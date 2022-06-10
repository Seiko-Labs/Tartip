import datetime

import pandas as pd
import numpy as np
import xlsxwriter
import json
import math


k = 3

norm_df = pd.read_json('Normas_parse.json')

gsm_df = pd.read_json('gsm_result.json')
print(gsm_df)
days_list = list(gsm_df.columns)

avg_cons_per_truck = {}
for truck in gsm_df.index:
    consumption_list = []
    w_days_list = []
    for day in days_list:
        try:
            consumption_val = gsm_df.at[truck, day]
            consumption_val = consumption_val[4]
            w_days_list.append(1)
        except:
            consumption_val = 0
        consumption_list.append(consumption_val)
    total_consumption = sum(consumption_list)
    w_days = sum(w_days_list)
    avg_consumption = total_consumption/w_days
    print(f'This is avg consumption for {truck}', avg_consumption)
    avg_cons_per_truck[truck] = avg_consumption


kpi_consumption = {}
for e in avg_cons_per_truck.keys():
    val = avg_cons_per_truck[e]
    try:
        norm = norm_df.at['Нормы ГСМ_зимнее.1', e]
        kpi = math.sqrt((norm/k)+1-(val/k))
        print('kpi exists')
    except:
        kpi = 0
    kpi_consumption[e] = kpi

with open ('consumption_kpi.json', 'w', encoding='utf-8') as out:
    json.dump(kpi_consumption, out, indent=4)