import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

base_payment = 160000
func_index = 1
actual_work_days = 24

df_w_days = pd.read_json('w_days_result.json')
with open('trip_kpi.json') as file:
    df_trip_kpi = json.load(file)
# df_trip_kpi = json.load(fp=r'C:\Users\Ayan\PycharmProjects\Tartyp_KPI\trip_kpi.json')
with open('consumption_kpi.json') as file:
    df_consumption_kpi = json.load(file)
with open('RD_Truck.json') as file:
    df_rd = json.load(file)
with open('weight_kpi.json') as file:
    df_weight_kpi = json.load(file)
# df_consumption_kpi = pd.read_json('consumption_kpi.json')
# df_rd= pd.read_json('RD_Truck.json')
# df_weight_kpi = pd.read_json('weight_kpi.json')

trucks_list = df_w_days.keys()
print(df_trip_kpi['609CM02'])
all_trucks_kpi = {}
for truck in trucks_list:
    try:
         trip_kpi = df_trip_kpi[truck]
         consumption_kpi = df_consumption_kpi[truck]
         weight_kpi = df_weight_kpi[truck]
         rd = len(df_rd[truck])/actual_work_days
         overall_kpi = base_payment * rd * (0.5*trip_kpi + 0.15*consumption_kpi + 0.15*weight_kpi + 0.2*func_index)
    except:
        overall_kpi = 0
        print(f'Exception Raised for {truck}! Check the truck numbers if they are correct')
    all_trucks_kpi[truck] = overall_kpi

with open('Trucks_KPI.json', 'w', encoding='utf-8') as out:
    json.dump(all_trucks_kpi, out, indent=4, ensure_ascii=False)