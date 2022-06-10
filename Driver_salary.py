import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections
from transliterate import translit

df_w_days = pd.read_json('w_days_result.json', encoding='utf-8')
df_name_parse = pd.read_json('Name_parse.json', encoding='utf-8', orient='index')
print(df_w_days)
print(df_name_parse)
col_list = list(df_w_days.columns)
truck_drivers_dict = {}
for i in col_list:
    temp_list = df_w_days[i].tolist()
    ll_list = [item for sublist in temp_list if sublist for item in sublist]
    ll_list = list(set(ll_list))
    for e in range(len(ll_list)):
        ll_list[e] = df_name_parse.at[ll_list[e], 'Full name']
    truck_drivers_dict[translit(i, language_code='ru', reversed=True)] = ll_list

with open('Teeest.json', 'w', encoding='utf-8') as out:
    json.dump(truck_drivers_dict, out, indent=4, ensure_ascii=False)


print(truck_drivers_dict)
print(truck_drivers_dict['287CF02'])
with open('Driver_KPI.json', encoding='utf-8') as file:
    df_drivers_kpi = json.load(file)
with open('Trucks_KPI.json', encoding='utf-8') as file:
    df_trucks_kpi = json.load(file)

print(df_drivers_kpi)
print(df_trucks_kpi)
for_each_driver = {}
for driver in df_drivers_kpi.keys():
    for_each_truck = {}
    for truck in df_drivers_kpi[driver].keys():
        truck = translit(truck, language_code='ru', reversed=True)
        driver_kpi = df_drivers_kpi[driver][truck]
        print(truck)
        ts_kpi = df_trucks_kpi[truck]
        drivers_list = truck_drivers_dict[truck]
        all_drivers_s_payment = []
        for dr in drivers_list:
            print(dr)
            single_payment = df_drivers_kpi[dr][truck]
            all_drivers_s_payment.append(single_payment)
        all_drivers_payment = sum(all_drivers_s_payment)
        for_each_truck[truck] = ts_kpi*(driver_kpi/all_drivers_payment)
    for_each_driver[driver] = for_each_truck

print(for_each_driver)

driver_salary = {}
for driver in for_each_driver.keys():
    salary = []
    for truck in for_each_driver[driver].keys():
        val = for_each_driver[driver][truck]
        salary.append(val)
    driver_salary[driver] = sum(salary)

with open('Driver_salary.json', 'w', encoding='utf-8') as out:
    json.dump(driver_salary, out, indent=4, ensure_ascii=False)