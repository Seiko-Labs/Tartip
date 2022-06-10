import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

base_payment = 160000
actual_work_days = 24


with open('trip_kpi_for_driver.json', encoding='utf-8') as file:
    df_trip_kpi = json.load(file)
print(df_trip_kpi)
# df_trip_kpi = json.load(fp=r'C:\Users\Ayan\PycharmProjects\Tartyp_KPI\trip_kpi.json')
with open('consumption_kpi.json', encoding='utf-8') as file:
    df_consumption_kpi = json.load(file)
with open('RD_Driver.json', encoding='utf-8') as file:
    df_rd = json.load(file)
with open('weight_kpi_for_driver.json', encoding='utf-8') as file:
    df_weight_kpi = json.load(file)
with open('Tablet_index.json', encoding='utf-8') as file:
    df_tablet = json.load(file)

df_foto_fix = pd.read_json('fix_foto_result.json', orient='index', encoding='utf-8')
df_ttn = pd.read_json('fix_ttn_result.json', orient='index', encoding='utf-8')

driver_list = list(df_trip_kpi.keys())
driver_dict = {}
for driver in driver_list:
    truck_dict = {}
    try:
        functional_kpi = df_tablet[driver]*df_foto_fix.at[driver, 'index']*df_ttn.at[driver, 'index']
    except:
        print(f'Exception raised for {driver} while evaluation of functional_index')
        functional_kpi = 0
    for truck in list(df_trip_kpi[driver].keys()):
        trip_kpi = df_trip_kpi[driver][truck]
        consumption_kpi = df_consumption_kpi[truck]
        weight_kpi = df_weight_kpi[driver][truck]
        rd = df_rd[driver][truck]/actual_work_days
        print(f'This is indexes {functional_kpi}, {trip_kpi}, {consumption_kpi}, {weight_kpi}, {rd}')
        overall_kpi = base_payment * rd * (0.5 * trip_kpi + 0.15 * consumption_kpi + 0.15 * weight_kpi + 0.2 * functional_kpi)
        print(f'This is overall: {overall_kpi} for {driver} with {truck}')
        truck_dict[truck] = overall_kpi
    driver_dict[driver] = truck_dict
print(driver_dict)

with open ('Driver_KPI.json', 'w', encoding='utf-8') as out:
    json.dump(driver_dict, out, indent=4, ensure_ascii=False)
