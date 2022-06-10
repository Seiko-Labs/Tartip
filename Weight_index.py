# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import xlsxwriter
import json

from googleapiclient.errors import HttpError

from Google import *

import os
from Google import Create_Service


# range_names = [
# 'Лист1!A1:H15'
# ]
# result = service.spreadsheets().values().get(
#     spreadsheetId='13GFy1Z0WveRw12KetuBmucRn8m0kxZDpHonjnSMd700', range=range_names).execute()
# ranges = result.get('valueRanges', [])
# # print('{0} ranges retrieved.'.format(len(ranges)))
# print(ranges)


a = batch_get_values("1UcaizxVvwXzCKTEUvV8SsnpwaNWf_TPuYfuB3Z7v-pM", 'ПЭС АЛМ ежедн. СВОД')
print(a.keys())
first = a['valueRanges'][0]['values'][5]
second = a['valueRanges'][0]['values'][6]
#df = pd.DataFrame(columns=a['valueRanges'][0]['values'][5])
print(len(first))
#print(len(second))
for i in range(len(first)):
    try:
        n = second[i]
    except:
        second.insert(i, '')


new_col_names = []
# arq2nd = list(df_dict_2nd_layer_col[i].columns)
# arq = list(df[i].columns)
# print(arq)
# print(arq2nd)
for e in range(len(first)):
    #print('Iteration: ', e)
    if first[e] != '' and second[e] != '':
        #print('1st option')
        new_col_names.append(first[e]+'_'+second[e])
    elif first[e] == '' and second[e] != '':
        #print('2nd option')
        first[e] = str(first[e-1])
        new_col_names.append(first[e]+'_'+second[e])
    elif second[e] == '' and first[e] != '':
        new_col_names.append(first[e])
        #print('3rd option')
    else:
        new_col_names.append(first[e-1]+'_'+second[e-1]+'.1')
        #print('Error occured!')
        #print(first[e], second[e])
new_col_names.append("ID")
print(new_col_names)
val_range = a['valueRanges'][0]['values'][9:]


#for i in range(len(val_range)):
    #del val_range[i][-3:]

for i in range(len(val_range)):
    for j in range(0, 48):
        try:
            n = val_range[i][j]
        except:
            val_range[i].insert(j, '')
    val_range[i].insert(49, i+1)
data = dict()
for i in range(len(new_col_names)):
    data[new_col_names[i]] = list(map(lambda v: v[i], val_range))
val_range = val_range[0:len(new_col_names)]
df = pd.DataFrame(val_range, columns=new_col_names)

zip(dict(a))
ex = a['valueRanges'][0]['values'][8:]
print(len(ex))
# df = pd.DataFrame(columns=a['valueRanges'][0]['values'][6])
df_col_list = df.columns

for e in range(len(ex)):
    len_of_ex = len(ex[e])
    for k in range (0, 48):
        df.at[e, df_col_list[k]] = ex[e][k]


exel = df.get(["ID", "Марка, модель ТС", "ГРНЗ", "Ф.И.О. водителя", "ДАТА", "Норма расхода ДТ, литр _на 100 км пробега", "Норма расхода ДТ, литр _на доп. работу оборудования", "Лимит на выдачу ДТ план, литр", "Неснижаемый остаток ДТ ", "Выход техники, 1/ 0", "Остаток ДТ на начало, литр", "Рекомендовано ДТ, литр", "Получено ДТ, литр_ВСЕГО:", "Пробег, км/ мото-час для грейф. Погрузчиков", "Общее кол-во рейсов", "в том числе:_на МСК", "в том числе:_Перевалка", "в том числе:_на ГП", "Общий объем ТБО, тонн", "Объем ТБО за рейс, тонн_1 рейс", "Объем ТБО за рейс, тонн_2 рейс", "Объем ТБО за рейс, тонн_3 рейс"])
exel_2 = df.get(["Объем ТБО за рейс, тонн_4 рейс", "Объем ТБО за рейс, тонн_5 рейс", "Расход ДТ, литр_ВСЕГО:", "Расход ДТ, литр_в том числе:", "Расход ДТ, литр_в том числе:.1", "Остаток ДТ на конец дня (расчетный), литр", "Остаток ДТ для пополнения", "Объем ДТ для заправки, литр"])
new_df = pd.concat([exel, exel_2]) #в один дф все столбцы не записались, записала в 2 разных и объединила в один

print(new_df)
df_new = new_df.get(['ID', 'ГРНЗ']).groupby([new_df['ДАТА']]).apply(list).reset_index()

#df_new.set_index('ДАТА', inplace=True)
#print ("Тест")
#print(df_new)

#new_df.set_index('ID', inplace=True)
result_ex = df_new.to_json(orient="index")
parsed_ex = json.loads(result_ex)
json_result_ex = json.dumps(parsed_ex, ensure_ascii=False, indent=4)
with open('weight_result_exel.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex)


# df = pd.DataFrame(columns=a['valueRanges'][0]['values'][7])
# list_col = list(df.columns)
# spec_dict = {}
# for i in range(len(list_col)):
#     try:
#         val = a['valueRanges'][0]['values'][8][i]
#     except:
#         val = ''
#     spec_dict[list_col[i]] = val
# # new_dict = dict(zip(list(df.columns), a['valueRanges'][0]['values'][8]))
# print(spec_dict)

weight_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Данные весовой АЗВ за февраль.xlsx'
df = pd.read_excel(weight_path)
df['Дата/Время'] = df['Дата/Время'].dt.strftime('%d.%m.%Y')
df_merged = df['Нетто'].groupby([df['Дата/Время'], df['№ транспорта']]).apply(list).reset_index()

print(df_merged)
my_dict = {k : g['№ транспорта'].tolist() for k, g in df_merged.groupby(df_merged['Дата/Время'])}

print(my_dict)

numb = len(my_dict)
days_list = list(range(1, numb+1))
i = 0
df_dict = {}
final_df = pd.DataFrame()

for key in my_dict:
    print('iter number ', key)
    name = 'temp_df_' + str(days_list[i])[8:10]
    key1 = str(key)
    df_dict[name] = pd.DataFrame(columns=['fio', key1])
    i += 1
    val = my_dict[key]
    print(len(val))

    for e in range(len(val)):
        # arc = df1.query('`Начало разговора (дата)`==@key and `Код оператора`==@my_dict[@key][@e]')['Оценка звонка']
        arc = df_merged[(df_merged['Дата/Время'] == key) & (df_merged['№ транспорта'] == my_dict[key][e])]
        v = arc['Нетто'].values[0]
        c = len(v)
        pos = v.count(1)
        f_v = [pos, c]
        df_dict[name] = df_dict[name].append({'fio': my_dict[key][e], str(key): v}, ignore_index=True)

    print(df_dict[name])

    # final_df = pd.merge()
    print('Final')

    if final_df.empty:
        final_df = df_dict[name]
    else:
        final_df = pd.merge(final_df, df_dict[name], how='outer', on='fio')

print(final_df)


final_df.set_index('fio', inplace=True)
result = final_df.to_json(orient='index')
parsed = json.loads(result)
json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
with open('weight_result.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result)

# weight_norm_df = pd.read_excel(weight_path, sheet_name='Лист1')
#
# weight_norm_df.set_index('ГРНЗ', inplace=True)
# weight_norm_df.to_json('weight_norm.json', indent=4, orient='index', force_ascii=False)

norm_df = pd.read_json('Normas_parse.json', orient='index')

df_new = pd.read_json('weight_result.json', orient='index')
days_list = list(df_new.columns)
w_kpi_dict = {}
for i in df_new.index:
    try:
        w_norm = norm_df.at[i, 'Норма грузоподьемности _зимнее']
    except:
        print(i)
        w_norm = 0
        print('Exception occurred')
    trip_count = []
    weight_count = []
    for day in days_list:
        try:
            trips_per_day = len(df_new.at[i, day])
        except:
            trips_per_day = 0

        try:
            weight_per_day = sum(df_new.at[i, day])
        except:
            weight_per_day = 0
        trip_count.append(trips_per_day)
        weight_count.append(weight_per_day)
    avg_weight = sum(weight_count)/sum(trip_count)
    if avg_weight >= w_norm:
        kpi_weight = 1
    else:
        kpi_weight = ((-1/(w_norm*w_norm))*(avg_weight*avg_weight)) + ((2/w_norm)*avg_weight)
    w_kpi_dict[i] = kpi_weight

print(w_kpi_dict)

with open('weight_kpi.json', 'w', encoding='utf_8') as outfile:
    json.dump(w_kpi_dict, outfile, indent=4)