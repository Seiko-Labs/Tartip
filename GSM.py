import datetime

import pandas as pd
import numpy as np
import xlsxwriter
import json

gsm_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\02_2022_АЗВ_ГСМ.xlsx'
xlsx_file = pd.ExcelFile(gsm_path).sheet_names
xlsx_file.pop(0)
print(xlsx_file)
df = pd.read_excel(gsm_path, sheet_name=xlsx_file, skiprows=4)
df_dict_2nd_layer_col = pd.read_excel(gsm_path, sheet_name=xlsx_file, skiprows=5)
total_by_day = {}
for i in xlsx_file:
    new_col_names = []
    arq = list(df[i].columns)
    arq2nd = list(df_dict_2nd_layer_col[i].columns)
    print(arq)
    print(arq2nd)
    for e in range(len(arq)):
        print('Iteration: ', e)
        if 'Unnamed' not in arq[e] and 'Unnamed' not in arq2nd[e]:
            print('1st option')
            new_col_names.append(arq[e]+'_'+arq2nd[e])
        elif 'Unnamed' in arq[e] and 'Unnamed' not in arq2nd[e]:
            print('2nd option')
            arq[e] = str(arq[e-1])
            new_col_names.append(arq[e]+'_'+arq2nd[e])
        elif 'Unnamed' in arq2nd[e] and 'Unnamed' not in arq[e]:
            new_col_names.append(arq[e])
            print('3rd option')
        else:
            new_col_names.append(arq[e-1]+'_'+arq2nd[e-1]+'.1')
            print('Error occured!')
            print(arq[e], arq2nd[e])

    print(new_col_names)
    # print(arq)
    # n_c = dict(zip(arq, new_col_names))
    # print('This is n_c', n_c)
    # print(len(n_c))
    df[i].set_axis(new_col_names, axis=1, inplace=True)
    # df[i].rename(columns=n_c, inplace=True)
    print(df[i].columns)
    df[i].dropna(how='all', inplace=True)
    if i == '13' or i == '14':
        df[i] = df[i].iloc[2:, :]
    else:
        df[i] = df[i].iloc[3:, :]
    total_by_day[i] = df[i].iloc[-1:, :].to_dict(orient='index')
    # df[i].drop(df[i][df[i]['№'] == 'ИТОГО'], inplace=True)

print(total_by_day)
df_merged = pd.concat(df.values(), ignore_index=True)
df_merged = df_merged.drop(df_merged[df_merged['№'] == 'ИТОГО:'].index)
print('This is merged: ', df_merged)


df_merged = df_merged[(df_merged['Выход техники, 1/ 0'] == 1)]
# df_merged['Дата'] = df_merged['Дата'].dt.stftime('%Y-%m-%d')
df_merged.info()
df_new = df_merged['ГРНЗ'].groupby([df_merged['Дата']]).apply(list).reset_index()
df_new.set_index('Дата', inplace=True)
print(df_new)

result = df_new.to_json(orient='index', date_format='%d.%m.%Y')
parsed = json.loads(result)
json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
with open('gsm_result.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result)

df_check = pd.read_json('gsm_result.json', orient='index')
print(df_check)
# df_check['Дата'] = [datetime.datetime.strptime(date[0:10], '%Y-%m-%d').date() for date in df_check['Дата']]
# w = df_check.at['140CG02', 'Дата'][1]
# q = [datetime.datetime.strptime(date[0:10], '%Y-%m-%d').date() for date in df_check.at['140CG02', 'Дата']]

new_dict = {}
for n in df_check.index:
    val_dict = {}
    truck_num_list = df_check.at[n, 'ГРНЗ']
    for num in truck_num_list:
        val = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Общее кол-во рейсов'].values[0]
        to_start = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Остаток ДТ на начало, литр'].values[0]
        given = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Получено ДТ, литр'].values[0]
        to_end = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Остаток ДТ на конец дня (расчетный), литр'].values[0]
        consumption = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Расход ДТ, литр_ВСЕГО:'].values[0]
        standard_cons = df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Норма расхода ДТ, литр _на доп. работу оборудования'].values[0] + df_merged[(df_merged['ГРНЗ'] == num) & (df_merged['Дата'] == n)]['Норма расхода ДТ, литр _на 100 км пробега'].values[0]
        val_dict[num] = [val, given, to_start, to_end, consumption, standard_cons]
    new_dict[n] = val_dict

print(new_dict)
temp_df = pd.DataFrame.from_dict(new_dict, orient='index')
print(temp_df)
temp_df.to_json('gsm_result.json', force_ascii=False, orient='index', indent=4)

df_alfa = pd.read_json('gsm_result.json', encoding='utf-8', orient='index')
print(df_alfa)
# for i in df_check.index:
#     temp = df_check.at[i, 'Дата']
#     for j in range(len(temp)):
#         try:
#             temp[j] = pd.to_datetime(temp[j], unit='ms').strftime('%Y-%m-%d')
#         except:
#             print('Error occured')
#             continue
#
# df_check.to_json('new_gsm_res.json', orient='index', force_ascii=False, indent=4)
# print(df_check)