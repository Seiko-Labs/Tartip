import math

import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections
from transliterate import translit

actual_w_days = 24
df_name_parse = pd.read_json('Name_parse.json', encoding='utf-8', orient='index')
truck_w_days_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Выход техники АЗВ февраль 2022 г.xlsx'
xlsx_file = pd.ExcelFile(truck_w_days_path).sheet_names
del xlsx_file[-6:]
print(xlsx_file)
df = pd.read_excel(truck_w_days_path, sheet_name=xlsx_file, skiprows=1)
with open ('Trucks_KPI.json', encoding='utf-8') as file:
    df_truck_kpi = json.load(file)

for i in xlsx_file:
    df[i].dropna(how='all', inplace=True)
    df[i]['Дата'] = i
print(df)
df_merged = pd.concat(df.values(), ignore_index=True)
df_merged = df_merged[df_merged['ФИО водителя '].notna()]
print(df_merged)
NUMB = df_merged['Гос.номер'].tolist()
NUMB = set(NUMB)
NUMB = list(NUMB)
nnn_dict= {}
for i in range(len(NUMB)):
    if 'KZ' or 'kz' in NUMB[i]:
        qq =NUMB[i].replace('KZ', '')
        qq = qq.replace('kz', '')
        qq = qq.replace(' ', '')
        print(qq[:1], qq[-2:])
        if qq[:1] == 'A' and qq[-2:] == '02':
            print('Success')
            qq = qq[1:]
        elif qq[:1] == 'А' and qq[-2:] == '02':
            print('Success')
            qq = qq[1:]
        print('This is qq: ', qq)
        nnn_dict[NUMB[i]] = qq
    else:
        qq = NUMB[i].replace(' ', '')
        print(qq[:1], qq[-2:])
        if qq[:1] == 'A' and qq[-2:] == '02':
            print('Success')
            qq = qq[1:]
        elif qq[:1] == 'А' and qq[-2:] == '02':
            print('Success')
            qq = qq[1:]
        print('This is qq: ', qq)
        nnn_dict[NUMB[i]] = qq

df_merged = df_merged.replace({'Гос.номер':nnn_dict})
# fio1 = df_merged['ФИО грузчика 1'].tolist()
# nname_list = []
# for i in range(len(fio1)):
#     try:
#         nname = df_name_parse.at[i, 'Full name']
#     except:
#         nname = i
#
for nammme in df_name_parse.index:
    try:
        df_merged['ФИО грузчика 1'].replace({nammme : df_name_parse.at[nammme, 'Full name']}, inplace=True)
        print(df_merged['ФИО грузчика 1'])
    except:
        pass
    try:
        df_merged['ФИО грузчика 2'].replace({nammme : df_name_parse.at[nammme, 'Full name']}, inplace=True)
    except:
        pass
    try:
        df_merged['ФИО грузчика 2'].replace({nammme : df_name_parse.at[nammme, 'Full name']}, inplace=True)
    except:
        pass

exception_list = ['Контейнер', 'Вывоз', 'Развозка', 'Загрузился', 'Разкозка',  str(math.nan)]

df_l1 = df_merged['Дата'].groupby([df_merged['ФИО грузчика 1'], df_merged['Гос.номер']]).apply(list).reset_index()
df_l2 = df_merged['Дата'].groupby([df_merged['ФИО грузчика 2'], df_merged['Гос.номер']]).apply(list).reset_index()
df_l3 = df_merged['Дата'].groupby([df_merged['ФИО грузчика 3'], df_merged['Гос.номер']]).apply(list).reset_index()

# print(df_l3['ФИО грузчика 3'].str.contains('Контейнер', na=False))
# df_l3 = df_l3.apply(lambda row: row[df_l3['ФИО грузчика 3'].isin(exception_list)])
# print(df_l3)
iter = 0
for i in exception_list:
    print('Iteration ', iter)
    df_l1 = df_l1[~df_l1['ФИО грузчика 1'].str.contains(i)]
    df_l2 = df_l2[~df_l2['ФИО грузчика 2'].str.contains(i)]
    df_l3 = df_l3[~df_l3['ФИО грузчика 3'].str.contains(i)]
    iter += 1
#     # except:
#     #     df_l1 = df_l1[df_l1['ФИО грузчика 1'] != i]
#     #     df_l2 = df_l2[df_l2['ФИО грузчика 2'] != i]
#     #     df_l3 = df_l3[df_l3['ФИО грузчика 3'] != i]
l1_list = df_l1['ФИО грузчика 1'].tolist()
l2_list = df_l2['ФИО грузчика 2'].tolist()
l3_list = df_l3['ФИО грузчика 3'].tolist()
print(f'This is l1_names {l1_list}')
print(f'This is l2_names {l2_list}')
print(f'This is l3_names {l3_list}')
names_list = []
names_list.append(l1_list)
names_list.append(l2_list)
names_list.append(l3_list)
print(f'This is appended l1_list {names_list}')
b_list = [item for sublist in names_list for item in sublist]
overall_list_names = list(set(b_list))
l1_list_trucks = df_l1['Гос.номер'].tolist()
l2_list_trucks = df_l2['Гос.номер'].tolist()
l3_list_trucks = df_l3['Гос.номер'].tolist()
trucks_list = []
trucks_list.append(l1_list_trucks)
trucks_list.append(l2_list_trucks)
trucks_list.append(l3_list_trucks)

c_list = [item for sublist in trucks_list for item in sublist]
overall_list_trucks = list(set(c_list))
print(f'This is overall_trucks_list {overall_list_trucks}')
# overall_list_trucks = df_merged['Гос.номер'].tolist()
# overall_list_trucks = list(set(overall_list_trucks))
print(overall_list_names)
my_dict_l1 = {k : g['Дата'].values[0] for k, g in df_l1.groupby([df_l1['ФИО грузчика 1'], df_l1['Гос.номер']])}
my_dict_l2 = {k : g['Дата'].values[0] for k, g in df_l2.groupby([df_l2['ФИО грузчика 2'], df_l2['Гос.номер']])}
my_dict_l3 = {k : g['Дата'].values[0] for k, g in df_l3.groupby([df_l3['ФИО грузчика 3'], df_l3['Гос.номер']])}
print('This is my_dict_1', my_dict_l1)
print(my_dict_l2)
print(my_dict_l3)

# with open('l1.json', 'w', encoding='utf-8') as out:
#     json.dump(my_dict_l1, out, indent=4, ensure_ascii=False)
# with open('l2.json', 'w', encoding='utf-8') as out:
#     json.dump(my_dict_l2, out, indent=4, ensure_ascii=False)
# with open('l3.json', 'w', encoding='utf-8') as out:
#     json.dump(my_dict_l3, out, indent=4, ensure_ascii=False)

new_dict_1 = {}
new_dict_2 = {}
new_dict_3 = {}
for i in my_dict_l1.keys():
    print(i)
    name = i[0]
    if name in new_dict_1:
        pass
    else:
        new_dict_1[name] = dict.fromkeys(overall_list_trucks)
    if i[0] in overall_list_names:
        truck = i[1]
        if i[0] == name:
            new_dict_1[name][truck] = my_dict_l1[i]
            testvar = my_dict_l1[i]
            testvar2 = new_dict_1[name]
            print()
        else:
            new_dict_1[name][truck] = None
            print('None value occurred')
    else:
        print(f'name {name} not in list')
        pass
print('This is new_dict1: ', new_dict_1)


for i in my_dict_l2.keys():
    name = i[0]
    if name in new_dict_2:
        pass
    else:
        new_dict_2[name] = dict.fromkeys(overall_list_trucks)
    if i[0] in overall_list_names:
        truck = i[1]
        if i[0] == name:
            new_dict_2[name][truck] = my_dict_l2[i]
        else:
            new_dict_2[name][truck] = None
            print('None value occurred')

    else:
        print(f'name {name} not in list')
        pass
print('This is new_dict2: ', new_dict_2)

for i in my_dict_l3.keys():
    name = i[0]
    if name in new_dict_3:
        pass
    else:
        new_dict_3[name] = dict.fromkeys(overall_list_trucks)
    if i[0] in overall_list_names:
        truck = i[1]
        if i[0] == name:
            new_dict_3[name][truck] = my_dict_l3[i]
        else:
            new_dict_3[name][truck] = None
    else:
        pass
print('This is new_dict3: ', new_dict_3)

loader_dict = {}

keys_list = [list(new_dict_1.keys()), list(new_dict_2.keys()), list(new_dict_3.keys())]
all_keys = [item for sublist in keys_list for item in sublist]
all_keys = list(set(all_keys))
print(f'This is allkeys {all_keys}')
new_dict = {}

# val = new_dict_1['Абдикаймов Ж']['609CM02']
# print('This is vaLL111111: ', val)


for name in all_keys:
    print(name)
    # try:
    #     new_name = df_name_parse.at[name, 'Full name']
    # except:
    #     new_name = name
    new_name = name
    sub_dict = {}
    for truck in overall_list_trucks:
       # print(f'The truck {truck} with name {new_name}')
        try:
            val1 = new_dict_1[name][truck]
        except:
            val1 = None
        try:
            val2 = new_dict_2[name][truck]
        except:
            val2 = None
        try:
            val3 = new_dict_3[name][truck]
        except:
            val3 = None
        ov_list = [val1, val2, val3]
        filled_list = [i for i in ov_list if i]
        flat_list = [item for sublist in filled_list for item in sublist]
        if not flat_list:
            pass
        else:
            sub_dict[truck] = flat_list
    print(f'Name {new_name} and {sub_dict}')
    a_keys_list = list(new_dict.keys())
    if new_name in new_dict:
        new_dict[new_name].update(sub_dict)
    else:
        new_dict[new_name] = sub_dict

print('This is new_dict: ', new_dict)


loader_dict = {}
for i in new_dict.keys():
    truck_dict = {}
    for j in new_dict[i].keys():
        if new_dict[i][j] == None:
            pass
        else:
            truck_dict[j] = new_dict[i][j]
    loader_dict[i] = truck_dict

print(loader_dict)

loader_salary_old = {}
loader_salary_new = {}
for loader in loader_dict.keys():
    old_f_dict= {}
    new_f_dict = {}
    print(loader)
    for truck in loader_dict[loader].keys():
        # truck = translit(truck, 'ru', reversed=True)
        truck_kpi = df_truck_kpi[truck]
        rd = len(loader_dict[loader][truck])/actual_w_days
        old_formula = truck_kpi/2*rd
        new_formula = truck_kpi*0.65
        old_f_dict[truck] = old_formula
        new_f_dict[truck] = new_formula
    loader_salary_old[loader] = old_f_dict
    loader_salary_new[loader] = new_f_dict

print(loader_salary_old)
print(loader_salary_new)

with open('Loaders_salary_old.json', 'w', encoding='utf-8') as out:
    json.dump(loader_salary_old, out, indent=4, ensure_ascii=False)
with open('Loaders_salary_new.json', 'w', encoding='utf-8') as out:
    json.dump(loader_salary_new, out, indent=4, ensure_ascii=False)