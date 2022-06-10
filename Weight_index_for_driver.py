import math

import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

def OccurrenceCount(a):
    k = {}
    for j in a:
        if j in k:
            k[j] += 1
        else:
            k[j] = 1
    return k


weight_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Данные весовой АЗВ за февраль.xlsx'
df = pd.read_excel(weight_path)
df['Дата/Время'] = df['Дата/Время'].dt.strftime('%d.%m.%Y')
df_merged = df['№ транспорта'].groupby([df['Дата/Время'], df['Водитель']]).apply(list).reset_index()

print(df_merged)
my_dict = {k : g['Дата/Время'].tolist() for k, g in df_merged.groupby(df_merged['Водитель'])}

print(my_dict)
new_dict = {}
for name in my_dict.keys():
    days_list = my_dict[name]
    day_truck_dict = {}
    for day in days_list:
        val = df_merged[(df_merged['Дата/Время'] == day) & (df_merged['Водитель'] == name)]['№ транспорта'].values[0]
        print(val)
        day_truck_dict[day] = val
    new_dict[name] = day_truck_dict

print(new_dict)

copy_dict = new_dict
print('This is copy_dict: ', copy_dict)
for name in new_dict.keys():
    for day in new_dict[name].keys():
        val = new_dict[name][day]
        new_val = OccurrenceCount(val)
        new_dict[name][day] = new_val
print(new_dict)

final_dict = {}
check_dict = {}
for name in new_dict.keys():
    dict_temp = {}
    w_days = {}
    for day in new_dict[name].keys():
        truck_list = new_dict[name][day].keys()
        for key in truck_list:
            temp_val = new_dict[name][day][key]
            if key in dict_temp:
                dict_temp[key] = dict_temp[key] + temp_val
            else:
                dict_temp[key] = temp_val

    final_dict[name] = dict_temp
    # check_dict[name] = w_days
print(final_dict)

nn_dict = {}

for name in new_dict.keys():
    w_days = []
    nn_dict[name] = []
    for day in new_dict[name].keys():
        truck_list = list(new_dict[name][day].keys())
        w_days.append(truck_list)
    nn_dict[name].extend(w_days)

for name in nn_dict:
    big_list = nn_dict[name]
    flat_list = [item for sublist in big_list for item in sublist]
    counter_val = OccurrenceCount(flat_list)
    nn_dict[name] = counter_val
print(nn_dict)

with open ('RD_Driver.json', 'w', encoding='utf-8') as out:
    json.dump(nn_dict, out, indent=4, ensure_ascii=False)


trip_index_for_driver_dict = {}
for name in final_dict.keys():
    trucks = final_dict[name].keys()
    trip_index_per_truck = {}
    for truck in trucks:
        trip_index_for_driver = final_dict[name][truck]/nn_dict[name][truck]
        trip_index_per_truck[truck] = trip_index_for_driver
    trip_index_for_driver_dict[name] = trip_index_per_truck

print(trip_index_for_driver_dict)

norm_df = pd.read_json('Normas_parse.json')

kpi_dict = {}
for name in trip_index_for_driver_dict.keys():
    trucks = list(trip_index_for_driver_dict[name].keys())
    s_dict = {}
    for truck in trucks:
        f_val = 2**(trip_index_for_driver_dict[name][truck]-norm_df.at['Норма рейсов', truck])
        s_dict[truck] = f_val
    kpi_dict[name] = s_dict
print(kpi_dict)

with open ('trip_kpi_for_driver.json', 'w', encoding='utf-8') as out:
    json.dump(kpi_dict, out, indent=4, ensure_ascii=False)




weight_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Данные весовой АЗВ за февраль.xlsx'
df = pd.read_excel(weight_path)
df['Дата/Время'] = df['Дата/Время'].dt.strftime('%d.%m.%Y')
df_merged = df['Нетто'].groupby([df['Дата/Время'], df['Водитель'], df['№ транспорта']]).apply(list).reset_index()

a = df_merged.groupby(df_merged['Водитель'])
print(a)
my_dict = {k : g['Нетто'].tolist() for k, g in df_merged.groupby([df_merged['Водитель'], df_merged['№ транспорта']])}
ll = list(my_dict.keys())
print(type(ll[0]))
print(list(my_dict.keys()))

ww_dict = {}
tt_dict = {}
name_list = []
truck_n_list = []
for i in range(len(ll)):
    name_list.append(ll[i][0])
tt_dict = dict.fromkeys(name_list)
for name in name_list:
    tt_dict[name] = {}
    for i in ll:
        # print(i)
        # print(name)
        if i[0] == name:
            print(f'correct name!')
            truck_n = i[1]
            # name = ll[i][0]
            # truck_num = ll[i][1]
            # n_key = name + '_' + truck_num
            t_list = my_dict[i]
            # print(f'This is t_list {t_list} with key {i}')
            val_list = [item for sublist in t_list for item in sublist]
            val_list = [x for x in val_list if not (math.isnan(x) == True)]
            tt_dict[name][truck_n] = val_list
        else:
            print('no name!')
            pass

# print(list(set(name_list)))
    # if ll[i][0] != ll[i-1][0]:
    #     zz_dict = {}
    #     n_name = ll[i][0]
    #     t_list = my_dict[ll[i]]
    #     val_list = [item for sublist in t_list for item in sublist]
    #     zz_dict[ll[i][1]] = val_list
    # else:
    #     zz_dict = {}
    #     t_list = my_dict[ll[i]]
    #     val_list = [item for sublist in t_list for item in sublist]
    #     zz_dict[ll[i][1]] = val_list


print(tt_dict)
kpi_dict = {}

for name in tt_dict.keys():
    truck_dict = {}
    for truck in tt_dict[name].keys():
        val_list = tt_dict[name][truck]
        avg_weight = sum(val_list)/len(val_list)

        print(sum(val_list), len(val_list))
        print(f'This is avg weight: {avg_weight} for {name} with {truck}')
        norm = norm_df.at['Норма грузоподьемности _зимнее', truck]
        # if avg_weight >= norm:
        #     kpi = 1
        # else:
        kpi = (-1/(norm**2)*(avg_weight**2)) + (2/norm*avg_weight)
        print(f'This is kpi: {kpi}')
        truck_dict[truck] = kpi
    kpi_dict[name] = truck_dict


with open ('weight_kpi_for_driver.json', 'w', encoding='utf-8') as out:
    json.dump(kpi_dict, out, indent=4, ensure_ascii=False)
# print(check_dict)
# arc = df_merged[(df_merged['Дата/Время'] == key) & (df_merged['№ транспорта'] == my_dict[key][e])]


# numb = len(my_dict)
# days_list = list(range(1, numb+1))
# i = 0
# df_dict = {}
# final_df = pd.DataFrame()
#
# for key in my_dict:
#     print('iter number ', key)
#     name = 'temp_df_' + str(days_list[i])[8:10]
#     key1 = str(key)
#     df_dict[name] = pd.DataFrame(columns=['fio', key1])
#     i += 1
#     val = my_dict[key]
#     print(len(val))
#
#     for e in range(len(val)):
#         # arc = df1.query('`Начало разговора (дата)`==@key and `Код оператора`==@my_dict[@key][@e]')['Оценка звонка']
#         arc = df_merged[(df_merged['Дата/Время'] == key) & (df_merged['№ транспорта'] == my_dict[key][e])]
#         v = arc['Нетто'].values[0]
#         c = len(v)
#         pos = v.count(1)
#         f_v = [pos, c]
#         df_dict[name] = df_dict[name].append({'fio': my_dict[key][e], str(key): v}, ignore_index=True)
#
#     print(df_dict[name])
#
#     # final_df = pd.merge()
#     print('Final')
#
#     if final_df.empty:
#         final_df = df_dict[name]
#     else:
#         final_df = pd.merge(final_df, df_dict[name], how='outer', on='fio')
#
# print(final_df)
#
# final_df.set_index('fio', inplace=True)
# result = final_df.to_json(orient='index')
# parsed = json.loads(result)
# json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
# with open('weight_result.json', "w", encoding='utf-8') as outfile:
#     outfile.write(json_result)
#
# # weight_norm_df = pd.read_excel(weight_path, sheet_name='Лист1')
# #
# # weight_norm_df.set_index('ГРНЗ', inplace=True)
# # weight_norm_df.to_json('weight_norm.json', indent=4, orient='index', force_ascii=False)
#
# norm_df = pd.read_json('Normas_parse.json', orient='index')
#
# df_new = pd.read_json('weight_result.json', orient='index')
# days_list = list(df_new.columns)
# w_kpi_dict = {}
# for i in df_new.index:
#     try:
#         w_norm = norm_df.at[i, 'Норма грузоподьемности _зимнее']
#     except:
#         print(i)
#         w_norm = 0
#         print('Exception occurred')
#     trip_count = []
#     weight_count = []
#     for day in days_list:
#         try:
#             trips_per_day = len(df_new.at[i, day])
#         except:
#             trips_per_day = 0
#
#         try:
#             weight_per_day = sum(df_new.at[i, day])
#         except:
#             weight_per_day = 0
#         trip_count.append(trips_per_day)
#         weight_count.append(weight_per_day)
#     avg_weight = sum(weight_count)/sum(trip_count)
#     if avg_weight >= w_norm:
#         kpi_weight = 1
#     else:
#         kpi_weight = ((-1/(w_norm*w_norm))*(avg_weight*avg_weight)) + ((2/w_norm)*avg_weight)
#     w_kpi_dict[i] = kpi_weight
#
# print(w_kpi_dict)
#
# with open('weight_kpi.json', 'w', encoding='utf_8') as outfile:
#     json.dump(w_kpi_dict, outfile, indent=4)