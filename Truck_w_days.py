import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

truck_w_days_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Выход техники АЗВ февраль 2022 г.xlsx'
xlsx_file = pd.ExcelFile(truck_w_days_path).sheet_names
del xlsx_file[-6:]
print(xlsx_file)
df = pd.read_excel(truck_w_days_path, sheet_name=xlsx_file, skiprows=1)

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
df_new = df_merged['ФИО водителя '].groupby([df_merged['Дата'], df_merged['Гос.номер']]).apply(list).reset_index()


print(df_new)

my_dict = {k : g['Гос.номер'].tolist() for k, g in df_new.groupby(df_new['Дата'])}


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
        arc = df_new[(df_new['Дата'] == key) & (df_new['Гос.номер'] == my_dict[key][e])]
        v = arc['ФИО водителя '].values[0]
        c = len(v)
        # pos = v.count(1)
        # f_v = [pos, c]
        df_dict[name] = df_dict[name].append({'fio': my_dict[key][e], str(key): v}, ignore_index=True)

    print(df_dict[name])

    # final_df = pd.merge()
    print('Final')

    if final_df.empty:
        final_df = df_dict[name]
    else:
        final_df = pd.merge(final_df, df_dict[name], how='outer', on='fio')

print(final_df)
# final_df = final_df.groupby([final_df['fio']]).apply(list).reset_index()
final_df.set_index('fio', inplace=True)

# print(final_df)
# new_dict= {}
# zz = list(final_df.index)
# print(final_df.index)
#
#
# final_df.rename(index=new_dict, inplace=True)
# ae = list(final_df.index)
#
# print([item for item, count in collections.Counter(ae).items() if count > 1])
# if ae[0] == ae[-5]:
#     print('TRUE')
# else:
#     print('FALSE')
# bb = len(ae)
# cc = len(set(ae))
# print(final_df.index.is_unique)
# print(bb, cc)
# print('This is ae: ', ae)
print(final_df)
final_df.to_json('w_days_result.json', orient='index', force_ascii=False, indent=4)
# parsed = json.loads(result)
# json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
# with open('w_days_result.json', "w", encoding='utf-8') as outfile:
#     outfile.write(json_result)
# df_dupl = df_new.loc[df_new.duplicated(subset='Марка спец. техники'), :]
