import pandas as pd
import numpy as np
import json

path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Макрос АЗВ февраль оконч..xlsm'

df = pd.read_excel(path, sheet_name='NORMA')
df_2nd_row = pd.read_excel(path, sheet_name='NORMA', skiprows=1)

print(df.columns)
print(df_2nd_row.columns)
list_col = list(df.columns)
second_list_col = list(df_2nd_row)
new_col_list = []
for i in range(len(list_col)):
    if 'Unnamed' in list_col[i] and 'Unnamed' not in second_list_col[i]:
        print(list_col[i-1])
        w = list_col[i-1] + '_' + second_list_col[i]
        new_col_list.append(w)
    elif 'Unnamed' not in second_list_col[i]:
        w = list_col[i] + '_' + second_list_col[i]
        new_col_list.append(w)
    else:
        new_col_list.append(list_col[i])
print(new_col_list)

rename_dict = dict(zip(list_col, new_col_list))
print(rename_dict)

df = df.rename(columns=rename_dict)
print(df.columns)
df.drop(index=0, axis=0, inplace=True)
df.set_index('Гос.номер СТ', inplace=True)

df.to_json('Normas_parse.json', force_ascii=False, indent=4, orient='index')
