import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

weight_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Данные весовой АЗВ за февраль.xlsx'
df = pd.read_excel(weight_path)
df['Дата/Время'] = df['Дата/Время'].dt.strftime('%d.%m.%Y')
df_merged = df['Нетто'].groupby([df['Дата/Время'], df['№ транспорта']]).apply(list).reset_index()

print(df_merged)
my_dict = {k : g['Дата/Время'].tolist() for k, g in df_merged.groupby(df_merged['№ транспорта'])}

print(my_dict)

with open ('RD_Truck.json', 'w', encoding='utf-8') as out:
    json.dump(my_dict, out, indent=4, ensure_ascii=False)


# weight_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Данные весовой АЗВ за февраль.xlsx'
# df = pd.read_excel(weight_path)
# df['Дата/Время'] = df['Дата/Время'].dt.strftime('%d.%m.%Y')
# df_merged = df['Нетто'].groupby([df['Дата/Время'], df['Водитель']]).apply(list).reset_index()
#
# print(df_merged)
# my_dict = {k : g['Дата/Время'].tolist() for k, g in df_merged.groupby(df_merged['Водитель'])}
#
# print(my_dict)
#
# with open ('RD_Driver.json', 'w', encoding='utf-8') as out:
#     json.dump(my_dict, out, indent=4, ensure_ascii=False)