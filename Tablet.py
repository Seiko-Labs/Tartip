import pandas as pd
import numpy as np
import xlsxwriter
import json

path = r'C:\Users\Ayan\Desktop\KPI_февраль\Работа с планшетами февраль.xlsx'

df = pd.read_excel(path, skiprows=5)
print(df)

new_dict = {}
for i in df.index:
    new_dict[df.at[i, 'Водитель']] = df.at[i, 'Кол-во нарядов с планшетом']/df.at[i, 'Общее кол-во нарядов']
print(new_dict)

with open ('Tablet_index.json', 'w', encoding='utf-8') as out:
    json.dump(new_dict, out, indent=4, ensure_ascii=False)