import pandas as pd
import numpy as np
import xlsxwriter
import json

fix_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\фотофиксация АВЗ февраль.xlsx'
df = pd.read_excel(fix_path, skiprows=6)

df = df['Процент покрытия'].groupby([df['Водитель']]).apply(float).reset_index()
df.set_index('Водитель', inplace=True)
print(df)
df.rename(columns={'Процент покрытия':'index'}, inplace=True)
df['index'] = df['index'].div(100)


result = df.to_json(orient='index')
parsed = json.loads(result)
json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
with open('fix_foto_result.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result)