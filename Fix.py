import pandas as pd
import numpy as np
import xlsxwriter
import json

fix_path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\фиксация АВЗ февраль .xlsx'
df = pd.read_excel(fix_path, skiprows=4)
df = df.dropna()

df = df['Процент'].groupby([df['Водитель']]).apply(float).reset_index()
df.set_index('Водитель', inplace=True)

df.rename(columns={'Процент':'index'}, inplace=True)
df['index'] = df['index'].div(100)
print(df)

result = df.to_json(orient='index')
parsed = json.loads(result)
json_result = json.dumps(parsed, ensure_ascii=False, indent=4)
with open('fix_ttn_result.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result)