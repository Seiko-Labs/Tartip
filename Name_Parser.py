import pandas as pd
import numpy as np
import json

path = r'C:\Users\Ayan\Desktop\KPI_февраль\АЗВ_февраль\Name_parser.xlsx'

df = pd.read_excel(path)
print(df)
df = df.set_index('Name ')
with open('Name_parse.json', 'w', encoding='utf-8') as out:
    df.to_json('Name_parse.json', indent=4, orient='index', force_ascii=False)