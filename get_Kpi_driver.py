import pandas as pd
import numpy as np
import xlsxwriter
import json
from googleapiclient.errors import HttpError
import os
from Google import Create_Service
import codecs
from Google import *

a = batch_get_values("1UcaizxVvwXzCKTEUvV8SsnpwaNWf_TPuYfuB3Z7v-pM", 'ПЭС АЛМ ежедн. СВОД')
print(a.keys())

#считывание всех уровней заголовок столбцов
first = a['valueRanges'][0]['values'][5]
second = a['valueRanges'][0]['values'][6]
third = a['valueRanges'][0]['values'][7]
for i in range(len(first)):
    try:
        n = second[i]
    except:
        second.insert(i, '')
    try:
        m = third[i]
    except:
        third.insert(i, '')

new_col_names = []
#парсинг заголовок столбцов
for e in range(len(first)):
    if first[e] != '' and second[e] != '':
        new_col_names.append(first[e]+'_'+second[e]+third[e])
    elif first[e] == '' and second[e] != '':
        first[e] = str(first[e-1])
        new_col_names.append(first[e]+'_'+second[e]+third[e])
    elif second[e] == '' and first[e] != '':
        new_col_names.append(first[e])
    else:
        if first[e] == '' and second[e] == '':
            new_col_names.append(first[e - 2] + '_' + second[e - 2] + '_' + third[e])
        else:
            new_col_names.append(first[e-1]+'_'+second[e-1]+'_'+third[e])

new_col_names.append("ID")
val_range = a['valueRanges'][0]['values'][9:]

for i in range(len(val_range)):
    for j in range(0, 48):
        try:
            n = val_range[i][j]
        except:
            val_range[i].insert(j, '')
    val_range[i].insert(48, i+1)

new_val_range=[] #новый массив без пустых строк
for i in range(len(val_range)):
    if val_range[i][0]!='' and val_range[i][2]!='' and val_range[i][8]!='':
        new_val_range.append(val_range[i])

data = dict()
for i in range(len(new_col_names)):
    data[new_col_names[i]] = list(map(lambda v: v[i], new_val_range))
print(new_col_names)
df = pd.DataFrame(data=data)
df.set_index('ID', inplace=True)
df_new = df['Ф.И.О. водителя'].groupby([df['ДАТА']]).apply(list)

result_ex = df_new.to_json(orient = 'index')
parsed_ex = json.loads(result_ex)
json_result_ex = json.dumps(parsed_ex, ensure_ascii=False, indent=4)

with open('weight_result_driver_exel.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex)
    
with open('weight_result_driver_exel.json') as f:
    file_content = f.read()
    GRNS_by_data = json.loads(file_content)


Days_all = list(GRNS_by_data.keys())
GRNS_all = list(GRNS_by_data.items())
Cp_GSM_dict = {}

for i in range(len(GRNS_all)): #счетчик дней
    Cp_GSM_day = {}
    for j in range(len(GRNS_all[i][1])): #счетчик транспорта
        GRNS_by_day = list(filter(lambda x: GRNS_all[i][1][j] in x and Days_all[i] in x, new_val_range))
        if (GRNS_by_day[0][30] != "0,0") and (GRNS_by_day[0][30] != '') and (GRNS_by_day[0][43] != ''):
            rashod = GRNS_by_day[0][43].replace(",", ".")
            probeg = GRNS_by_day[0][30].replace(",", ".")
            Cp_GSM = (float(rashod) / float(probeg)) * 100
            Cp_GSM_day [GRNS_all[i][1][j]] = Cp_GSM, GRNS_by_day[0][19] #значение Ср_ГСМ и норма расхода ГСМ
        else:
            continue
    Cp_GSM_dict [Days_all[i]] = Cp_GSM_day

#Вытаскиваем все ГРНЗ номера
df_GRNS = df['ДАТА'].groupby([df['ГРНЗ']]).apply(list)
GRNS = list(df_GRNS.keys())

Cp_GSM_month = {}
Cp_GSM_sum = ''
for i in range (len(GRNS)):
    Cp_GSM_sum = 0;
    days_count = 0; #счетчик рабочих дней
    for j in range(len(Days_all)):
        day = Days_all[j]
        GRNS_number = GRNS[i]
        a = Cp_GSM_dict[day].get(GRNS_number, False) #проверка на наличие ключа в словаре
        if Cp_GSM_dict[day].get(GRNS_number, False):
            Cp_GSM_one_day = Cp_GSM_dict[day][GRNS_number][0]
            Cp_GSM_sum = Cp_GSM_sum + float(Cp_GSM_one_day) #складываем Ср_ГСМ за все дни
            days_count = days_count +1
            Norma = Cp_GSM_dict[day][GRNS_number][1]
        else:
            continue
    if days_count > 0:
        Cp_GSM_sum = Cp_GSM_sum / days_count
    else:
        Cp_GSM_sum = 0
    Cp_GSM_month[GRNS[i]] = Cp_GSM_sum, Norma

print ("Ср ГСМ за месяц по ГРНЗ:")
print(Cp_GSM_month)

k = 3
KPI_dict = {}
for i in range (len(GRNS)):
    GRNS_number = GRNS[i]
    Cp_GSM = float(Cp_GSM_month[GRNS_number][0])
    Ngsm = Cp_GSM_month[GRNS_number][1].replace(",", ".")
    Ngsm = float(Ngsm)
    if (Cp_GSM-Ngsm)>3:
        KPI_dict[GRNS_number] = 0
    else:
        KPI = (Ngsm/k) + 1 - (1/k*Cp_GSM)
        KPI_final = KPI ** (0.5)
        KPI_dict [GRNS_number] = KPI_final

print ("Итоговый KPI транспорта:")
print (KPI_dict)




