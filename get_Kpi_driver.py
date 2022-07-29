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

# Считывание всех уровней заголовок столбцов___________________________________________________________________________
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

# Парсинг заголовок столбцов___________________________________________________________________________________________
new_col_names = []
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

new_val_range = []# новый массив без пустых строк
for i in range(len(val_range)):
    if val_range[i][0]!='' and val_range[i][2]!='' and val_range[i][8]!=''and val_range[i][9]!='':
        new_val_range.append(val_range[i])

data = dict()
for i in range(len(new_col_names)):
    data[new_col_names[i]] = list(map(lambda v: v[i], new_val_range))
df = pd.DataFrame(data=data)
df.set_index('ID', inplace=True)

# _____________________________________________________________________________________________________________________
# Сортировка датафрейма для грузчиков__________________________________________________________________________________
df_loader1 = df['ГРНЗ'].groupby([df['Ф.И.О. грузчик-1']]).apply(list)
result_ex_1 = df_loader1.to_json(orient = 'index')
parsed_ex_1 = json.loads(result_ex_1)
json_result_ex_1 = json.dumps(parsed_ex_1, ensure_ascii=False, indent=4)

df_loader2 = df['ГРНЗ'].groupby([df['Ф.И.О. грузчик-2']]).apply(list)
result_ex_2 = df_loader2.to_json(orient = 'index')
parsed_ex_2 = json.loads(result_ex_2)
json_result_ex_2 = json.dumps(parsed_ex_2, ensure_ascii=False, indent=4)

df_loader3 = df['ГРНЗ'].groupby([df['Ф.И.О. грузчик-3']]).apply(list)
result_ex_3 = df_loader3.to_json(orient = 'index')
parsed_ex_3 = json.loads(result_ex_3)
json_result_ex_3 = json.dumps(parsed_ex_3, ensure_ascii=False, indent=4)

df_loader4 = df['ГРНЗ'].groupby([df['Ф.И.О. грузчик-4']]).apply(list)
result_ex_4 = df_loader4.to_json(orient = 'index')
parsed_ex_4 = json.loads(result_ex_4)
json_result_ex_4 = json.dumps(parsed_ex_4, ensure_ascii=False, indent=4)

with open('Loader/loader1.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex_1)
with open('Loader/loader2.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex_2)
with open('Loader/loader3.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex_3)
with open('Loader/loader4.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex_4)

fileObj_1 = codecs.open( "Loader/loader1.json", "r", "utf_8_sig" )
file_content_1 = fileObj_1.read()  # или читайте по строке
Loader_by_GRNS_1 = json.loads(file_content_1)

fileObj_2 = codecs.open( "Loader/loader2.json", "r", "utf_8_sig" )
file_content_2 = fileObj_2.read()  # или читайте по строке
Loader_by_GRNS_2 = json.loads(file_content_2)

fileObj_3 = codecs.open( "Loader/loader3.json", "r", "utf_8_sig" )
file_content_3 = fileObj_3.read()  # или читайте по строке
Loader_by_GRNS_3 = json.loads(file_content_3)

fileObj_4 = codecs.open( "Loader/loader4.json", "r", "utf_8_sig" )
file_content_4 = fileObj_4.read()  # или читайте по строке
Loader_by_GRNS_4 = json.loads(file_content_4)

# Получение всех ФИО грузчиков_________________________________________________________________________________________
Loader_1 = list(Loader_by_GRNS_1.keys())
Loader_2 = list(Loader_by_GRNS_2.keys())
Loader_3 = list(Loader_by_GRNS_3.keys())
Loader_4 = list(Loader_by_GRNS_4.keys())

Loader_all = [*Loader_1, *Loader_2, *Loader_3, *Loader_4]
Loader_FIOs = list(set(Loader_all))
Loader_Grns = {}
for i in range(len(Loader_FIOs)):
    FIO = Loader_FIOs[i]
    GRNS_by_day = list(filter(lambda x: FIO in x, new_val_range))
    try:
        District = GRNS_by_day[0][0]
    except:
        District = None
    GRNS = []
    GRNS_list_1 = []
    GRNS_list_2 = []
    GRNS_list_3 = []
    GRNS_list_4 = []
    if Loader_by_GRNS_1.get(FIO, False):
        GRNS_list_1 = list(Loader_by_GRNS_1[FIO])
        GRNS.append(list(set(GRNS_list_1)))
    if Loader_by_GRNS_2.get(FIO, False):
        GRNS_list_2 = list(Loader_by_GRNS_2[FIO])
        GRNS.append(list(set(GRNS_list_2)))
    if Loader_by_GRNS_3.get(FIO, False):
        GRNS_list_3 = list(Loader_by_GRNS_3[FIO])
        GRNS.append(list(set(GRNS_list_3)))
    if Loader_by_GRNS_4.get(FIO, False):
        GRNS_list_4 = list(Loader_by_GRNS_4[FIO])
        GRNS.append(list(set(GRNS_list_4)))
    GRNS_list = [*GRNS_list_1, *GRNS_list_2, *GRNS_list_3, *GRNS_list_4]
    GRNS_all_list = list(set(GRNS_list))
    Loader_Grns[FIO] = {"job": "Грузчик", "District": District, "GRNS_list": GRNS_all_list, "Work_days": len(GRNS_by_day)}
print(Loader_Grns)

with open("Json/Loaders_GRNS.json", "w", encoding="utf-8") as file:
    json.dump(Loader_Grns, file, sort_keys=False, indent=4, ensure_ascii=False)

# _____________________________________________________________________________________________________________________
# Сортировка датафрейма для водителей__________________________________________________________________________________
df_new = df['ГРНЗ'].groupby([df['Ф.И.О. водителя']]).apply(list)

result_ex = df_new.to_json(orient = 'index')
parsed_ex = json.loads(result_ex)
json_result_ex = json.dumps(parsed_ex, ensure_ascii=False, indent=4)

with open('weight_result_driver_exel.json', "w", encoding='utf-8') as outfile:
    outfile.write(json_result_ex)

fileObj = codecs.open( "weight_result_driver_exel.json", "r", "utf_8_sig" )
file_content = fileObj.read()
GRNS_by_data = json.loads(file_content)

# Получение всех ФИО водителей_________________________________________________________________________________________
FIO_all = list(GRNS_by_data.keys())

with open('Json/Premiya_auto.json') as f:
    file_content = f.read()
    GRNS_all = json.loads(file_content)
GRNS = list(GRNS_all.keys())

# Рассчет Kpi_gsm для водителей________________________________________________________________________________________
Cp_GSM_dict = {}
for i in range(len(FIO_all)):
    Cp_GSM_day = {}
    work_days = 0 # счетчик рабочих дней
    for k in range(len(GRNS)):
        list_GRNS = []
        GRNS_by_day = list(filter(lambda x: FIO_all[i] in x and GRNS[k] in x, new_val_range))
        MSK = 0
        all_trip = 0
        tonns = 0
        Cp_GSM = 0
        days_count = 0  # счетчик рабочих дней на одной машине
        District = None
        if len(GRNS_by_day) > 0:
            for j in range (0, len(GRNS_by_day)):
                District = GRNS_by_day[j][0]
                if (GRNS_by_day[j][30] != "0,0") and (GRNS_by_day[j][30] != '') and (GRNS_by_day[j][43] != ''):
                    rashod = GRNS_by_day[j][43].replace(",", ".")
                    probeg = GRNS_by_day[j][30].replace(",", ".")
                    tonn = GRNS_by_day[j][36].replace(",", ".")
                    MSK = MSK + int(GRNS_by_day[j][32])
                    all_trip = all_trip + int(GRNS_by_day[j][31])
                    Cp_GSM = Cp_GSM + ((float(rashod) / float(probeg)) * 100)
                    tonns = tonns + float(tonn)
                    days_count = days_count + 1
                    work_days = work_days + 1
                else:
                    Cp_GSM = 0
                    MSK = 0
                    tonns = 0

            Cp_GSM_day[GRNS[k]] = {'District': District, 'Cp_GSM': Cp_GSM, 'Norm_100': GRNS_by_day[0][19], 'MSK': MSK,
                                   'All_trips': all_trip, 'Weight': tonns, 'Work_days': days_count, 'All_days': work_days}
        list_GRNS.append(Cp_GSM_day)
    Cp_GSM_dict[FIO_all[i]] = list_GRNS


Cp_GSM_month = {}
for k in range (len(FIO_all)):
    FIOs_number = FIO_all[k]
    list_GRNS_GSM = []
    for i in range(len(Cp_GSM_dict[FIOs_number][0])):
        GRNS_list = Cp_GSM_dict[FIOs_number][0]
        Trips_count = 0  # счетчик всех рейсов
        Weight = 0
        GRNS_GSM = {}
        GRNS = list(GRNS_list.keys())
        GRNS = GRNS[i]
        Cp_GSM_sum = Cp_GSM_dict[FIOs_number][0][GRNS]['Cp_GSM']
        MSK_count = Cp_GSM_dict[FIOs_number][0][GRNS]['MSK']
        Weight = Cp_GSM_dict[FIOs_number][0][GRNS]['Weight']
        work_days = int(Cp_GSM_dict[FIOs_number][0][GRNS]['Work_days'])
        all_days = int(Cp_GSM_dict[FIOs_number][0][GRNS]['All_days'])
        Trips_count = int(Cp_GSM_dict[FIOs_number][0][GRNS]['All_trips'])

        if all_days > 0:
            Cp_GSM_sum = Cp_GSM_sum / all_days
            RD = MSK_count / all_days
            Cpp = Trips_count / all_days
        else:
            Cp_GSM_sum = 0
            RD = 0
            Cpp = 0

        if Trips_count > 0:
            Cpt = Weight / Trips_count
        else:
            Cpt = 0

        GRNS_GSM [GRNS] = {'Cp_GSM_sum': Cp_GSM_sum, 'RD': RD, 'Cpp': Cpp, 'Cpt': Cpt}
        list_GRNS_GSM.append(GRNS_GSM)
    Cp_GSM_month[FIOs_number] = list_GRNS_GSM

k = 3
KPI_GSM_dict = {}
for j in range (len(FIO_all)):
    FIOs_number = FIO_all[j]
    Kpi_list = []
    for i in range(len(Cp_GSM_month[FIOs_number])):
        GRNS_list = Cp_GSM_dict[FIOs_number][0]
        GRNS = list(GRNS_list.keys())
        GRNS = GRNS[i]
        KPI_GSM_dict_one = {}
        RD = Cp_GSM_month[FIOs_number][i][GRNS]['RD']
        Cpp = Cp_GSM_month[FIOs_number][i][GRNS]['Cpp']
        Cp_GSM = Cp_GSM_month[FIOs_number][i][GRNS]['Cp_GSM_sum']
        Cp_GSM = float(Cp_GSM)
        Ngsm = Cp_GSM_dict[FIOs_number][0][GRNS]['Norm_100'].replace(",", ".")
        Ngsm = float(Ngsm)
        if (Cp_GSM-Ngsm) > 3 and (Cp_GSM == 0):
            KPI_GSM_dict_one[GRNS] = {'Kpi_gsm': 0, 'RD': RD, 'Cpp': Cpp}
        else:
            KPI = (Ngsm/k) + 1 - (1/k*Cp_GSM)
            KPI_final = KPI ** (0.5)
            KPI_GSM_dict_one[GRNS] = {'Kpi_gsm': KPI_final, 'RD': RD, 'Cpp': Cpp}
            #print(KPI_GSM_dict_one)
        Kpi_list.append(KPI_GSM_dict_one)
    KPI_GSM_dict[FIOs_number] = Kpi_list

with open("Json/Kpi_GSM_driver.json", "w", encoding="utf-8") as file:
    json.dump(KPI_GSM_dict, file, sort_keys=False, indent=4, ensure_ascii=False)

# Получение норм транспорта____________________________________________________________________________________________
with open('Norms/Norms_auto.json') as f:
    file_content = f.read()
    Norms = json.loads(file_content)

# Расчет Kpi_рейсов____________________________________________________________________________________________________
KPI_trips_dict={}
for j in range (len(FIO_all)):
    FIOs_number = FIO_all[j]
    Kpi_list = []
    for i in range(len(Cp_GSM_month[FIOs_number])):
        KPI_one_trips_dict={}
        GRNS_list = Cp_GSM_dict[FIOs_number][0]
        GRNS = list(GRNS_list.keys())
        GRNS = GRNS[i]
        Cpp = Cp_GSM_month[FIOs_number][i][GRNS]['Cpp']
        if Norms.get(GRNS, False):  # проверка на наличие ключа в словаре
            Nt = Norms[GRNS]['norm_trips']
        else:
            Nt = 0
        Kpi = 2 ** (Cpp - Nt)
        KPI_one_trips_dict[GRNS] = Kpi
        Kpi_list.append(KPI_one_trips_dict)
    KPI_trips_dict[FIOs_number] = Kpi_list

# Рассчет Kpi-Тоннаж___________________________________________________________________________________________________
Kpi_tonn_dict = {}
for j in range (len(FIO_all)):
    FIOs_number = FIO_all[j]
    Kpi_list = []
    for i in range(len(Cp_GSM_month[FIOs_number])):
        KPI_one_tonns_dict = {}
        GRNS_list = Cp_GSM_dict[FIOs_number][0]
        GRNS = list(GRNS_list.keys())
        GRNS = GRNS[i]
        if Norms.get(GRNS, False):  # проверка на наличие ключа в словаре
            now = datetime.datetime.now()
            if (now.month > 3) and (now.month > 10):
                Nt = Norms[GRNS]['norm_weight_summer']
            else:
                Nt = Norms[GRNS]['norm_weight_winter']
        else:
            NT = 0
        Cpt = Cp_GSM_month[FIOs_number][i][GRNS]['Cpt'] * 1000
        if Cpt > Nt:
            Kpi = 1
        else:
            Kpi = ((-1/Nt ** 2) * (Cpt ** 2)) + (2 * Cpt/Nt)
        Kpi = ((-1 / Nt ** 2) * Cpt ** 2) + (2 * Cpt / Nt)
        KPI_one_tonns_dict[GRNS] = Kpi
        Kpi_list.append(KPI_one_tonns_dict)
    Kpi_tonn_dict[FIOs_number] = Kpi_list

fileObj = codecs.open("Json/Kpi_tablet.json", "r", "utf_8_sig")
text = fileObj.read()
Tablet = eval(text.replace('null', 'None'))

# Рассчет Pst__________________________________________________________________________________________________________
Pst_dict = {}
for j in range (len(FIO_all)):
    FIOs_number = FIO_all[j]
    Pst_list = []
    for i in range(len(Cp_GSM_month[FIOs_number])):
        Pst_one_dict = {}
        GRNS_list = Cp_GSM_dict[FIOs_number][0]
        GRNS = list(GRNS_list.keys())
        GRNS = GRNS[i]
        S = 160000
        if Tablet.get(FIOs_number, False):  # проверка на наличие ключа в словаре
            Kpi_tablet = Tablet[FIOs_number]['KPI']
            Kpi_tablet = Tablet[FIOs_number]['KPI']
        else:
            Kpi_tablet = 0
        RD = KPI_GSM_dict[FIOs_number][i][GRNS]['RD']
        Kpi_trip = KPI_trips_dict[FIOs_number][i][GRNS]
        Kpi_tonn = Kpi_tonn_dict[FIOs_number][i][GRNS]
        Kpi_GSM = KPI_GSM_dict[FIOs_number][i][GRNS]['Kpi_gsm']
        All_days = int(Cp_GSM_dict[FIOs_number][0][GRNS]['All_days'])
        Pst = S * RD * (0.5 * Kpi_trip + 0.15 * Kpi_tonn + 0.15 * Kpi_GSM + 0.2 * Kpi_tablet)
        District = Cp_GSM_dict[FIOs_number][0][GRNS]['District']
        Pst_one_dict[GRNS] = Pst
        Pst_list.append(Pst_one_dict)
    Pst_dict[FIOs_number] = {"job": "Водитель", "District": District, "Pst_list": Pst_list}

print(Pst_dict)
with open("Json/Premiya_drivers.json", "w", encoding="utf-8") as file:
    json.dump(Pst_dict, file, sort_keys=False, indent=4, ensure_ascii=False)
