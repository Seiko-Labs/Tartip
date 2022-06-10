import pandas as pd
import numpy as np
import xlsxwriter
import json
import collections

norm_df = pd.read_json('Normas_parse.json')
print(norm_df)


trip_count_df = pd.read_json('gsm_result.json')
print(trip_count_df)
truck_number_list = list(trip_count_df.index)
days_list = list(trip_count_df.columns)
trips_per_truck_dict = {}
for truck in truck_number_list:
    trips = []
    w_days = []
    for day in days_list:
        val = trip_count_df.at[truck, day]
        if isinstance(val, list):
            trips.append(val[0])
            w_days.append(1)
        else:
            print('Not working day')
    trip_sum = sum(trips)
    days_sum = sum(w_days)
    trips_per_truck_dict[truck] = trip_sum/days_sum
print(trips_per_truck_dict)

kpi_trips = {}
for e in trips_per_truck_dict.keys():
    avg = trips_per_truck_dict[e]
    try:
        norm = norm_df.at['Норма рейсов', e]
        kpi_trips[e] = 2 ** (avg - norm)
    except:
        print(f'Exception raised on {e}')
        norm = avg
        kpi_trips[e] = 0

with open ('trip_kpi.json', 'w', encoding='utf-8') as out:
    json.dump(kpi_trips, out, indent=4)