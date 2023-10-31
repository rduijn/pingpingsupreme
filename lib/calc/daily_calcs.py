import pandas as pd
import os
from pathlib import Path
import numpy as np
import openpyxl
from datetime import datetime, timedelta
from dateutil.relativedelta import *

def get_project_root() -> Path:
    return Path(__file__).parent.parent


PATH = 'C:\\Users\\rduij\\downloads'

import os
FILENAMES = [filename for filename in os.listdir(PATH) if filename.startswith("NL61INGB0442453558_")]

transactions_dfs = [pd.read_csv(PATH + "/" + filename, sep=';', decimal=',') for filename in FILENAMES]

df = pd.concat(transactions_dfs, axis=0)
print(len(df))

df = df.drop_duplicates()
df = df[df['Af Bij'] == 'Af']

print(len(df))
df['CATEGORY'] = ''
df_filter_base = pd.read_excel(str(get_project_root().parent) + '\\lib\\tables\\MAPPINGTABLE.xlsx', sheet_name='BASE')

for _, filter_base in df_filter_base.iterrows():
    temp = filter_base.DATE
    #ToDo - instead of if else -> if cell has value, add filter
    if not pd.isnull(filter_base.DATE):
        date_ing_format = temp.year*10000 + temp.month*100 + temp.day
        df.loc[(df['Naam / Omschrijving'] == filter_base.NAME) & (df['Datum'] == date_ing_format), 'CATEGORY'] = filter_base.CATEGORY
    elif not isinstance(filter_base.ACCOUNTTO, str) and not isinstance(filter_base.NOTE, str):
        try:
            df.loc[df['Naam / Omschrijving'].str.contains(filter_base.NAME), 'CATEGORY'] = filter_base.CATEGORY
        except:
            pass
    elif not isinstance(filter_base.ACCOUNTTO, str) and isinstance(filter_base.NOTE, str):
        df.loc[(df['Naam / Omschrijving'].str.contains(filter_base.NAME)) & (df['Mededelingen'].str.contains(filter_base.NOTE)), 'CATEGORY'] = filter_base.CATEGORY
    else:
        df.loc[(df['Naam / Omschrijving'] == filter_base.NAME) & (df['Tegenrekening'] == filter_base.ACCOUNTTO), 'CATEGORY'] = filter_base.CATEGORY

cf_mapped = sum(df['Bedrag (EUR)'][df.CATEGORY != ''])
cf_not_mapped = sum(df['Bedrag (EUR)'][df.CATEGORY == ''])

print(f"--- MAPPING STATS ({min(df.Datum)}-{max(df.Datum)})---")
print(f"MAPPED (%): {cf_mapped / (cf_mapped + cf_not_mapped) * 100:20.2f}")
print(f"MAPPED (EUR):     {cf_mapped:14,.2f}")
print(f"NOT MAPPED (EUR): {cf_not_mapped:14,.2f}")
print(f"\n--- COST PER TYPE ---")
df_mapped = df[df.CATEGORY != ''].sort_values('Bedrag (EUR)', ascending=False)
pivot = df_mapped.groupby(['CATEGORY'])['Bedrag (EUR)'].sum().sort_values(ascending=False)
print(pivot.head(50))


print(f"\n--- COST PER TYPE ---")
print("\n\n")
df_not_mapped = df[df.CATEGORY == ''].sort_values('Naam / Omschrijving', ascending=False)
pivot = df_not_mapped.groupby(['Naam / Omschrijving'])['Bedrag (EUR)'].sum().sort_values(ascending=False)
print(pivot.head(50))


df.loc[df.CATEGORY == '', 'CATEGORY'] = 'OTHER'
df = df[df.CATEGORY != 'NOCOST']
df['datetime'] = [datetime(int(str(x)[:4]),int(str(x)[4:6]),int(str(x)[6:8])) for x in df.Datum]

def prev_weekday(adate):
    adate -= timedelta(days=1)
    while adate.weekday() > 4: # Mon-Fri are 0-4
        adate -= timedelta(days=1)
    return adate



yyyy= 2022
d = 20

lb_date = datetime(yyyy, 6, d) +relativedelta(months=-1)
ub_date = datetime(yyyy, 6, d)
for period in range(1, 18):

    lb_date_adj = prev_weekday(lb_date)
    ub_date_adj = prev_weekday(ub_date)
    print(f"{period} {lb_date_adj.strftime('%Y-%m-%d')} {ub_date_adj.strftime('%Y-%m-%d')} >>> {lb_date_adj.year*100 + lb_date_adj.month} ")
    df.loc[(df['datetime']>=lb_date_adj) & (df['datetime']<ub_date_adj),'PERIOD'] = str(lb_date_adj.year*10000 + lb_date_adj.month * 100 + lb_date_adj.day)+"-" + str(ub_date_adj.year*10000 + ub_date_adj.month * 100 + ub_date_adj.day-1)
    lb_date = lb_date + relativedelta(months=1)
    ub_date = ub_date + relativedelta(months=1)

final_pivot = pd.pivot_table(df, values='Bedrag (EUR)', index=['CATEGORY'],
                       columns=['PERIOD'], aggfunc=np.sum)

final_pivot['sum_cols'] = final_pivot.sum(axis=1)
final_pivot = final_pivot.sort_values('sum_cols' , ascending=False)
final_pivot.to_clipboard()
temp = 3
df.to_clipboard()
temp=4
