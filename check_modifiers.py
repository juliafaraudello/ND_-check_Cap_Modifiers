# Let's import the libraries needed.
from xlrd import open_workbook
import xlsxwriter
from datetime import date
import pandas as pd
import numpy as np
import os
import time
import warnings
from pandas.core.frame import DataFrame
from datetime import timedelta, datetime
from optparse import OptionParser
import glob
import requests
import win32com.client as win32


#Forecaster logs root
root_forecaster_pull = r'C:\Users\zfaraude\Desktop\Develop\ND_Check_CapsModifiers\Forecaster Pull\*.csv'
root_station_config = r'C:\Users\zfaraude\Desktop\Develop\ND_Check_CapsModifiers\station_config.csv'

def chime_tabulation(df,table_title,text):
    df = df.fillna("-")
    columns = [ x for x in df.columns]

    df = df.rename(columns={columns[0]: f'|{columns[0]}'})
    df = df.add_suffix('|')

    for x,y in zip(df.columns, range(len(df.columns))):

        if y < int(len(df.columns)) - 1:
            df[x] = "|" + df[x].astype(str)
        else:
            df[x] = "|" + df[x].astype(str) + "|"

    start = "|---|"
    rest = "---|"

    markdown_sep = [start if x == 0 else rest for x in range(len(df.columns))]

    df.loc[-1] = markdown_sep
    df.index = df.index + 1
    df = df.sort_index()

    output = df.to_string(index=False).strip()

    # output_final = output.replace(" ","")

    last_output = '''/md


# {}

{}


{}
'''.format(table_title,text,output)
    requests.post("https://hooks.chime.aws/incomingwebhooks/858e6fec-99d8-4352-9584-384124c7eb19?token=UUd6b0prY0t8MXxTWjdKNC1vN2dPYmRkcFFtWEl3V2JQYXNpVHNkdWZzdWFLMTM1OU9mSE5V",
                        json={'Content': f'{last_output}'})
    #print(f'{last_output}')
    return last_output

#Select the last daily log file modification and create a df
def last_updated(root_forecaster_pull):
    list_of_files = glob.glob(root_forecaster_pull)
    list_of_files.sort(key=os.path.getctime,reverse=False)
    print('\n'.join(list_of_files))  
    file_name = list_of_files[-1]    # Last file from the folder
    print("lastFile:")
    print(file_name)

    status= os.stat(file_name)
    date = time.localtime(status.st_mtime)
    date = datetime(date[0], date[1], date[2], date[3], date[4], date[5])
    print ("The last modification of the file log is at: ", date)

    df_cap_modifiers=pd.read_csv(file_name,header=0)
    return (df_cap_modifiers, date)

def check_caps(df_merge,day_cap, date_last_mod):
    df_D1=df_merge[(df_merge.adjustment_factor != 1) & (df_merge.day_adjustment == day_cap)]
    df_country_count=df_D1.groupby(['country']).size().reset_index(name='N_DS')
    DS_list=df_D1.groupby(['country']).apply(lambda x: pd.Series({'DS': x['delivery_station'].tolist()}))
    df_report=pd.merge(df_country_count,DS_list, on='country')
    if (df_report['N_DS'].sum() > 267):
            print("Exceeds the maximum number of nodes to display in Chime, please check the file for D+"+ str(day_cap))
    if (df_report['N_DS'].sum() != 0):
        for index,row in df_report.iterrows():
            if row['N_DS'] >=3:
                print("DS for D+" + str(day_cap) + " have Active capping factors. There are more than three factors active in any Country, @PresentMembers will be notified.")
                table_title = ('Active capping factors - D+' + str(day_cap))
                text = ('@Present - More than three factors active in any Country - Last updated: ' + str(date_last_mod))
                chime_tabulation(df_report,table_title,text)
                print(df_report)
            else:
                print("DS for D+" + str(day_cap) + " have Active capping factors.")
                table_title = ('Active capping factors - D+' + str(day_cap))
                text = ('Last updated: ' + str(date_last_mod))
                chime_tabulation(df_report,table_title,text)
                print(df_report)
    else:
        print("No active capping factors for D+" + str(day_cap))
        payload = {'Content': "No active capping factors for D+" + str(day_cap) + ' - Last updated: ' + str(date_last_mod)}
        requests.post("https://hooks.chime.aws/incomingwebhooks/858e6fec-99d8-4352-9584-384124c7eb19?token=UUd6b0prY0t8MXxTWjdKNC1vN2dPYmRkcFFtWEl3V2JQYXNpVHNkdWZzdWFLMTM1OU9mSE5V",
                        json= payload)
    return df_report

def check_D3(df_merge, date_last_mod):
    actual_date = date.today()   
    if actual_date.weekday() == 5:
        df_D3=df_merge[(df_merge.adjustment_factor != 1) & (df_merge.day_adjustment == 3) & (df_merge.country == 'DE')]
        df_country_count=df_D3.groupby(['country']).size().reset_index(name='N_DS')
        DS_list=df_D3.groupby(['country']).apply(lambda x: pd.Series({'DS': x['delivery_station'].tolist()}))
        df_report=pd.merge(df_country_count,DS_list, on='country')
        if (df_report['N_DS'].sum() > 267):
            print("Exceeds the maximum number of nodes to display in Chime, please check the file for D+3")
        if (df_report['N_DS'].sum() != 0):
            if (df_report['N_DS'].sum() >=3):
                table_title = ('MEU: Active capping factors - D+3')
                text = ('@Present - More than three factors active in any MEU - Last updated: ' + str(date_last_mod)) 
                chime_tabulation(df_report,table_title,text)
                print(df_report)
            else:
                table_title = ('MEU: Active capping factors - D+3')
                text = ('Last updated: ' + str(date_last_mod)) 
                chime_tabulation(df_report,table_title,text)
                print(df_report)
        else:
            print("MEU: No active capping factors for D+3")
            payload = {'Content': "MEU: No active capping factors for D+3 - Last updated: " + str(date_last_mod)}
            requests.post("https://hooks.chime.aws/incomingwebhooks/858e6fec-99d8-4352-9584-384124c7eb19?token=UUd6b0prY0t8MXxTWjdKNC1vN2dPYmRkcFFtWEl3V2JQYXNpVHNkdWZzdWFLMTM1OU9mSE5V",
                        json= payload)
    else:
        print('DE: only on Saturdays D+3 is analyzed')
    return 


df_cap, date_last_mod = last_updated(root_forecaster_pull)

df_station_config=pd.read_csv(root_station_config,header=0)
df_merge = pd.merge(df_cap, df_station_config, left_on='delivery_station', right_on='Station')
df_merge= df_merge[['delivery_station', 'adjustment_factor','day_adjustment','country']]

check_caps(df_merge,1, date_last_mod)
check_caps(df_merge,2, date_last_mod)
check_D3(df_merge, date_last_mod)

