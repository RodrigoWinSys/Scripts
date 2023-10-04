# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 15:31:21 2023

@author: r.deapellaniz
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Apr  4 10:22:04 2023

@author: r.deapellaniz
"""


import pandas as pd
import json
import datetime as dt
import os
from ftplib import FTP
import ftplib
import numpy as np
import sys
import time
import datetime as dt

#files_directory=r'\\10.1.10.26\StatisticsDatabase'
files_directory=r'../csv acumulados'
script_dir=os.getcwd()
os.chdir(files_directory)
list_of_files=os.listdir()
list_of_files.sort(key=os.path.getctime)
print (list_of_files)
#latest_file_name = list_of_files[-1]
latest_file_name='Statistics_02_05_2023.csv'
previous_file_name='Statistics_02_04_2023.csv'
column_names=['Slot','themeId','denom','RTP','coinIn','coinOut','jackpotHandpay','ProgressiveCoinOut','gamesPlayed',
              'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']
print ('Latest file: ',latest_file_name,' Previous File: ',previous_file_name)
os.chdir(files_directory)
latest_file=pd.read_csv(latest_file_name,names=column_names,sep=';')
previous_file=pd.read_csv(previous_file_name,names=column_names,sep=';')


column_names=['Slot','themeId','denom','RTP','coinIn','coinOut','jackpotHandpay','ProgressiveCoinOut','gamesPlayed',
              'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']

# Convierto todo a numerico
cols=[i for i in latest_file.columns if i not in ['Slot','themeId','RTP','denom']]
for col in cols:
    latest_file[col]=pd.to_numeric(latest_file[col])
    
for col in cols:
    previous_file[col]=pd.to_numeric(previous_file[col])
    
previous_file=previous_file[previous_file['themeId'].notnull()]
latest_file=latest_file[latest_file['themeId'].notnull()]
"""
out=latest_file.merge(previous_file.drop_duplicates(),how='left',indicator=True,on=['Slot','denom','themeId','RTP'],suffixes=(None,"_y"))
out=out[out['_merge']=='left_only']
out=out[column_names]

for col in out.columns:
    if (col in cols):
        out[col].values[:] = 0
  
previous_file=previous_file.append(out,ignore_index=True)
"""
# Resta entre DFs
latest_file=latest_file.fillna(0)
previous_file=previous_file.fillna(0)
previous_file=previous_file[previous_file['themeId']!=0]


df_difference=latest_file.set_index(['Slot','themeId','RTP','denom'])[cols].subtract(previous_file.set_index(['Slot','themeId','RTP','denom'])[cols])

# Corrijo los ram clears
"""

indexes=df_difference[df_difference['coinIn']<0].index
latest_file.set_index(['Slot','themeId','RTP','denom'],inplace=True)

#latest_file.reset_index(inplace=True)

df_difference.loc[indexes]=latest_file.loc[indexes]
"""

df_difference.reset_index(inplace=True)
df_difference['Last Date']=latest_file_name.split('_')[1]+'-'+latest_file_name.split('_')[2]+'-'+latest_file_name.split('_')[3].split('.')[0]
df_difference['First Date']=previous_file_name.split('_')[1]+'-'+previous_file_name.split('_')[2]+'-'+previous_file_name.split('_')[3].split('.')[0]
df_difference['Last Date']=pd.to_datetime(df_difference['Last Date'],dayfirst=True)
df_difference['Last Date']=df_difference['Last Date'].dt.date
df_difference['First Date']=pd.to_datetime(df_difference['First Date'],dayfirst=True)
df_difference['First Date']=df_difference['First Date'].dt.date
df_difference['Date']=pd.to_datetime(df_difference['Last Date'])
df_difference['Day']=df_difference['Date'].dt.day
df_difference['month']=df_difference['Date'].dt.month
df_difference['year']=df_difference['Date'].dt.year

os.chdir(script_dir)
os.chdir('..\..\Inventario')
inventario=pd.read_excel(r'Inventario Winsystems Online.xlsx',sheet_name="Inventario total")
inventario=inventario[inventario['Online']==1]


df_difference=df_difference.merge(inventario,how='left',left_on='Slot',right_on='Match online')
#df_difference.to_excel(r'C:\Users\r.deapellaniz\OneDrive - Win Systems Group\Escritorio\Python scrpits\maquinas_online\Puerto Rico\Acumulated checks.xlsx',index=False)

os.chdir(script_dir)
os.chdir('..')
game_detail=pd.read_excel('Games Detail EGC 2022.xlsx')

df_difference=df_difference.merge(game_detail,how='left',left_on='themeId',right_on='Game')
#os.chdir('../2 Inventarios')

os.chdir(script_dir)
os.chdir('../')
full_db=pd.read_excel(r'games_report.xlsx')
#full_db=pd.DataFrame()
# Modificar la línea para que no elimine a lo pelotudo
#full_db=full_db[full_db['First Date'].dt.strftime('%Y%m')!=dt.datetime.today().strftime('%Y%m')]

df_difference['First Date']=pd.to_datetime(df_difference['First Date'],dayfirst=True)
df_difference['Last Date']=pd.to_datetime(df_difference['Last Date'],dayfirst=True)


full_db=full_db.append(df_difference,ignore_index=True)



full_db=full_db[['Grupo','Casino','Slot ID','Match online','Serie','Cabinet',
                 'Modelo','Mix','RTP TH','Hold TH','Jackpot','Tipo JP','Area',
                 'Isla','Posicion','Pais','Estado','Ciudad','Muncipio','TG SP',
                 'TG CI','TG NW','Slot_x','time_x','Estatus','Modo',
                 '% Participacion','Online','ON','Positions Approached',
                 'Positions Denied','Positions Confirmed',
                 'Infraestructure Needed','Infrastructure Exist','Scheduled',
                 'Positions Conected','Coments','# Isla','Connected',
                 'Programed','Pending','Date Connected','Date Programmed',
                 'Programmed Week','Conected Week','Dated','Instalada',
                 'Tipo Movimiento','Detalle Movimiento','Fecha Movimiento',
                 'Origen','Promotoria','Detalle Prom','Nombre Prom',
                 'Fecha de inicio de Op','Slot','coinIn','coinOut',
                 'jackpotHandpay','ProgressiveCoinOut','gamesPlayed',
                 'GameBaseCoinOut','ScatterWon','FreeGamePlays',
                 'FreeGamesCoinOut','BounsCointOut','Day','month','year','Date',
                 'Fecha alta','Owner','Razon Social','Serial',
                 'themeId','RTP','denom','Last Date','First Date',
                 'Platform','Game','Release','Lines','Vol','V','Math','HF',
                 'Base RTP','FF','FG F','FG RTP','BG F','BG RTP','MxPrz',
                 'Game Cycle','Max Liability','Side Bet (FB)','Cluster','RTP TH']]

full_db.to_excel(r'games_report.xlsx',index=False)
