# -*- coding: utf-8 -*-

"""

Created on Tue Aug  3 16:03:58 2021
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

"""
Para cambiar el directorio donde se guardan los csv acumulados cambiar la línea 22.

"""
#files_directory=r'\\10.1.10.26\StatisticsDatabase'

files_directory=r'../csv test'

#bi_directory=r'C:\Users\r.deapellaniz\OneDrive - Win Systems Group\Escritorio\4Online'

#inventario_directory=r'C:\Users\r.deapellaniz\OneDrive - Win Systems Group\PBI Data Bases\Inventario'

script_dir=os.getcwd()

os.chdir(files_directory)

"""

Lee los úlitmos 2 csv que fueron agregados a la carpeta, en orden en que fueron agregados
y le pone el nombre a las columnas

"""

os.chdir(files_directory)
list_of_files=os.listdir()
list_of_files.sort(key=os.path.getctime)
print (list_of_files)
latest_file_name = list_of_files[-1]
previous_file_name =list_of_files[-2]

#latest_file_name = 'Statistics_05_04_2023.csv'

#previous_file_name ='Statistics_04_04_2023.csv'
print (latest_file_name, previous_file_name)

column_names=['Slot','themeId','denom','RTP','coinIn','coinOut','jackpotHandpay',

              'ProgressiveCoinOut','gamesPlayed','GameBaseCoinOut',

              'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']

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

# Divido por 100 las columnas que están en currency    

currency=['coinIn','coinOut','jackpotHandpay','ProgressiveCoinOut','GameBaseCoinOut', 'FreeGamesCoinOut', 'BounsCointOut']

for col in currency:
    latest_file[col]=latest_file[col]/100
    previous_file[col]=previous_file[col]/100

previous_file=previous_file[previous_file['themeId'].notnull()]

latest_file=latest_file[latest_file['themeId'].notnull()]

 

previous_file=pd.pivot_table(previous_file,index='Slot',values=['coinIn','coinOut','jackpotHandpay',

                                                                 'ProgressiveCoinOut','gamesPlayed','GameBaseCoinOut',

                                                                 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 

                                                                 'BounsCointOut'],aggfunc=np.sum)

previous_file.reset_index(inplace=True)

latest_file=pd.pivot_table(latest_file,index='Slot',values=['coinIn','coinOut','jackpotHandpay',

                                                                 'ProgressiveCoinOut','gamesPlayed','GameBaseCoinOut',

                                                                 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 

                                                                 'BounsCointOut'],aggfunc=np.sum)

latest_file.reset_index(inplace=True)  

out=latest_file.merge(previous_file.drop_duplicates(),how='left',indicator=True,on=['Slot'],suffixes=(None,"_y"))

out=out[out['_merge']=='left_only']

column_names=['Slot','coinIn','coinOut','jackpotHandpay','ProgressiveCoinOut','gamesPlayed',

              'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']


# Resta entre DFs

latest_file=latest_file.fillna(0)
previous_file=previous_file.fillna(0)

#previous_file=previous_file[previous_file['themeId']!=0]

df_difference=latest_file.set_index(['Slot'])[cols].subtract(previous_file.set_index(['Slot'])[cols])

# Corrijo los ram clears
latest_file=latest_file[['Slot','coinIn','coinOut','jackpotHandpay','ProgressiveCoinOut','gamesPlayed','GameBaseCoinOut',
                        'ScatterWon','FreeGamePlays','FreeGamesCoinOut','BounsCointOut']]

latest_file.set_index(['Slot'],inplace=True)
indexes=df_difference[df_difference['coinIn']<0].index
df_difference.loc[indexes]=latest_file.loc[indexes]
df_difference.reset_index(inplace=True)

"""
Lee el archivo inventario
"""
os.chdir(script_dir)
os.chdir('../Inventario')

inventario=pd.read_excel(r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\PBI Data Bases\Inventario\Inventario Winsystems Online.xlsx',sheet_name="Inventario total")
inventario=inventario[inventario['Online']==1]
inventario=inventario[inventario['ON']==1]

df_difference=inventario.merge(df_difference,how='left',left_on='Match online',right_on='Slot')
df_difference['Day']=latest_file_name.split('_')[1]
df_difference['month']=latest_file_name.split('_')[2]
df_difference['year']=latest_file_name.split('_')[3].split('.')[0]
df_difference['Date']=df_difference['year']+'-'+df_difference['month']+'-'+df_difference['Day']
df_difference['Date']=pd.to_datetime(df_difference['Date'])
df_difference['Date']=df_difference['Date'].dt.date

# Agrego transformación de restarle un día

df_difference['Date']=df_difference['Date']-pd.Timedelta(days=1)
os.chdir(script_dir)
os.chdir('../')
print ("Leyendo full_online_db_new")


# Lista de nombres de columnas
columnas = [
    'Grupo', 'Casino', 'Slot ID', 'Match online', 'Serie', 'Cabinet', 'Modelo', 'Mix', 'RTP TH', 'Hold TH',
    'Jackpot', 'Tipo JP', 'Area', 'Isla', 'Posicion', 'Pais', 'Estado', 'Ciudad', 'Municipio', 'TG SP',
    'TG CI', 'TG NW', 'Slot_x', 'time_x', 'Estatus', 'Modo', '% Participacion', 'Online', 'ON',
    'Positions Approached', 'Positions Denied', 'Positions Confirmed', 'Infraestructure Needed',
    'Infrastructure Exist', 'Scheduled', 'Positions Conected', 'Coments', '# Isla', 'Connected',
    'Programed', 'Pending', 'Date Connected', 'Date Programmed', 'Programmed Week', 'Conected Week',
    'Dated', 'Instalada', 'Tipo Movimiento', 'Detalle Movimiento', 'Fecha Movimiento', 'Origen',
    'Promotoria', 'Detalle Prom', 'Nombre Prom', 'Fecha de inicio de Op', 'Slot', 'coinIn', 'coinOut',
    'jackpotHandpay', 'ProgressiveCoinOut', 'gamesPlayed', 'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays',
    'FreeGamesCoinOut', 'BounsCointOut', 'Day', 'month', 'year', 'Date', 'Fecha alta', 'Owner', 'Razon Social',
    'Serial', 'RTP TH.1', '% JP', 'RTP Tot', 'TG CI.1', 'TG NW.1', 'TG Spins', 'TG Bet', 'TG RTP',
    'Status 2', 'Launched', 'Tipo de conexion', 'Gamemix'
]
""""""
def create_empty_excel_file(file_path, columns_list):
    # Creamos un DataFrame vacío con las columnas especificadas
    df = pd.DataFrame(columns=columns_list)

    # Creamos el archivo Excel utilizando el DataFrame vacío
    df.to_excel(file_path, index=False)

# Ruta del archivo Excel
archivo_excel = "results.xlsx"

# Crear el archivo si no existe o sobrescribirlo si ya existe
create_empty_excel_file(archivo_excel, columnas)

# Verificamos si el archivo ya existe
if os.path.exists(archivo_excel):
    # Si el archivo existe, sobrescribirlo
    create_empty_excel_file(archivo_excel, columnas)
    print(f"Se ha sobrescrito el archivo: {archivo_excel}")
else:
    # Si el archivo no existe, crear uno nuevo
    create_empty_excel_file(archivo_excel, columnas)
    print(f"Se ha creado un nuevo archivo: {archivo_excel}")


full_db=pd.read_excel(archivo_excel, engine='openpyxl')#full_db=full_db.append(df_difference,ignore_index=True)
full_db=full_db.append(df_difference)
print (full_db)

try:
    full_db.drop(['time','Slot_y','Unnamed: 0','Muncipio'],axis=1,inplace=True)

except:
    pass

# Agregar y/o Sacar columnas para que quede perfecto

full_db['Date'].fillna(dt.datetime.today()-pd.Timedelta(days=1),inplace=True)
del full_db['Muncipio']
print ("Escribiendo a Excel...")
full_db.to_excel(archivo_excel,index=False)