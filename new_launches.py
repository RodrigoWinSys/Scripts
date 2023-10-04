# -*- coding: utf-8 -*-
"""
Created on Mon May 17 11:38:08 2021

@author: r.deapellaniz
"""

from bs4 import BeautifulSoup as bs
import pandas as pd
import os
import datetime as dt
import numpy as np

from tkinter import *
from tkinter import ttk
import tkinter as tk
script_dir=os.getcwd()
def execute():
    directory=e1.get()    
    output_path=e2.get()
    slots=[]
    date=dt.datetime.today()
    denoms=[]
    names=[]
    values=[]
    themes=[]
    dates=[]
    casinos=[]
    first_games=[]
    last_games=[]
    slots_dates=[]
    rtps=[]
    os.chdir(directory)
   
    def ad_timestamp(timestamp):
        if timestamp != 0:
            return (dt.datetime(1,1,1)+dt.timedelta(seconds=int(timestamp)/10000000))
       
   
    for casino in os.listdir():
        #print("Leyendo carpeta")
        print (casino)
        os.chdir(directory+'//'+str(casino))
        content=[]
        for machine in os.listdir():           
            with open((directory+"/"+str(casino)+"/"+str(machine)+"/"+"DeviceManagerData.xml_1"), "r",encoding="latin-1") as file:
                # Read each line in the file, readlines() returns a list of lines
                content = file.readlines()
                # Combine the lines in the list into a string
                content = "".join(content)
                bs_content = bs(content, "lxml")         
                print (machine)
                performance =bs_content.find_all('d4p1:perfmeter')
                cab_meters=bs_content.find_all('d4p1:cabmeter')
                
            for meter in cab_meters:
                if (meter['d4p1:metername']=='lastGamePlayedDT'):
                    last_game=meter['d4p1:metervalue']
                    last_games.append(ad_timestamp(int(last_game)))
                if (meter['d4p1:metername']=='firstGamePlayedDT'):
                    firstgame=meter['d4p1:metervalue']
                    first_games.append(ad_timestamp(int(firstgame)))
                    slots_dates.append(machine)  
                    print(slots_dates)
                
            for k in performance:
                casinos.append(str(casino))
                slots.append(str(machine))
                denoms.append(k['d4p1:denomid'])
                themes.append(k['d4p1:themeid'])
                names.append(k['d4p1:metername'])
                values.append(k['d4p1:metervalue'])
                rtps.append(k['d4p1:paytableid'])
        

                #print("Hubo un error, no encuentro el archivo")
    df=pd.DataFrame({'Casino':casinos,'denom':denoms,'Slot':slots,'RTP THEO':rtps,'themeId':themes,'meter':names,'value':values})
    df['value']=pd.to_numeric(df['value'])
    print (df['value'])
    df=pd.pivot_table(df,index=['Casino','denom','Slot','themeId','RTP THEO'],columns=['meter'],values=['value'],aggfunc='first')
    df=df['value']
    df.reset_index(inplace=True)
    # Modificar la ruta relativa al archivo Games Detail
    os.chdir(script_dir)
    os.chdir('..')
    games_info=pd.read_excel('C:/Users/Rodrigo Rodriguez/OneDrive - Win Systems Group/Documents/Games Detail EGC 2022.xlsx')
    games_info=games_info[['Game','Lines','HF','Math','V','Vol','Base RTP','FF',	'FG F','FG RTP','BG RTP','Game Cycle','MxPrz','Side Bet (FB)','Cluster']]
    df=df.merge(games_info,how='left',left_on='themeId',right_on='Game')
    dates_df=pd.DataFrame({'Slot':slots_dates,'First Game':first_games,'Last Game':last_games})
    #dates_df.to_csv(output_path+'\DMD output dates'+str(dt.datetime.today().date())+'.csv',sep=',',decimal='.',index=False)
    #dates_df['First Game']=dates_df['First Game'].dt.date
    #dates_df['Last Game']=dates_df['Last Game'].dt.date
    #dates_df['DAYS']=(dates_df['Last Game']-dates_df['First Game']).dt.days
    #df=df.merge(dates_df,how='left',on='Slot')
    os.chdir(script_dir)
    #os.chdir('..\..\Inventario')
    inventario=pd.read_excel('C:/Users/Rodrigo Rodriguez/OneDrive - Win Systems Group/PBI Data Bases/Inventario/Inventario Winsystems Online.xlsx',sheet_name='Inventario total')
    
    df=df.merge(inventario,how='left',left_on='Slot',right_on='Match online')
    print (df)
    for col in df.columns:
        print(col)
    df.rename(columns={'Casino_x':'Casino','First Game':'First Date','Last Game':'Last Date','RTP THEO':'RTPO'},inplace=True)
    
    for i in ['Grupo','Casino','Slot ID','Match online','Serie','Cabinet',
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
                     'themeId','RTPO','denom','Last Date','First Date',
                     'Platform','Game','Release','Lines','Vol','V','Math','HF',
                     'Base RTP','FF','FG F','FG RTP','BG F','BG RTP','MxPrz',
                     'Game Cycle','Max Liability','Side Bet (FB)','Cluster','RTP TH','% JP','RTP Tot','TG CI','TG NW','TG Spins','TG Bet','TG RTP','Status 2','launched','Gamemix']:
        if (i not in df):
            df[str(i)]=""
            print ("Agregada columna vacía: ",i)
    try:      
        df=df[['Grupo','Casino','Slot ID','Match online','Serie','Cabinet',
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
                         'themeId','RTPO','denom','Last Date','First Date',
                         'Platform','Game','Release','Lines','Vol','V','Math','HF',
                         'Base RTP','FF','FG F','FG RTP','BG F','BG RTP','MxPrz',
                         'Game Cycle','Max Liability','Side Bet (FB)','Cluster','RTP TH','% JP','RTP Tot','TG CI','TG NW','TG Spins','TG Bet','TG RTP','Status 2','launched','Gamemix']]
    except:
        pass

    df.to_excel(output_path+'\DMD output Games Report'+str(dt.datetime.today().date())+'.xlsx',index=False,sheet_name='Data Base GA')

    
    #dates_df=pd.DataFrame({'Slot':slots_dates,'First Game':first_games,'Last Game':last_games})
    #dates_df.to_csv(output_path+'\DMD output dates'+str(dt.datetime.today().date())+'.csv',sep=',',decimal='.',index=False)
     
    os.chdir(directory)
    rtps=[]
    casinos=[]
    slots=[]
    themes=[]
    actives=[]
    denoms=[]
    for casino in os.listdir():
        print (casino)
        os.chdir(directory+'//'+str(casino))
        content=[]
        for machine in os.listdir():
            
            with open((directory+"/"+str(casino)+"/"+str(machine)+"/AurumSetup.xml"), "r",encoding="latin-1") as file:
                # Read each line in the file, readlines() returns a list of lines
                content = file.readlines()
                # Combine the lines in the list into a string
                content = "".join(content)
                bs_content = bs(content, "lxml")           
                print (machine)
            combo =bs_content.find_all('combo')
                          
            for k in combo:
                casinos.append(str(casino))
                slots.append(str(machine))
                denoms.append(k['d7p1:denomid'])
                themes.append(k['d7p1:themeid'])
                rtps.append(k['d7p1:theoreticalpaybackpercentage'])
                actives.append(k['d7p1:comboactive'])
    df=pd.DataFrame({'Casino':casinos,'denom':denoms,'Slot':slots,'theme':themes,'RTP':rtps,'Active':actives})
    df=df[df['Active']=='true']
    df.to_csv(output_path+'\DMD_config.csv',sep=',',decimal='.',index=False)
    
    # Ahora saco el archivo por slot
    
    
    root.destroy()



root = Tk()
root.wm_title('Hola Win')
frm = ttk.Frame(root, padding=10)
frm.grid()
ttk.Label(frm, text="Ingrese la ruta a los DMD: ").grid(column=0, row=0)
e1=ttk.Entry(frm)
e1.grid(column=1,row=0)
ttk.Label(frm, text="Ingrese la ruta donde quiere el output: ").grid(column=0,row=3)
e2=ttk.Entry(frm)
e2.grid(column=1,row=3)
"""
ttk.Label(frm, text="Ingrese la ruta al archivo iventario: ").grid(column=0,row=5)
e3=ttk.Entry(frm)
e3.grid(column=1,row=5)
"""
ttk.Button(frm, text="Iniciar", command=execute).grid(column=0, row=7)
root.mainloop()

