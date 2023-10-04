import os
import pandas as pd
import numpy as np
import datetime as dt

script_dir = os.getcwd()
older = r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\PBI Data Bases\New Launches\Mexico\DMDs Adventure\Mexico\Grand Palacio\19-Sept\grand_palacio_third.xlsx'
newer = r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\PBI Data Bases\New Launches\Mexico\DMDs Adventure\Mexico\Grand Palacio\28-Sept\grand_palacio_fourth.xlsx'
previous_file = pd.read_excel(older)
latest_file = pd.read_excel(newer)

# Columnas a restar
numeric_columns = ['coinIn', 'coinOut', 'jackpotHandpay', 'ProgressiveCoinOut', 'gamesPlayed',
                   'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']

# Convierto todo a numerico
cols = [i for i in latest_file.columns if i in numeric_columns]
for col in cols:
    latest_file[col] = pd.to_numeric(latest_file[col])

for col in cols:
    previous_file[col] = pd.to_numeric(previous_file[col])

# Restar las columnas y almacenar la diferencia en df_difference
df_difference = latest_file.groupby(['Slot', 'themeId', 'denom'])[numeric_columns].sum().subtract(
    previous_file.groupby(['Slot', 'themeId', 'denom'])[numeric_columns].sum(), fill_value=0)

# Sumar los totales por las combinaciones de 'Slot', 'themeId' y 'denom'
df_total = df_difference.groupby(['Slot', 'themeId', 'denom'])[numeric_columns].sum().reset_index()

# Dividir las columnas entre 100
#numeric_columns_to_divide = ['coinIn', 'coinOut', 'jackpotHandpay', 'ProgressiveCoinOut', 'GameBaseCoinOut', 'FreeGamesCoinOut', 'BounsCointOut']
#df_total[numeric_columns_to_divide] = df_total[numeric_columns_to_divide] / 100

# Agregar una columna 'Verificar' basada en los totales
df_total['Verificar'] = np.where(df_total[numeric_columns].lt(0).any(axis=1), 'revisar', 'check')

os.chdir(script_dir)
os.chdir('../')
df_total.to_excel('Totales' + str(dt.datetime.today().date()) + '.xlsx', index=False)
print("Documento guardado en " + script_dir)
