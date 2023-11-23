import pandas as pd

# Cargar el archivo Excel
archivo_excel = 'C:/Users/Rodrigo Rodriguez/OneDrive - Win Systems Group/Desktop/prestige_fusion.xlsx'
hoja1 = pd.read_excel(archivo_excel, sheet_name='Hoja1')
hoja2 = pd.read_excel(archivo_excel, sheet_name='Hoja2')

# Definir las columnas a sumar
columnas_a_sumar = ['coinIn', 'coinOut', 'jackpotHandpay', 'ProgressiveCoinOut', 'gamesPlayed',
                    'GameBaseCoinOut', 'ScatterWon', 'FreeGamePlays', 'FreeGamesCoinOut', 'BounsCointOut']

# Sumar las columnas para la Hoja 1
suma_hoja1 = hoja1.groupby(['Slot', 'themeId', 'denom'])[columnas_a_sumar].sum().reset_index()

# Unir los resultados de la suma de la Hoja 1 con la Hoja 2
df_resultado = pd.merge(suma_hoja1, hoja2, on=['Slot', 'themeId', 'denom'], how='inner', suffixes=('_hoja1', '_hoja2'))

# Crear columnas totales
columnas_totales = [col + '_total' for col in columnas_a_sumar]

# Sumar las columnas correspondientes para obtener las columnas totales
for col in columnas_a_sumar:
    df_resultado[col + '_total'] = df_resultado[col + '_hoja1'] + df_resultado[col + '_hoja2']

# Seleccionar solo las columnas deseadas en el resultado final
columnas_resultado_final = ['Slot', 'themeId', 'denom'] + columnas_totales
df_resultado_final = df_resultado[columnas_resultado_final].copy()

# Guardar el resultado en una tercera hoja
nombre_nueva_hoja = 'Hoja3'
with pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a') as writer:
    df_resultado_final.to_excel(writer, sheet_name=nombre_nueva_hoja, index=False)
