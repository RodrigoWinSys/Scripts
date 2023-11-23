import pandas as pd
import re

# Función para extraer la fecha y hora en el formato deseado
def extract_datetime(datetime_str):
    match = re.search(r'(\w{3} \w{3} \d{2} \d{4} \d{2}:\d{2}:\d{2})', datetime_str)
    if match:
        return pd.to_datetime(match.group(), format='%a %b %d %Y %H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
    else:
        return None

# Rutas de los dos archivos CSV a combinar
abiertos = r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\Documents\Tickets\abiertos_17_nov.csv'
resueltos = r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\Documents\Tickets\resueltos_17_nov.csv'

# Nombre del archivo Excel de salida (acumulado)
archivo_salida = r'C:\Users\Rodrigo Rodriguez\OneDrive - Win Systems Group\Documents\Tickets\data_base_tickets_resume.xlsx'

def combinar_archivos_csv(archivo1, archivo2, archivo_salida):
    # Leer los archivos CSV en DataFrames
    df1 = pd.read_csv(abiertos)
    df2 = pd.read_csv(resueltos)

    # Añadir columnas faltantes 'Fecha Atención' y 'Serie de refaccion origen' a df2
    if 'Fecha Atención' not in df2.columns:
        df2['Fecha Atención'] = None
    if 'Serie de refaccion origen' not in df2.columns:
        df2['Serie de refaccion origen'] = None

    # Alinear los DataFrames por los encabezados de columnas de archivo 1
    df2 = df2[df1.columns]

    # Concatenar los DataFrames
    df_acumulado = pd.concat([df1, df2], ignore_index=True, sort=False)

    # Cambiar "awaiting_agent" a "opened" en la columna "status_id"
    df_acumulado['status_id'] = df_acumulado['status_id'].replace('awaiting_agent', 'opened')

    # Formatear las columnas 'date_created' en el formato deseado y separar la hora
    df_acumulado['date_created'] = df_acumulado['date_created'].apply(lambda x: extract_datetime(x) if pd.notna(x) else x)
    df_acumulado['hour_created'] = df_acumulado['date_created'].apply(lambda x: x.split(' ')[-1] if pd.notna(x) else x)
    df_acumulado['date_created'] = df_acumulado['date_created'].apply(lambda x: x.split(' ')[0] if pd.notna(x) else x)

    # Formatear la columna 'Fecha Atención' en el formato deseado y separar la hora
    df_acumulado['Fecha Atención'] = df_acumulado['Fecha Atención'].apply(lambda x: extract_datetime(x) if pd.notna(x) else x)
    df_acumulado['hour_atencion'] = df_acumulado['Fecha Atención'].apply(lambda x: x.split(' ')[-1] if pd.notna(x) else x)
    df_acumulado['Fecha Atención'] = df_acumulado['Fecha Atención'].apply(lambda x: x.split(' ')[0] if pd.notna(x) else x)

    # Formatear la columna 'date_resolved' en el formato deseado y separar la hora
    df_acumulado['date_resolved'] = df_acumulado['date_resolved'].apply(lambda x: extract_datetime(x) if pd.notna(x) else x)
    df_acumulado['hour_resolved'] = df_acumulado['date_resolved'].apply(lambda x: x.split(' ')[-1] if pd.notna(x) else x)
    df_acumulado['date_resolved'] = df_acumulado['date_resolved'].apply(lambda x: x.split(' ')[0] if pd.notna(x) else x)

    # Guardar el DataFrame resultante en un archivo Excel con una segunda hoja
    with pd.ExcelWriter(archivo_salida, engine='xlsxwriter') as writer:
        df_acumulado.to_excel(writer, sheet_name='Hoja1', index=False)
        print("La primera hoja (Hoja1) ha sido creada exitosamente.")

        # Verificar si df_acumulado contiene datos antes de crear la segunda hoja
        if not df_acumulado.empty:
            # Crear una segunda hoja (Hoja2) con ciertas columnas del DataFrame acumulado
            columnas_deseadas = ['id', 'agent_name', 'status_id', 'date_created', 'followers','hour_created', 'Serie Maquina', 'Maquina Offline', 'Modelo Maquina', 'Casino', 'Grupo', 'Refacción Necesaria', 'Descripción de refacción', 'ST', 'Pedido realizado - Status', 'Fecha Atención', 'hour_atencion', 'date_resolved', 'hour_resolved', 'Foraneo']  # Reemplaza con las columnas deseadas
            df2_hoja2 = df_acumulado[columnas_deseadas]
            df2_hoja2.to_excel(writer, sheet_name='Resume', index=False)


combinar_archivos_csv(abiertos, resueltos, archivo_salida)
print("El archivo acumulado se encuentra en:", archivo_salida)
