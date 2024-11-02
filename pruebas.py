import pandas as pd
import re
from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo

# Cargar archivos
saeplus = procesar_archivo_excel_solo('saeplus.xlsx')
olt2 = procesar_archivo_csv_solo('olt.csv')

# Filtrar los datos de 'olt' para tener solo los registros "Online"
olt = olt2[olt2['Status'] == 'Online']

# Fusionar los datos de abonados y OLT
if not saeplus.empty and not olt.empty:
    # Fusionar solo los registros coincidentes
    resultado = pd.merge(saeplus, olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
    resultado = resultado.dropna(subset=['EQUIPO MACO'])
    resultado.columns = resultado.columns.str.lower()

# Fusionar los datos incluyendo los registros no coincidentes para los diferentes registros
if not saeplus.empty and not olt2.empty:
    # Confirmar que existen las columnas 'EQUIPO MACO' y 'NSN'
    if 'EQUIPO MACO' in saeplus.columns and 'NSN' in olt2.columns:
        # Hacer la fusión con 'indicator=True' para identificar los registros coincidentes
        resultado = pd.merge(saeplus, olt2, how='outer', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'), indicator=True)
        
        if '_merge' in resultado.columns:
            # Filtrar registros donde 'EQUIPO MACO' y 'NSN' no coinciden
            resultado_diferente = resultado[resultado['_merge'] != 'both']
            
            # Dividir en diferentes DataFrames según la existencia de valores en 'EQUIPO MACO' y 'NSN'
            resultado_todos_diferentes = resultado_diferente  # Todos los registros diferentes
            resultado_solo_equipo_mac = resultado_diferente.dropna(subset=['EQUIPO MAC'])  # Solo registros con 'EQUIPO MAC'
            resultado_solo_nsn = resultado_diferente.dropna(subset=['SN'])  # Solo registros con 'SN'

            # Crear el archivo Excel con las tres hojas
            with pd.ExcelWriter('olt_diferente.xlsx') as writer:
                resultado_todos_diferentes.to_excel(writer, sheet_name='Todos los Registros Diferentes', index=False)
                resultado_solo_equipo_mac.to_excel(writer, sheet_name='Solo EQUIPO MAC', index=False)
                resultado_solo_nsn.to_excel(writer, sheet_name='Solo NSN', index=False)
            
            print("Archivo 'olt_diferente.xlsx' creado con las tres hojas.")
        else:
            print("No se encontró la columna '_merge' en el resultado.")
    else:
        print("Columnas 'EQUIPO MACO' o 'NSN' faltantes en uno de los DataFrames.")
else:
    print("Uno o ambos DataFrames están vacíos.")



