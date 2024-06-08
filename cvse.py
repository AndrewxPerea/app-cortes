import pandas as pd
import re

def procesar_archivo_csv(archivo):
    try:
        # Leer el archivo CSV con Pandas
        df = pd.read_csv(archivo, usecols=['SN', 'Name', 'OLT', 'CATV',"Administrative status", 'Service port upload speed'])
       
        
        # Crear una nueva columna con el primer dato antes del guion
        df['codigo'] = df['Name'].str.split(' - ').str[0]
        
        # Mostrar los primeros cinco registros
        print(df.head())
        
        return df
    
    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel(archivo):
    try:
        # Leer el archivo Excel con Pandas
        df = pd.read_excel(archivo, usecols=[0, 1, 2, 3, 4])
        
        # Fusionar las columnas "Nombre" y "Apellido" en una sola columna
        df['Nombre'] = df.apply(lambda row: ' '.join([str(row['Nombre']), str(row['Apellido'])]), axis=1)
        
        # Eliminar la columna "Apellido"
        df.drop(columns=['Apellido'], inplace=True)
        
        
        
        return df

    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

# Rutas de los archivos
abonados_file = "cortes.xlsx" 
cortes_file = "olt.csv"

# Procesar archivos
df_abonados = procesar_archivo_excel(abonados_file)
df_cortes = procesar_archivo_csv(cortes_file)

# Fusionar DataFrames si no son vacíos
if not df_abonados.empty and not df_cortes.empty:
    resultado = pd.merge(df_abonados, df_cortes, how='right', left_on='N° Abonado', right_on='codigo', suffixes=('_abonados', '_cortes'))
    resultado = resultado.dropna(subset=['N° Abonado'])
    resultado.columns = resultado.columns.str.lower()
    print(resultado.head())
    resultado.to_excel('fusion_resultado.xlsx', index=False)
    resultado_filtrado = resultado[
        (resultado['observaciones'].isna()) & 
        ((resultado['catv'] == 'Enabled') |
         (resultado['administrative status'] == 'Enabled') )
    ]
    resultado_filtrado.to_excel('fusion_resultado2.xlsx', index=False)
    print(resultado_filtrado.head())
