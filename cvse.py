import pandas as pd

def procesar_archivo_csv(archivo):
    try:
        # Leer el archivo CSV sin fragmentarlo
        df = pd.read_csv(archivo, low_memory=False)
        df['NSN'] = df['SN'].astype(str).str[-8:]  # Crear la columna NSN con los últimos 8 dígitos
        return df
    
    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel(archivo):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo)
        df['EQUIPO MACO'] = df['EQUIPO MAC'].astype(str).str[-8:]
        return df

    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

# Rutas de los archivos
solointernet = "dosquebradas.xlsx" 
olt = "guaviare2.csv"

# Procesar archivos
dfsolointernet = procesar_archivo_excel(solointernet)
dfolt = procesar_archivo_csv(olt)


# Fusionar DataFrames si no son vacíos
if not dfsolointernet.empty and not dfolt.empty:
    resultado = pd.merge(dfsolointernet, dfolt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
    resultado = resultado.dropna(subset=['EQUIPO MACO'])
    resultado.to_excel('fusion_resultado.xlsx', index=False)
    abonados_filtrados = resultado[
        (resultado['Detalle Suscripcion'].str.contains('@', na=False)) & 
                ((resultado['CATV'] == 'Enabled'))]
    if not abonados_filtrados.empty:
        abonados_filtrados.to_excel('abonados__filtrados.xlsx', index=False)
        print("Archivo 'abonados__filtrados.xlsx' creado con éxito.")
    else:
        print("No se encontraron abonados que no coincidan con '@' en 'Detalle Suscripcion'.")
else:
    print("Uno o ambos DataFrames están vacíos.")
