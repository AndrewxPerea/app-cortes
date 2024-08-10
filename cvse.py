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
abonados = "abonados.xlsx" 
olt = "olt.csv"

# Procesar archivos
dfequipos = procesar_archivo_excel(abonados)
dfolt = procesar_archivo_csv(olt)


# Fusionar DataFrames si no son vacíos
if not dfequipos.empty and not dfolt.empty:
    # Realiza la fusión de los DataFrames
    resultado = pd.merge(dfequipos, dfolt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes'))
    resultado = resultado.dropna(subset=['EQUIPO MACO'])
    
    # Filtra los datos según las condiciones dadas
    abonados_filtrados = resultado[
        (resultado['Estatus'].str.lower().isin(['activo', 'por instalar']) == False) &  # Filtra los que no sean "estatus activo" o "estatus por instalar"
        ((resultado['Administrative status'].str.lower() == 'enabled') |  # Y que tengan "administrative status" o "catv" en "enabled"
         (resultado['CATV'].str.lower() == 'enabled'))
    ]
    
    # Guarda el DataFrame fusionado
    resultado.to_excel('fusion_resultado.xlsx', index=False)
    
    # Verifica si existen abonados filtrados y guarda el archivo
    if not abonados_filtrados.empty:
        abonados_filtrados.to_excel('abonados_filtrados.xlsx', index=False)
        print("Archivo 'abonados_filtrados.xlsx' creado con éxito.")
    else:
        print("No se encontraron abonados que coincidan con las condiciones especificadas.")
else:
    print("Uno o ambos DataFrames están vacíos.")

