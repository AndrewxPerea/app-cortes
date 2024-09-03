

import pandas as pd

def procesar_excel(archivo_excel):
    df = pd.read_excel(archivo_excel)

    # Definir el mapeo de valores del plan
    valor_plan_mapping = {
        '100 MG': 106000,
        '100 MG PA 5': 117000,
        '100 MG PA 6': 128000,
        '15 MG': 71000,
        '200 MG': 204000,
        '200 MG PA 5': 215000,
        '300 MG': 314000,
        '400 MG': 418000,
        '50 MG': 77000,
        '50 MG PA 5': 88000,
        'SOLO @ 50 MG': 64000,
        '30 MG': 70000,
        '70 MG': 87000
    }

    # Procesar cada fila
    df['Valor Plan'] = df['Plan Nuevo'].map(valor_plan_mapping)
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str)
    df['Nombre Cliente'] = df['Nombre Cliente'].apply(lambda x: x.split()[0].capitalize())
    df['Plan Nuevo'] = df['Plan Nuevo'].astype(str)

    # Generar el mensaje personalizado
    df['Mensaje'] = df.apply(lambda row: (
            f"Estimad@ {row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud de cambio de plan a "
            f"{row['Plan Nuevo']} Mbps de solo internet, por un valor mensual de $ {row['Valor Plan']} ha sido efectuado exitosamente. "
            "Con esto, procedemos a finalizar tu petición. ¡Te deseamos un feliz día!"
            if '@' in row['Plan Nuevo'] else
            f"Estimad@ {row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud de cambio de plan a "
            f"{row['Plan Nuevo']}Mbps de internet, por un valor mensual de $ {row['Valor Plan']} ha sido efectuado exitosamente. "
            "Con esto, procedemos a finalizar tu petición. ¡Te deseamos un feliz día!"
        ), axis=1)

    # Guardar el resultado en un nuevo archivo Excel
    output_file = "resultado_procesado.xlsx"
    df.to_excel(output_file, index=False)

    return output_file


def procesar_archivo_csv(archivo):
    try:
        df = pd.read_csv(archivo, usecols=['SN', 'Name', 'OLT', 'CATV', "Administrative status", 'Service port upload speed'])
        df['codigo'] = df['Name'].str.split(' - ').str[0]
        return df
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel(archivo):
    try:
        df = pd.read_excel(archivo, usecols=[0, 1, 2, 3, 4])
        df['Nombre'] = df.apply(lambda row: ' '.join([str(row['Nombre']), str(row['Apellido'])]), axis=1)
        df.drop(columns=['Apellido'], inplace=True)
        return df
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_csv_solo(archivo):
    try:
        # Leer el archivo CSV sin fragmentarlo
        df = pd.read_csv(archivo, low_memory=False)
        df['NSN'] = df['SN'].astype(str).str[-8:] # Crear la columna NSN con los últimos 8 dígitos
        return df
    
    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return pd.DataFrame()  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel_solo(archivo):
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
    