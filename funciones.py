from flask import  render_template
import pandas as pd
import numpy as np
from datetime import datetime

def obtener_valor_plan(plan):
    valor_plan_mapping = {
        '150 MG': 106000,
        '100 MG PA 5': 117000,
        '100 MG PA 6': 128000,
        '15 MG': 71000,
        '200 MG': 204000,
        '200 MG PA 5': 215000,
        '300 MG': 314000,
        '400 MG': 418000,
        '100 MG': 77000,
        '50 MG PA 5': 88000,
        'SOLO @ 50 MG': 64000,
        '30 MG': 70000,
        '70 MG': 87000,
        '20 MG': 54000,    }
    return valor_plan_mapping.get(plan, 'Plan no encontrado')

def generar_mensaje(row):
    if '@' in row['Plan Nuevo']:
        return (f"Estimad@ {row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud "
                f"de cambio de plan a {row['Plan Nuevo']}bps de solo internet, por un valor mensual "
                f"de $ {row['Valor Plan']} ha sido efectuado exitosamente. Con esto, procedemos a finalizar "
                f"tu petición. ¡Te deseamos un feliz día!")
    else:
        return (f"Estimad@ {row['Nombre Cliente']}, TuCable te informa que el estado de tu solicitud "
                f"de cambio de plan a {row['Plan Nuevo']}bps de internet, por un valor mensual "
                f"de $ {row['Valor Plan']} ha sido efectuado exitosamente. Con esto, procedemos a finalizar "
                f"tu petición. ¡Te deseamos un feliz día!")

def procesar_excel(archivo_excel):
    df = pd.read_excel(archivo_excel)

    df['Valor Plan'] = df['Plan Nuevo'].apply(obtener_valor_plan)
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str).apply(lambda x: x.split()[0].capitalize())
    df['Plan Nuevo'] = df['Plan Nuevo'].astype(str)

    # Generar los mensajes
    df['Mensaje'] = df.apply(generar_mensaje, axis=1)

    # Guardar el resultado en un archivo Excel
    output_file = f"resultado_procesado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(output_file, index=False)
    return output_file




def procesar_archivo_csv_solo(archivo):
    try:
        # Leer el archivo CSV sin fragmentarlo
        df = pd.read_csv(archivo, low_memory=False)
        df['NSN'] = df['SN'].astype(str).str[-8:] # Crear la columna NSN con los últimos 8 dígitos
        return df
    
    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return render_template('error.html', error=str(e))
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return render_template('error.html', error=str(e))  # Devolver un DataFrame vacío en caso de error

def procesar_archivo_excel_solo(archivo):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo)
        df['EQUIPO MACO'] = df['EQUIPO MAC'].astype(str).str[-8:]
        return df

    except FileNotFoundError:
        print(f"El archivo {archivo} no se encontró.")
        return render_template('error.html', error=str(e))
    
    except Exception as e:
        print(f"Ocurrió un error al procesar {archivo}:", e)
        return render_template('error.html', error=str(e)) 

