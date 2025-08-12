from flask import  render_template
import pandas as pd
import numpy as np
from datetime import datetime

def obtener_valor_plan(plan):
    valor_plan_mapping = {
        'GPON 150 MG $112.000': 112000,
        'GPON 100 MG PA 5 $117.000': 117000,
        'GPON 100 MG PA 6 $128.000': 128000,
        'GPON 15 MG $71000': 71000,
        'GPON 250 MG $211000': 211000,
        'GPON 200 MG PA 5 $215000': 215000,
        'GPON 350 MG $324000': 324000,
        'GPON 450 MG $431000': 431000,
        'GPON 550 MG $537000': 537000,
        'GPON 100 MG $82.000': 82000,
        'GPON 50 MG PA 5 $88000': 88000,
        'SOLO @ 100 MG $80.000': 80000,
        'GPON 30 MG $70000': 70000,
        'GPON 70 MG $87000': 87000,
        'GPON 20 MG $54000': 54000,  
        'VIP 50 MG $70.000' : 70000,
        'SOLO @ 30 MG' : 66000,
        'GPON 30 MG $58.000' : 58000,
        }
    return valor_plan_mapping.get(plan, 'Plan no encontrado')

def generar_mensaje(row):
    if '@' in row['Plan Nuevo']:
        return (f"Estimad@ {row['Nombre']}, TuCable te informa que el estado de tu solicitud "
                f"de cambio de plan a {row['Plan Nuevo']}bps de solo internet, por un valor mensual "
                f"de $ {row['Valor Plan']} ha sido efectuado exitosamente. Con esto, procedemos a finalizar "
                f"tu petición. ¡Te deseamos un feliz día!")
    else:
        return (f"Estimad@ {row['Nombre']}, TuCable te informa que el estado de tu solicitud "
                f"de cambio de plan a {row['Plan Nuevo']}bps de internet, por un valor mensual "
                f"de $ {row['Valor Plan']} ha sido efectuado exitosamente. Con esto, procedemos a finalizar "
                f"tu petición. ¡Te deseamos un feliz día!")

def procesar_excel(archivo_excel):
    df = pd.read_excel(archivo_excel)
    df.columns = df.columns.str.title   ()

    df['Valor Plan'] = df['Plan Nuevo'].apply(obtener_valor_plan)
    df['Nombre'] = df['Nombre'].astype(str).apply(lambda x: x.split()[0].capitalize())
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

def normalizar_columnas(df, rename_col):
                df = df.rename(columns={df.columns[0]: rename_col})
                df.columns = df.columns.str.lower()      
                return df