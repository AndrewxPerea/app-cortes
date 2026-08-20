from datetime import datetime
import io
import os
import zipfile
from xml.etree import ElementTree as ET

import pandas as pd


SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NSMAP = {'a': SPREADSHEET_NS}

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
    df = leer_excel_seguro(archivo_excel)
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
        df = pd.read_csv(archivo, low_memory=False)
        if 'SN' not in df.columns:
            raise ValueError("El archivo CSV no contiene la columna requerida 'SN'.")
        df['NSN'] = df['SN'].astype(str).str[-8:]
        return df

    except FileNotFoundError as e:
        raise FileNotFoundError(f"El archivo {archivo} no se encontró.") from e

    except Exception as e:
        raise ValueError(f"Ocurrió un error al procesar el CSV: {e}") from e

def procesar_archivo_excel_solo(archivo):
    try:
        df = leer_excel_seguro(archivo)
        if 'EQUIPO MAC' not in df.columns:
            raise ValueError("El archivo Excel no contiene la columna requerida 'EQUIPO MAC'.")
        df['EQUIPO MACO'] = df['EQUIPO MAC'].astype(str).str[-8:]
        return df

    except FileNotFoundError as e:
        raise FileNotFoundError(f"El archivo {archivo} no se encontró.") from e

    except Exception as e:
        raise ValueError(f"Ocurrió un error al procesar el Excel: {e}") from e

def normalizar_columnas(df, rename_col):
    df = df.rename(columns={df.columns[0]: rename_col})
    df.columns = df.columns.str.lower()
    return df


def obtener_bytes_archivo(archivo):
    if isinstance(archivo, (str, os.PathLike)):
        with open(archivo, 'rb') as stream:
            return stream.read()

    stream = getattr(archivo, 'stream', archivo)
    if hasattr(stream, 'seek'):
        stream.seek(0)

    data = stream.read()

    if hasattr(stream, 'seek'):
        stream.seek(0)

    if isinstance(data, str):
        return data.encode('utf-8')

    return data


def es_error_fill_openpyxl(error):
    return "expected <class 'openpyxl.styles.fills.Fill'>" in str(error)


def reparar_fills_estilos_xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data), 'r') as origen:
        archivos = {nombre: origen.read(nombre) for nombre in origen.namelist()}

    styles_path = 'xl/styles.xml'
    if styles_path not in archivos:
        return data

    root = ET.fromstring(archivos[styles_path])
    fills = root.find('a:fills', NSMAP)
    if fills is None:
        return data

    max_fill_id = 1
    for xf in root.findall('.//a:xf', NSMAP):
        fill_id = xf.attrib.get('fillId')
        if fill_id and fill_id.isdigit():
            max_fill_id = max(max_fill_id, int(fill_id))

    cantidad_existente = len(list(fills))
    cantidad_fills = max(cantidad_existente, max_fill_id + 1, 2)

    fills.clear()
    fills.set('count', str(cantidad_fills))

    for indice in range(cantidad_fills):
        fill = ET.SubElement(fills, f'{{{SPREADSHEET_NS}}}fill')
        pattern_fill = ET.SubElement(fill, f'{{{SPREADSHEET_NS}}}patternFill')
        if indice == 0:
            pattern_fill.set('patternType', 'none')
        elif indice == 1:
            pattern_fill.set('patternType', 'gray125')
        else:
            pattern_fill.set('patternType', 'solid')

    archivos[styles_path] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    reparado = io.BytesIO()
    with zipfile.ZipFile(reparado, 'w', zipfile.ZIP_DEFLATED) as destino:
        for nombre, contenido in archivos.items():
            destino.writestr(nombre, contenido)

    return reparado.getvalue()


def _ejecutar_lectura_excel(archivo, lector):
    source = getattr(archivo, 'stream', archivo)

    try:
        if hasattr(source, 'seek'):
            source.seek(0)
        return lector(source)
    except Exception as error:
        if not es_error_fill_openpyxl(error):
            raise

    data = obtener_bytes_archivo(archivo)
    reparado = reparar_fills_estilos_xlsx(data)
    return lector(io.BytesIO(reparado))


def leer_excel_seguro(archivo, **kwargs):
    return _ejecutar_lectura_excel(
        archivo,
        lambda source: pd.read_excel(source, **kwargs)
    )


def abrir_excel_seguro(archivo, **kwargs):
    return _ejecutar_lectura_excel(
        archivo,
        lambda source: pd.ExcelFile(source, **kwargs)
    )

def clasificar_estado_potencia(potencia):
    if potencia <= -33:
        return "Arpoón Critico"
    elif potencia <= -30:
        return "Arpón Alarmado"
    else:
        return "Normal"
