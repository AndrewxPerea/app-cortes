from pathlib import Path

import pandas as pd

from funciones import abrir_excel_seguro
from services.common import excel_desde_hojas


PREFIJOS_CIUDAD = [
    ('TCF', 'PereiraCentro'),
    ('CH', 'Chinchiná'),
    ('C0', 'Cartago'),
    ('DQ', 'Dosquebradas'),
    ('SR', 'Santa Rosa'),
    ('VG', 'Virginia'),
    ('TC', 'Pereira'),
    ('PQ', 'Parque Industrial'),
    ('SG', 'Guaviare'),
]


def obtener_ciudad_desde_abonado(valor):
    if pd.isna(valor):
        return None

    texto = str(valor).strip().upper()
    if not texto:
        return None

    for prefijo, ciudad in PREFIJOS_CIUDAD:
        if texto.startswith(prefijo):
            return ciudad

    return None


def agregar_columna_ciudad(df):
    if 'ABONADO' not in df.columns:
        raise ValueError("El archivo no contiene la columna requerida 'ABONADO'.")

    resultado = df.copy()
    resultado['CIUDAD'] = resultado['ABONADO'].apply(obtener_ciudad_desde_abonado)
    return resultado


def procesar_ciudades_abonado(archivo_excel):
    excel = abrir_excel_seguro(archivo_excel)
    hojas_procesadas = []
    vistas_previas = []

    for nombre_hoja in excel.sheet_names:
        df = excel.parse(sheet_name=nombre_hoja)

        if 'ABONADO' not in df.columns:
            continue

        hoja_resultado = agregar_columna_ciudad(df)
        hojas_procesadas.append((nombre_hoja[:31], hoja_resultado))

        vista = hoja_resultado.copy()
        vista.insert(0, 'HOJA ORIGEN', nombre_hoja)
        vistas_previas.append(vista)

    if not hojas_procesadas:
        raise ValueError("No se encontró la columna 'ABONADO' en ninguna hoja del archivo.")

    data = pd.concat(vistas_previas, ignore_index=True, sort=False)
    return {
        'data': data,
        'columns': data.columns.tolist(),
        'num_casos': int(data.shape[0]),
        'excel': excel_desde_hojas(hojas_procesadas),
    }


def guardar_ciudades_abonado(input_path, output_path):
    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"No existe el archivo: {input_path}")

    resultado = procesar_ciudades_abonado(input_path)
    output_path.write_bytes(resultado['excel'].getvalue())
    return resultado
