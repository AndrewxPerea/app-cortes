import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_hojas


def cargar_datos_equipos(saeplus_file, olt_file):
    saeplus = procesar_archivo_excel_solo(saeplus_file)
    olt = procesar_archivo_csv_solo(olt_file)

    if 'EQUIPO MACO' not in saeplus.columns or 'NSN' not in olt.columns:
        raise ValueError("Los archivos no contienen las columnas requeridas 'EQUIPO MACO' y 'NSN'.")

    return saeplus, olt


def ordenar_por_abonado(df):
    if 'N° Abonado' not in df.columns:
        return df

    trabajo = df.copy()
    texto = trabajo['N° Abonado'].fillna('').astype(str).str.strip()
    numerico = pd.to_numeric(texto, errors='coerce')
    trabajo['_orden_es_texto'] = numerico.isna().astype(int)
    trabajo['_orden_numerico'] = numerico.fillna(float('inf'))
    trabajo['_orden_texto'] = texto.str.upper()

    trabajo = trabajo.sort_values(
        by=['_orden_es_texto', '_orden_numerico', '_orden_texto'],
        ascending=[True, True, True],
        na_position='last'
    )
    return trabajo.drop(columns=['_orden_es_texto', '_orden_numerico', '_orden_texto'])


def describir_no_coincidencia(valor):
    if valor == 'left_only':
        return 'No coincide en SmartOLT'
    if valor == 'right_only':
        return 'No coincide en SAEPlus'
    return ''


def construir_coincidencias(saeplus, olt):
    resultado = pd.merge(
        saeplus,
        olt,
        how='inner',
        left_on='EQUIPO MACO',
        right_on='NSN',
        suffixes=('_abonados', '_smartolt')
    )
    resultado = resultado.dropna(subset=['EQUIPO MACO', 'NSN']).copy()
    resultado = resultado.dropna(axis=1, how='all')
    resultado = ordenar_por_abonado(resultado)

    return resultado.reset_index(drop=True)


def construir_no_coincidencias(saeplus, olt):
    resultado = pd.merge(
        saeplus,
        olt,
        how='outer',
        left_on='EQUIPO MACO',
        right_on='NSN',
        suffixes=('_abonados', '_cortes'),
        indicator=True
    )
    resultado = resultado[resultado['_merge'] != 'both'].copy()
    resultado = resultado.dropna(axis=1, how='all')
    resultado.insert(0, 'no coincide en', resultado['_merge'].map(describir_no_coincidencia))
    resultado = ordenar_por_abonado(resultado)

    return resultado.reset_index(drop=True)


def procesar_comparativo_equipos(saeplus_file, olt_file):
    saeplus, olt = cargar_datos_equipos(saeplus_file, olt_file)
    coincidencias = construir_coincidencias(saeplus, olt)
    no_coincidencias = construir_no_coincidencias(saeplus, olt)

    return {
        'data': coincidencias,
        'columns': coincidencias.columns,
        'num_casos': int(coincidencias.shape[0] + no_coincidencias.shape[0]),
        'excel': excel_desde_hojas(
            [
                ('Equipos que coinciden', coincidencias),
                ('Equipos que no coinciden', no_coincidencias),
            ]
        ),
    }
