import re

import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe, validar_columnas


def extraer_velocidad(detalle):
    if pd.isna(detalle):
        return None

    texto = str(detalle).upper()
    patrones = [
        r'(\d+)\s*MG\b',
        r'@\s*(\d+)(?=\s*(?:\$|$))',
        r'\bSOLO\b.*?(\d+)(?=\s*(?:\$|$))',
    ]

    for patron in patrones:
        match = re.search(patron, texto)
        if match:
            return match.group(1) + 'MG'

    return None


def normalizar_velocidad_red(valor):
    if pd.isna(valor):
        return None

    match = re.search(r'(\d+)', str(valor).upper())
    if match:
        return match.group(1) + 'MG'
    return None


def es_plan_solo_internet(detalle):
    if pd.isna(detalle):
        return False
    return bool(re.search(r'\bsolo\b', str(detalle), flags=re.IGNORECASE))


def catv_esta_activo(valor):
    if pd.isna(valor):
        return False
    return str(valor).strip().lower() in {'enable', 'enabled'}


def procesar_verificacion_velocidad(saeplus_file, olt_file):
    saeplus = procesar_archivo_excel_solo(saeplus_file)
    olt = procesar_archivo_csv_solo(olt_file)

    resultado = pd.merge(
        saeplus, olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_OLT')
    )
    resultado = resultado.dropna(subset=['EQUIPO MACO'])
    resultado.columns = resultado.columns.str.lower()

    validar_columnas(
        resultado,
        [
            'detalle suscripcion', 'estatus', 'n° abonado', 'documento', 'nombre', 'name',
            'nombre franquicia', 'equipo maco', 'sn', 'olt',
            'service port upload speed', 'service port download speed', 'tipo tecnología.', 'catv'
        ],
        'resultado de velocidad'
    )

    resultado['velocidad_detalle'] = resultado['detalle suscripcion'].apply(extraer_velocidad)
    resultado['velocidad_red'] = resultado['service port download speed'].apply(normalizar_velocidad_red)
    resultado['condicion_velocidad'] = (
        resultado['velocidad_detalle'].fillna('') != resultado['velocidad_red'].fillna('')
    )
    resultado['condicion_solo_internet_con_catv'] = (
        resultado['detalle suscripcion'].apply(es_plan_solo_internet) &
        resultado['catv'].apply(catv_esta_activo)
    )
    resultado['motivo_revision'] = ''
    resultado.loc[resultado['condicion_velocidad'], 'motivo_revision'] = 'Velocidad no coincide'
    resultado.loc[
        resultado['condicion_solo_internet_con_catv'],
        'motivo_revision'
    ] = resultado.loc[
        resultado['condicion_solo_internet_con_catv'],
        'motivo_revision'
    ].apply(
        lambda valor: (
            'Plan solo internet con CATV activo'
            if not valor
            else f'{valor} | Plan solo internet con CATV activo'
        )
    )

    abonados_filtrados = resultado[
        (resultado['estatus'] == 'ACTIVO') &
        (
            resultado['condicion_velocidad'] |
            resultado['condicion_solo_internet_con_catv']
        )
    ]

    columnas_deseadas = [
        'n° abonado', 'documento', 'nombre', 'name', 'estatus',
        'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt',
        'service port upload speed', 'service port download speed', 'catv', 'tipo tecnología.', 'motivo_revision'
    ]
    abonados_filtrados = abonados_filtrados[columnas_deseadas]

    return {
        'data': abonados_filtrados,
        'num_casos': abonados_filtrados.shape[0],
        'excel': excel_desde_dataframe(abonados_filtrados, 'Resultado Filtrado'),
    }
