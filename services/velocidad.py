import re

import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe, validar_columnas


def extraer_velocidad(detalle):
    match = re.search(r'(\d+)\s*MG', detalle.upper())
    if match:
        return match.group(1) + 'MG'
    return None


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
            'service port upload speed', 'service port download speed', 'tipo tecnología.'
        ],
        'resultado de velocidad'
    )

    resultado['velocidad_detalle'] = resultado['detalle suscripcion'].apply(extraer_velocidad)
    abonados_filtrados = resultado[
        (resultado['velocidad_detalle'] != resultado['service port download speed']) &
        (resultado['estatus'] == 'ACTIVO')
    ]

    columnas_deseadas = [
        'n° abonado', 'documento', 'nombre', 'name', 'estatus',
        'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt',
        'service port upload speed', 'service port download speed', 'tipo tecnología.'
    ]
    abonados_filtrados = abonados_filtrados[columnas_deseadas]

    return {
        'data': abonados_filtrados,
        'num_casos': abonados_filtrados.shape[0],
        'excel': excel_desde_dataframe(abonados_filtrados, 'Resultado Filtrado'),
    }
