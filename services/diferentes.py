import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_hojas


def procesar_diferentes(saeplus_file, olt_file):
    saeplus = procesar_archivo_excel_solo(saeplus_file)
    olt = procesar_archivo_csv_solo(olt_file)

    if 'EQUIPO MACO' not in saeplus.columns or 'NSN' not in olt.columns:
        raise ValueError("Los archivos no contienen las columnas requeridas 'EQUIPO MACO' y 'NSN'.")

    resultado = pd.merge(
        saeplus,
        olt,
        how='outer',
        left_on='EQUIPO MACO',
        right_on='NSN',
        suffixes=('_abonados', '_cortes'),
        indicator=True
    )

    resultado_diferente = resultado[resultado['_merge'] != 'both']
    resultado_solo_equipo_mac = resultado_diferente.dropna(subset=['EQUIPO MAC']).dropna(axis=1, how='all')
    resultado_solo_nsn = resultado_diferente.dropna(subset=['NSN']).dropna(axis=1, how='all')

    return {
        'data': resultado_diferente,
        'num_casos': resultado_diferente.shape[0],
        'excel': excel_desde_hojas(
            [
                ('Diferentes', resultado_diferente),
                ('Solo en SAEPLUS', resultado_solo_equipo_mac),
                ('Solo en OLT', resultado_solo_nsn),
            ]
        ),
    }
