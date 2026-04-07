import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe


def procesar_coincidencias_saeplus_smartolt(saeplus_file, olt_file):
    saeplus = procesar_archivo_excel_solo(saeplus_file)
    olt = procesar_archivo_csv_solo(olt_file)

    if 'EQUIPO MACO' not in saeplus.columns or 'NSN' not in olt.columns:
        raise ValueError("Los archivos no contienen las columnas requeridas 'EQUIPO MACO' y 'NSN'.")

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

    if 'N° Abonado' in resultado.columns:
        resultado = resultado.sort_values(by='N° Abonado')

    return {
        'data': resultado,
        'num_casos': resultado.shape[0],
        'excel': excel_desde_dataframe(resultado, 'Coincidencias'),
    }
