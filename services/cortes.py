import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe, validar_columnas


def procesar_cortes(abonados_file, cortes_file, sae_file):
    df_cortes = procesar_archivo_excel_solo(abonados_file)
    df_olt = procesar_archivo_csv_solo(cortes_file)
    df_saeplus = procesar_archivo_excel_solo(sae_file)

    resultado = pd.merge(
        df_saeplus, df_cortes, how='right', left_on='N° Abonado', right_on='N° Abonado'
    )
    resultado = resultado.dropna(subset=['N° Abonado'])
    resultado = pd.merge(
        resultado,
        df_olt,
        left_on='EQUIPO MACO_y',
        right_on='NSN',
        suffixes=('_abonados', '_cortes')
    )
    resultado = resultado.dropna(subset=['EQUIPO MACO_y'])
    resultado.columns = resultado.columns.str.lower()

    validar_columnas(
        resultado,
        [
            'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
            'estatus', 'observaciones', 'sn', 'olt',
            'catv', 'administrative status', 'status', 'ingeniero'
        ],
        'resultado de cortes'
    )

    columnas_deseadas = [
        'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
        'estatus', 'observaciones', 'sn', 'olt',
        'catv', 'administrative status', 'status', 'ingeniero'
    ]
    resultado_filtrado = resultado[columnas_deseadas]
    resultado_filtrado = resultado_filtrado[
        (resultado_filtrado['observaciones'].isna()) &
        (resultado_filtrado['estatus'] != 'ACTIVO') &
        (
            (resultado_filtrado['status'] == 'Online') |
            (resultado_filtrado['catv'] != 'Disabled') |
            (resultado_filtrado['administrative status'] == 'Enabled')
        )
    ]
    resultado_filtrado = resultado_filtrado.dropna(subset=['estatus'])

    return {
        'data': resultado_filtrado,
        'num_casos': resultado_filtrado.shape[0],
        'excel': excel_desde_dataframe(resultado_filtrado, 'Resultado Filtrado'),
    }
