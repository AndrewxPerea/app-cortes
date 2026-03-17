import pandas as pd

from funciones import normalizar_columnas
from services.common import excel_desde_dataframe, validar_columnas


def procesar_reconexiones(abonados_file, cortes_file):
    df_cortes = pd.read_excel(cortes_file)
    df_abonados = pd.read_excel(abonados_file)

    df_cortes = normalizar_columnas(df_cortes, 'Abonados')
    df_abonados = normalizar_columnas(df_abonados, 'Abonados')
    df_cortes.columns = df_cortes.columns.str.lower()
    df_abonados.columns = df_abonados.columns.str.lower()
    validar_columnas(df_cortes, ['abonados'], 'cortes')
    validar_columnas(df_abonados, ['abonados'], 'abonados')

    df_resultado = pd.merge(df_cortes, df_abonados, on="abonados", how="inner")
    validar_columnas(
        df_resultado,
        ['abonados', 'documento_x', 'nombre_x', 'apellido_x', 'observaciones', 'estatus_y'],
        'resultado de reconexiones'
    )

    df_resultado = df_resultado[
        ['abonados', 'documento_x', 'nombre_x', 'apellido_x', 'observaciones', 'estatus_y']
    ]
    df_resultado = df_resultado[
        (df_resultado['observaciones'].isna() | (df_resultado['observaciones'] == '')) &
        (df_resultado['estatus_y'] == 'ACTIVO')
    ]

    return {
        'data': df_resultado,
        'num_casos': df_resultado.shape[0],
        'excel': excel_desde_dataframe(df_resultado, 'Resultado'),
    }
