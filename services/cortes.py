import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe, validar_columnas


def _normalizar_nombre_columna(columna):
    return str(columna).strip().lower()


def _columnas_por_nombre_normalizado(df, nombres):
    nombres = {_normalizar_nombre_columna(nombre) for nombre in nombres}
    return [
        columna
        for columna in df.columns
        if _normalizar_nombre_columna(columna) in nombres
    ]


def _descartar_columnas_por_nombre_normalizado(df, nombres):
    columnas = _columnas_por_nombre_normalizado(df, nombres)
    return df.drop(columns=columnas, errors='ignore')


def _normalizar_columna_estatus_saeplus(df):
    if 'estatus' in df.columns:
        return df

    columnas_estatus = _columnas_por_nombre_normalizado(
        df,
        ['estatus', 'estatus sae', 'estatus saeplus']
    )
    if not columnas_estatus:
        return df

    return df.rename(columns={columnas_estatus[0]: 'estatus'})


def procesar_cortes(abonados_file, cortes_file, sae_file):
    df_cortes = procesar_archivo_excel_solo(abonados_file)
    df_olt = procesar_archivo_csv_solo(cortes_file)
    df_saeplus = _normalizar_columna_estatus_saeplus(procesar_archivo_excel_solo(sae_file))
    df_cortes = _descartar_columnas_por_nombre_normalizado(df_cortes, ['estatus', 'ingeniero'])

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
    resultado.columns = resultado.columns.str.strip().str.lower()

    validar_columnas(
        resultado,
        [
            'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
            'estatus', 'observaciones', 'sn', 'olt',
            'catv', 'administrative status', 'status'
        ],
        'resultado de cortes'
    )

    columnas_deseadas = [
        'n° abonado', 'documento_x', 'nombre_x', 'apellido_x',
        'estatus', 'observaciones', 'sn', 'olt',
        'catv', 'administrative status', 'status'
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
