import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_dataframe, excel_desde_hojas, validar_columnas


ANALISIS_NAVEGACION = {
    'activos_sin_navegar': 'Activos sin navegar',
    'desactivos_con_internet': 'Desactivos con internet',
    'activos_sin_catv': 'Activos sin Catv',
    'solo_con_arroba_y_catv_activo': 'Solo con @ y catv activo',
}


def preparar_resultado_navegacion(abonados_file, olt_file):
    df_abonados = procesar_archivo_excel_solo(abonados_file)
    df_olt = procesar_archivo_csv_solo(olt_file)

    resultado = pd.merge(
        df_abonados,
        df_olt,
        how='right',
        left_on='EQUIPO MACO',
        right_on='NSN',
        suffixes=('_abonados', '_cortes')
    )
    df_resultado = resultado.dropna(subset=['EQUIPO MACO']).copy()
    df_resultado.columns = df_resultado.columns.str.lower()

    validar_columnas(
        df_resultado,
        [
            'n° abonado', 'documento', 'nombre', 'name', 'estatus', 'status',
            'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt', 'board', 'port',
            'service port upload speed', 'service port download speed', 'tipo tecnología.', 'catv', 'administrative status'
        ],
        'resultado de navegacion'
    )

    columnas_deseadas = [
        'n° abonado', 'documento', 'nombre', 'name', 'estatus', 'status',
        'detalle suscripcion', 'nombre franquicia', 'equipo maco', 'sn', 'olt', 'board', 'port',
        'service port upload speed', 'service port download speed', 'tipo tecnología.', 'catv', 'administrative status'
    ]
    return df_resultado[columnas_deseadas]


def construir_analisis_navegacion(df_resultado):
    estatus = df_resultado['estatus'].fillna('')
    status = df_resultado['status'].fillna('')
    detalle = df_resultado['detalle suscripcion'].fillna('')
    catv = df_resultado['catv'].fillna('')
    administrative_status = df_resultado['administrative status'].fillna('')

    return {
        'activos_sin_navegar': df_resultado[
            (estatus == 'ACTIVO') &
            (
                (administrative_status == 'Disabled') |
                (status != 'Online')
            )
        ].copy(),
        'desactivos_con_internet': df_resultado[
            (~estatus.str.lower().isin(['activo', 'por instalar'])) &
            (status.str.lower() == 'online')
        ].copy(),
        'activos_sin_catv': df_resultado[
            (~detalle.str.contains('@', na=False)) &
            (estatus == 'ACTIVO') &
            (catv == 'Disabled')
        ].copy(),
        'solo_con_arroba_y_catv_activo': df_resultado[
            detalle.str.contains('@', na=False) &
            (catv == 'Enabled')
        ].copy(),
    }


def crear_resultado_analisis(df, nombre_hoja):
    return {
        'data': df,
        'columns': df.columns,
        'num_casos': df.shape[0],
        'excel': excel_desde_dataframe(df, nombre_hoja),
    }


def procesar_sin_navegar(abonados_file, olt_file):
    df_resultado = preparar_resultado_navegacion(abonados_file, olt_file)
    analisis = construir_analisis_navegacion(df_resultado)

    return {
        'data': analisis['activos_sin_navegar'],
        'columns': df_resultado.columns,
        'num_casos': sum(df.shape[0] for df in analisis.values()),
        'excel': excel_desde_hojas(
            [
                (ANALISIS_NAVEGACION['activos_sin_navegar'], analisis['activos_sin_navegar']),
                (ANALISIS_NAVEGACION['desactivos_con_internet'], analisis['desactivos_con_internet']),
                (ANALISIS_NAVEGACION['activos_sin_catv'], analisis['activos_sin_catv']),
                (ANALISIS_NAVEGACION['solo_con_arroba_y_catv_activo'], analisis['solo_con_arroba_y_catv_activo']),
            ]
        ),
    }


def procesar_navegacion_activos_sin_navegar(abonados_file, olt_file):
    df_resultado = preparar_resultado_navegacion(abonados_file, olt_file)
    analisis = construir_analisis_navegacion(df_resultado)
    return crear_resultado_analisis(
        analisis['activos_sin_navegar'],
        ANALISIS_NAVEGACION['activos_sin_navegar']
    )


def procesar_navegacion_desactivos_con_internet(abonados_file, olt_file):
    df_resultado = preparar_resultado_navegacion(abonados_file, olt_file)
    analisis = construir_analisis_navegacion(df_resultado)
    return crear_resultado_analisis(
        analisis['desactivos_con_internet'],
        ANALISIS_NAVEGACION['desactivos_con_internet']
    )


def procesar_navegacion_activos_sin_catv(abonados_file, olt_file):
    df_resultado = preparar_resultado_navegacion(abonados_file, olt_file)
    analisis = construir_analisis_navegacion(df_resultado)
    return crear_resultado_analisis(
        analisis['activos_sin_catv'],
        ANALISIS_NAVEGACION['activos_sin_catv']
    )


def procesar_navegacion_solo_con_arroba_y_catv_activo(abonados_file, olt_file):
    df_resultado = preparar_resultado_navegacion(abonados_file, olt_file)
    analisis = construir_analisis_navegacion(df_resultado)
    return crear_resultado_analisis(
        analisis['solo_con_arroba_y_catv_activo'],
        ANALISIS_NAVEGACION['solo_con_arroba_y_catv_activo']
    )
