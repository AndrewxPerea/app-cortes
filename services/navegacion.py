import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import excel_desde_hojas, validar_columnas


def procesar_sin_navegar(abonados_file, olt_file):
    df_abonados = procesar_archivo_excel_solo(abonados_file)
    df_olt = procesar_archivo_csv_solo(olt_file)

    resultado = pd.merge(
        df_abonados, df_olt, how='right', left_on='EQUIPO MACO', right_on='NSN', suffixes=('_abonados', '_cortes')
    )
    df_resultado = resultado.dropna(subset=['EQUIPO MACO'])
    df_resultado.columns = resultado.columns.str.lower()

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
    df_resultado = df_resultado[columnas_deseadas]

    df_resultado1 = df_resultado[
        (df_resultado['estatus'] == 'ACTIVO') &
        (
            (df_resultado['administrative status'] == 'Disabled') |
            ~(df_resultado['status'] == 'Online')
        )
    ]
    df_resultado2 = df_resultado[
        ((df_resultado['estatus'].str.lower().isin(['activo', 'por instalar']) == False) &
         (df_resultado['status'].str.lower() == 'online'))
    ]
    df_resultado3 = df_resultado[
        (~df_resultado['detalle suscripcion'].str.contains('@', na=False)) &
        (df_resultado['estatus'] == 'ACTIVO') &
        (df_resultado['catv'] == 'Disabled')
    ]
    df_resultado4 = df_resultado[
        (df_resultado['detalle suscripcion'].str.contains('@', na=False)) &
        (df_resultado['catv'] == 'Enabled')
    ]

    return {
        'data': df_resultado1,
        'columns': df_resultado.columns,
        'num_casos': df_resultado1.shape[0] + df_resultado2.shape[0] + df_resultado3.shape[0] + df_resultado4.shape[0],
        'excel': excel_desde_hojas(
            [
                ('Activos sin navegar', df_resultado1),
                ('Desactivos con internet', df_resultado2),
                ('Activos sin Catv', df_resultado3),
                ('Solo con @ y catv activo', df_resultado4),
            ]
        ),
    }
