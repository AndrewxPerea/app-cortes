import pandas as pd

from funciones import (
    leer_excel_seguro,
    normalizar_columnas,
    procesar_archivo_csv_solo,
    procesar_archivo_excel_solo,
)
from services.common import excel_desde_hojas, validar_columnas


def tomar_columna(df, *candidatas):
    for columna in candidatas:
        if columna in df.columns:
            return df[columna]
    raise KeyError(f"No se encontró ninguna de las columnas esperadas: {', '.join(candidatas)}")


def procesar_auditoria_reconexiones(drive_file, saeplus_file, epayco_file, olt_file):
    df_drive = leer_excel_seguro(drive_file)
    df_saeplus = leer_excel_seguro(saeplus_file)
    df_epayco = leer_excel_seguro(epayco_file)
    saeplus_file.seek(0)
    df_abonado_cortes = procesar_archivo_excel_solo(saeplus_file)
    df_olt_cortes = procesar_archivo_csv_solo(olt_file)

    df_drive = normalizar_columnas(df_drive, 'abonados')
    df_saeplus = normalizar_columnas(df_saeplus, 'abonados')
    df_epayco = normalizar_columnas(df_epayco, 'abonados')
    validar_columnas(df_drive, ['abonados'], 'drive')
    validar_columnas(df_saeplus, ['abonados'], 'saeplus')
    validar_columnas(df_epayco, ['abonados'], 'epayco')

    df_resultado1 = pd.merge(df_drive, df_saeplus, on="abonados", how="inner")
    df_resultado1.columns = df_resultado1.columns.str.lower()
    df_resultado2 = pd.merge(df_drive, df_epayco, on="abonados", how="inner")
    df_resultado2.columns = df_resultado2.columns.str.lower()
    df_resultado3 = pd.merge(
        df_abonado_cortes,
        df_olt_cortes,
        how='right',
        left_on='EQUIPO MACO',
        right_on='NSN',
        suffixes=('_abonados', '_cortes')
    ).dropna(subset=['EQUIPO MACO'])
    df_resultado3.columns = df_resultado3.columns.str.lower()

    df_resultado1 = pd.DataFrame({
        'abonados': tomar_columna(df_resultado1, 'abonados'),
        'documento_x': tomar_columna(df_resultado1, 'documento_x', 'documento'),
        'nombre_x': tomar_columna(df_resultado1, 'nombre_x', 'nombre'),
        'apellido_x': tomar_columna(df_resultado1, 'apellido_x', 'apellido'),
        'observaciones': tomar_columna(df_resultado1, 'observaciones'),
        'estatus_y': tomar_columna(df_resultado1, 'estatus_y', 'estatus'),
        'detalle suscripcion_x': tomar_columna(df_resultado1, 'detalle suscripcion_x', 'detalle suscripcion'),
        'saldo_y': tomar_columna(df_resultado1, 'saldo_y', 'saldo'),
    })
    df_resultado2 = pd.DataFrame({
        'abonados': tomar_columna(df_resultado2, 'abonados'),
        'documento_x': tomar_columna(df_resultado2, 'documento_x', 'documento'),
        'nombre': tomar_columna(df_resultado2, 'nombre', 'nombre_y', 'nombre_x'),
        'apellido': tomar_columna(df_resultado2, 'apellido', 'apellido_y', 'apellido_x'),
        'observaciones': tomar_columna(df_resultado2, 'observaciones', 'observaciones_x', 'observaciones_y'),
        'estatus_x': tomar_columna(df_resultado2, 'estatus_x', 'estatus_y', 'estatus'),
        'detalle suscripcion': tomar_columna(df_resultado2, 'detalle suscripcion', 'detalle suscripcion_y', 'detalle suscripcion_x'),
    })
    df_resultado3 = df_resultado3[
        ['n° abonado', 'documento', 'nombre', 'apellido', 'estatus', 'status', 'olt', 'catv', 'administrative status', 'detalle suscripcion']
    ]

    reconexiones = df_resultado1[
        (df_resultado1['observaciones'].isna() | (df_resultado1['observaciones'] == '')) &
        (df_resultado1['estatus_y'] == 'ACTIVO')
    ]
    df_resultado1 = df_resultado1[(df_resultado1['estatus_y'] == 'ACTIVO')]
    abonados_epayco = df_resultado2[
        (df_resultado2['observaciones'].isna() | (df_resultado2['observaciones'] == ''))
    ]

    df_resultado3 = df_resultado3[
        (df_resultado3['estatus'].str.lower().isin(['activo'])) &
        (
            (df_resultado3['administrative status'].str.lower() != 'enabled') |
            (df_resultado3['catv'].str.lower() != 'enabled')
        ) |
        (df_resultado3['status'].str.lower() != 'online')
    ]

    df_resultado3 = df_resultado3.rename(columns={df_resultado3.columns[0]: 'abonados'})
    df_resultado3 = pd.merge(df_resultado3, df_resultado1, on="abonados", how="inner")

    desactivado = df_resultado3[
        (
            (df_resultado3['status'].str.lower() != 'online') |
            (
                (df_resultado3['detalle suscripcion'].str.contains('@', na=False)) &
                ((df_resultado3['catv'].str.lower() == 'enabled')) |
                (df_resultado3['administrative status'].str.lower() != 'enabled')
            ) |
            (
                ~df_resultado3['detalle suscripcion'].str.contains('@', na=False) &
                ((df_resultado3['catv'].str.lower() != 'enabled')) |
                (df_resultado3['administrative status'].str.lower() != 'enabled')
            )
        )
    ]
    pagos_saeplus = pd.merge(reconexiones, abonados_epayco, on="abonados", how="left", indicator=True)
    pagos_saeplus = pagos_saeplus[pagos_saeplus['_merge'] == 'left_only']
    pagos_saeplus = pagos_saeplus[['abonados', 'nombre_x', 'estatus_y', 'detalle suscripcion_x', 'saldo_y']]

    return {
        'data': desactivado,
        'num_casos': desactivado.shape[0],
        'excel': excel_desde_hojas(
            [
                ('Reconexion sin observaciones', reconexiones),
                ('Pagos de epayco', abonados_epayco),
                ('Abonados sin activar', desactivado),
                ('pagos saeplus', pagos_saeplus),
            ]
        ),
    }
