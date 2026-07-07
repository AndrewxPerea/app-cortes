import pandas as pd

from services.common import save_excel_sheets, validar_columnas


def read_csv(archivo):
    return pd.read_csv(archivo, low_memory=False)


def normalize_cols(df):
    normalizado = df.copy()
    normalizado.columns = normalizado.columns.astype(str).str.strip().str.lower()
    return normalizado


def _serie_texto(serie):
    return serie.fillna('').astype(str).str.strip()


def _serie_categoria(serie):
    return _serie_texto(serie).replace('', 'Sin dato')


def _contar_estado(serie, estado):
    return int(_serie_texto(serie).str.lower().eq(estado).sum())


def _contar_categoria(serie, categoria):
    return int(_serie_texto(serie).str.lower().str.contains(categoria, regex=False).sum())


def _redondear_columnas(df, columnas):
    resultado = df.copy()
    for columna in columnas:
        if columna in resultado.columns:
            resultado[columna] = resultado[columna].round(2)
    return resultado


def _conteo_por_columnas(df, columnas):
    trabajo = df.copy()
    for columna in columnas:
        trabajo[columna] = _serie_categoria(trabajo[columna])

    return (
        trabajo.groupby(columnas, dropna=False)
        .size()
        .reset_index(name='cantidad')
        .sort_values(columnas)
        .reset_index(drop=True)
    )


def job_estadisticos_olt(df_olt):
    df = normalize_cols(df_olt)
    columnas_requeridas = [
        'sn', 'olt', 'board', 'port', 'status', 'signal',
        'signal 1310', 'signal 1490', 'catv', 'administrative status',
    ]
    validar_columnas(df, columnas_requeridas, 'CSV SmartOLT')

    df = df.copy()
    df['signal 1310'] = pd.to_numeric(df['signal 1310'], errors='coerce')
    df['signal 1490'] = pd.to_numeric(df['signal 1490'], errors='coerce')

    for columna in ['olt', 'board', 'port', 'status', 'signal', 'catv', 'administrative status']:
        df[columna] = _serie_categoria(df[columna])

    resumen_olt = (
        df.groupby('olt', dropna=False)
        .agg(**{
            'total abonados': ('sn', 'size'),
            'promedio signal 1310': ('signal 1310', 'mean'),
            'promedio signal 1490': ('signal 1490', 'mean'),
            'peor signal 1310': ('signal 1310', 'min'),
            'peor signal 1490': ('signal 1490', 'min'),
        })
        .reset_index()
        .sort_values('olt')
        .reset_index(drop=True)
    )
    resumen_olt = _redondear_columnas(
        resumen_olt,
        [
            'promedio signal 1310',
            'promedio signal 1490',
            'peor signal 1310',
            'peor signal 1490',
        ],
    )

    status_olt = _conteo_por_columnas(df, ['olt', 'status'])
    signal_olt = _conteo_por_columnas(df, ['olt', 'signal'])

    board_port = (
        df.groupby(['olt', 'board', 'port'], dropna=False)
        .agg(**{
            'cantidad usuarios': ('sn', 'size'),
            'promedio signal 1310': ('signal 1310', 'mean'),
            'promedio signal 1490': ('signal 1490', 'mean'),
            'online': ('status', lambda serie: _contar_estado(serie, 'online')),
            'offline': ('status', lambda serie: _contar_estado(serie, 'offline')),
        })
        .reset_index()
        .sort_values(['olt', 'board', 'port'])
        .reset_index(drop=True)
    )
    board_port = _redondear_columnas(
        board_port,
        ['promedio signal 1310', 'promedio signal 1490'],
    )

    catv = _conteo_por_columnas(df, ['olt', 'catv'])
    administrative_status = _conteo_por_columnas(df, ['olt', 'administrative status'])

    olt_validas = _serie_texto(df['olt'])
    olt_validas = olt_validas[olt_validas.ne('') & olt_validas.ne('Sin dato')]
    summary = {
        'Total abonados': int(df.shape[0]),
        'Total OLT': int(olt_validas.nunique()),
        'Online': _contar_estado(df['status'], 'online'),
        'Offline': _contar_estado(df['status'], 'offline'),
        'Warning': _contar_categoria(df['signal'], 'warning'),
        'Critical': _contar_categoria(df['signal'], 'critical'),
    }

    sheets = [
        ('Resumen por OLT', resumen_olt),
        ('Status por OLT', status_olt),
        ('Signal por OLT', signal_olt),
        ('Board Port', board_port),
        ('CATV', catv),
        ('Administrative Status', administrative_status),
    ]

    preview_df = resumen_olt
    return sheets, preview_df, summary


def procesar_estadisticos_olt(olt_file):
    df_olt = normalize_cols(read_csv(olt_file))
    sheets, preview_df, summary = job_estadisticos_olt(df_olt)

    return {
        'data': preview_df,
        'columns': list(preview_df.columns),
        'num_casos': summary['Total abonados'],
        'summary': summary,
        'excel': save_excel_sheets(sheets),
    }
