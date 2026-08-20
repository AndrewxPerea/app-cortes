import pandas as pd

from services.common import save_excel_sheets, validar_columnas


ESTADOS_PUERTO = [
    ('online', 'Online'),
    ('log in', 'Log in'),
    ('los', 'LOS'),
    ('offline', 'Offline'),
    ('power fail', 'Power fail'),
    ('sync mib', 'Sync Mib'),
]
COLUMNAS_ESTADO_PUERTO = [columna for columna, _ in ESTADOS_PUERTO]
ORDEN_PRIORIDAD_PUERTO = {
    'ALTA': 0,
    'MEDIA': 1,
    'BAJA': 2,
    'NORMAL': 3,
}


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


def _normalizar_estado(valor):
    texto = str(valor).strip().lower()
    texto = texto.replace('_', ' ').replace('-', ' ')
    return ' '.join(texto.split())


def _contar_estado(serie, estado):
    estado_normalizado = _normalizar_estado(estado)
    return int(_serie_texto(serie).apply(_normalizar_estado).eq(estado_normalizado).sum())


def _contar_administrative_status(serie, estado):
    estado_normalizado = _normalizar_estado(estado)
    valores = _serie_texto(serie).apply(_normalizar_estado)
    return int(valores.str.startswith(estado_normalizado).sum())


def _administrative_status_es(serie, estado):
    estado_normalizado = _normalizar_estado(estado)
    return _serie_texto(serie).apply(_normalizar_estado).str.startswith(estado_normalizado)


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


def _referencia_address(serie):
    valores = _serie_texto(serie)
    valores = valores[valores.ne('') & valores.ne('Sin dato')]
    if valores.empty:
        return ''

    moda = valores.mode()
    if not moda.empty:
        return moda.iloc[0]
    return valores.iloc[0]


def _porcentaje(numerador, denominador):
    if not denominador:
        return 0.0
    return round((numerador / denominador) * 100, 2)


def _clasificar_danio_dominante(fila):
    grupos = {
        'ENERGIA': int(fila.get('power fail', 0)),
        'FIBRA': int(fila.get('los', 0)),
        'SINCRONIZACION': int(fila.get('log in', 0)) + int(fila.get('sync mib', 0)),
        'OFFLINE': int(fila.get('offline', 0)),
        'OTROS': int(fila.get('otros estados', 0)),
    }
    activos = {nombre: cantidad for nombre, cantidad in grupos.items() if cantidad > 0}
    if not activos:
        return 'SIN_ALERTA'

    mayor = max(activos.values())
    dominantes = [nombre for nombre, cantidad in activos.items() if cantidad == mayor]
    if len(dominantes) > 1:
        return 'MIXTO'
    return dominantes[0]


def _clasificar_estado_arpon(fila):
    total = int(fila.get('cantidad usuarios', 0))
    online = int(fila.get('online', 0))
    afectados = int(fila.get('usuarios afectados', 0))
    porcentaje_afectados = float(fila.get('porcentaje afectados', 0))

    if total == 0 or afectados == 0:
        return 'NORMAL'
    if online == 0:
        return 'CAIDO'
    if porcentaje_afectados >= 80:
        return 'CRITICO'
    if porcentaje_afectados >= 50:
        return 'ALARMADO'
    return 'EN OBSERVACION'


def _clasificar_prioridad_puerto(fila):
    afectados = int(fila.get('usuarios afectados', 0))
    porcentaje_afectados = float(fila.get('porcentaje afectados', 0))
    estado_arpon = fila.get('estado arpon', 'NORMAL')

    if estado_arpon == 'CAIDO' or porcentaje_afectados >= 80 or afectados >= 8:
        return 'ALTA'
    if porcentaje_afectados >= 50 or afectados >= 4:
        return 'MEDIA'
    if afectados > 0:
        return 'BAJA'
    return 'NORMAL'


def _construir_resumen_board_port(df):
    trabajo = df.copy()
    trabajo['__admin_enable'] = _administrative_status_es(trabajo['administrative status'], 'enable')
    trabajo['__admin_disable'] = _administrative_status_es(trabajo['administrative status'], 'disable')
    trabajo['__status_norm'] = _serie_texto(trabajo['status']).apply(_normalizar_estado)
    trabajo['__signal_1310_contable'] = trabajo['signal 1310'].where(trabajo['__admin_enable'])
    trabajo['__signal_1490_contable'] = trabajo['signal 1490'].where(trabajo['__admin_enable'])

    agregaciones = {
        'cantidad usuarios': ('__admin_enable', 'sum'),
        'administrative status enable': ('__admin_enable', 'sum'),
        'administrative status disable': ('__admin_disable', 'sum'),
        'referencia direccion': ('address', _referencia_address),
        'promedio signal 1310': ('__signal_1310_contable', 'mean'),
        'promedio signal 1490': ('__signal_1490_contable', 'mean'),
    }
    for columna, estado in ESTADOS_PUERTO:
        columna_interna = f'__estado_{columna}'
        trabajo[columna_interna] = (
            trabajo['__admin_enable'] &
            trabajo['__status_norm'].eq(_normalizar_estado(estado))
        ).astype(int)
        agregaciones[columna] = (columna_interna, 'sum')

    board_port = (
        trabajo.groupby(['olt', 'board', 'port'], dropna=False)
        .agg(**agregaciones)
        .reset_index()
        .sort_values(['olt', 'board', 'port'])
        .reset_index(drop=True)
    )

    board_port['estados contabilizados'] = board_port[COLUMNAS_ESTADO_PUERTO].sum(axis=1)
    board_port['otros estados'] = (
        board_port['cantidad usuarios'] - board_port['estados contabilizados']
    ).clip(lower=0)
    board_port['usuarios afectados'] = (
        board_port['cantidad usuarios'] - board_port['online']
    ).clip(lower=0)
    board_port['porcentaje online'] = board_port.apply(
        lambda fila: _porcentaje(fila['online'], fila['cantidad usuarios']),
        axis=1,
    )
    board_port['porcentaje afectados'] = board_port.apply(
        lambda fila: _porcentaje(fila['usuarios afectados'], fila['cantidad usuarios']),
        axis=1,
    )
    board_port['danio dominante'] = board_port.apply(_clasificar_danio_dominante, axis=1)
    board_port['estado arpon'] = board_port.apply(_clasificar_estado_arpon, axis=1)
    board_port['prioridad'] = board_port.apply(_clasificar_prioridad_puerto, axis=1)

    columnas = [
        'olt', 'board', 'port', 'referencia direccion', 'prioridad',
        'estado arpon', 'danio dominante', 'cantidad usuarios',
        'administrative status enable',
        'administrative status disable', 'online', 'log in', 'los', 'offline',
        'power fail', 'sync mib', 'otros estados', 'estados contabilizados',
        'usuarios afectados', 'porcentaje online', 'porcentaje afectados',
        'promedio signal 1310', 'promedio signal 1490',
    ]
    board_port = board_port[columnas].copy()
    return _redondear_columnas(
        board_port,
        [
            'porcentaje online',
            'porcentaje afectados',
            'promedio signal 1310',
            'promedio signal 1490',
        ],
    )


def _construir_puertos_priorizados(board_port):
    priorizados = board_port[board_port['usuarios afectados'] > 0].copy()
    if priorizados.empty:
        return priorizados

    priorizados['orden prioridad'] = priorizados['prioridad'].map(ORDEN_PRIORIDAD_PUERTO).fillna(99)
    priorizados = (
        priorizados
        .sort_values(
            ['orden prioridad', 'usuarios afectados', 'porcentaje afectados', 'cantidad usuarios'],
            ascending=[True, False, False, False],
        )
        .drop(columns=['orden prioridad'])
        .reset_index(drop=True)
    )
    return priorizados


def job_estadisticos_olt(df_olt):
    df = normalize_cols(df_olt)
    columnas_requeridas = [
        'sn', 'olt', 'board', 'port', 'status', 'signal',
        'signal 1310', 'signal 1490', 'catv', 'administrative status',
    ]
    validar_columnas(df, columnas_requeridas, 'CSV SmartOLT')

    df = df.copy()
    if 'address' not in df.columns:
        df['address'] = ''

    df['signal 1310'] = pd.to_numeric(df['signal 1310'], errors='coerce')
    df['signal 1490'] = pd.to_numeric(df['signal 1490'], errors='coerce')

    for columna in ['olt', 'board', 'port', 'status', 'signal', 'catv', 'administrative status', 'address']:
        df[columna] = _serie_categoria(df[columna])

    df['__admin_enable'] = _administrative_status_es(df['administrative status'], 'enable')
    df_contable = df[df['__admin_enable']].copy()

    resumen_olt = (
        df_contable.groupby('olt', dropna=False)
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

    status_olt = _conteo_por_columnas(df_contable, ['olt', 'status'])
    signal_olt = _conteo_por_columnas(df_contable, ['olt', 'signal'])

    board_port = _construir_resumen_board_port(df)
    puertos_priorizados = _construir_puertos_priorizados(board_port)

    catv = _conteo_por_columnas(df_contable, ['olt', 'catv'])
    administrative_status = _conteo_por_columnas(df, ['olt', 'administrative status'])

    olt_validas = _serie_texto(df_contable['olt'])
    olt_validas = olt_validas[olt_validas.ne('') & olt_validas.ne('Sin dato')]
    summary = {
        'Total abonados': int(df_contable.shape[0]),
        'Total OLT': int(olt_validas.nunique()),
        'Total puertos': int((board_port['cantidad usuarios'] > 0).sum()),
        'Puertos afectados': int((board_port['usuarios afectados'] > 0).sum()),
        'Puertos caidos': int((board_port['estado arpon'] == 'CAIDO').sum()),
        'Administrative enable': _contar_administrative_status(df['administrative status'], 'enable'),
        'Administrative disable': _contar_administrative_status(df['administrative status'], 'disable'),
        'Online': _contar_estado(df_contable['status'], 'online'),
        'Log in': _contar_estado(df_contable['status'], 'log in'),
        'LOS': _contar_estado(df_contable['status'], 'los'),
        'Offline': _contar_estado(df_contable['status'], 'offline'),
        'Power fail': _contar_estado(df_contable['status'], 'power fail'),
        'Sync Mib': _contar_estado(df_contable['status'], 'sync mib'),
        'Warning': _contar_categoria(df_contable['signal'], 'warning'),
        'Critical': _contar_categoria(df_contable['signal'], 'critical'),
    }

    sheets = [
        ('Resumen por OLT', resumen_olt),
        ('Status por OLT', status_olt),
        ('Signal por OLT', signal_olt),
        ('Board Port', board_port),
        ('Puertos Priorizados', puertos_priorizados),
        ('CATV', catv),
        ('Administrative Status', administrative_status),
    ]

    preview_df = puertos_priorizados if not puertos_priorizados.empty else board_port
    return sheets, preview_df, summary


def procesar_estadisticos_olt(olt_file):
    df_olt = normalize_cols(read_csv(olt_file))
    sheets, preview_df, summary = job_estadisticos_olt(df_olt)

    return {
        'data': preview_df,
        'columns': list(preview_df.columns),
        'num_casos': summary['Puertos afectados'],
        'summary': summary,
        'excel': save_excel_sheets(sheets),
    }
