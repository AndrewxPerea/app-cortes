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
ORDEN_RESULTADO_CAMBIO = {
    'EMPEORO': 0,
    'NUEVO': 1,
    'RETIRADO': 2,
    'CAMBIO OPERATIVO': 3,
    'MEJORO': 4,
    'SIN CAMBIOS': 5,
}
COLUMNAS_COMPARATIVO_TEXTO = [
    'referencia direccion',
    'prioridad',
    'estado arpon',
    'danio dominante',
]
COLUMNAS_COMPARATIVO_NUMERICAS = [
    'cantidad usuarios',
    'administrative status enable',
    'administrative status disable',
    'online',
    'log in',
    'los',
    'offline',
    'power fail',
    'sync mib',
    'otros estados',
    'estados contabilizados',
    'usuarios afectados',
    'porcentaje online',
    'porcentaje afectados',
    'promedio signal 1310',
    'promedio signal 1490',
]
COLUMNAS_DETALLE_CAMBIO = [
    'prioridad',
    'estado arpon',
    'danio dominante',
    'cantidad usuarios',
    'usuarios afectados',
    'porcentaje afectados',
    'online',
    'log in',
    'los',
    'offline',
    'power fail',
    'sync mib',
    'promedio signal 1310',
    'promedio signal 1490',
]


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


def _normalizar_clave_comparativo(valor):
    texto = str(valor).strip()
    if texto.lower() in {'', 'nan', 'none', 'nat', '<na>'}:
        return 'sin dato'

    texto = ' '.join(texto.split())
    if texto.endswith('.0'):
        posible_entero = texto[:-2]
        if posible_entero.replace('-', '', 1).isdigit():
            texto = posible_entero

    return texto.lower()


def _texto_comparativo(valor):
    if pd.isna(valor):
        return ''
    return ' '.join(str(valor).strip().lower().split())


def _numero_comparativo(valor):
    if pd.isna(valor):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _formatear_valor_comparativo(valor):
    if pd.isna(valor):
        return ''

    numero = _numero_comparativo(valor)
    if numero is not None:
        if numero.is_integer():
            return str(int(numero))
        return str(round(numero, 2))

    return str(valor)


def _valor_cambio(fila, columna):
    antiguo = fila.get(f'{columna}_antiguo')
    nuevo = fila.get(f'{columna}_nuevo')

    if columna in COLUMNAS_COMPARATIVO_NUMERICAS:
        numero_antiguo = _numero_comparativo(antiguo)
        numero_nuevo = _numero_comparativo(nuevo)
        if numero_antiguo is None and numero_nuevo is None:
            return False
        if numero_antiguo is None or numero_nuevo is None:
            return True
        return abs(numero_nuevo - numero_antiguo) > 0.009

    return _texto_comparativo(antiguo) != _texto_comparativo(nuevo)


def _detalle_cambios_board_port(fila):
    estado = fila.get('estado comparativo')
    if estado == 'NUEVO':
        return 'Puerto presente solo en el archivo nuevo.'
    if estado == 'RETIRADO':
        return 'Puerto presente solo en el archivo antiguo.'
    if estado == 'SIN CAMBIOS':
        return 'Sin cambios relevantes.'

    cambios = []
    for columna in COLUMNAS_DETALLE_CAMBIO:
        if _valor_cambio(fila, columna):
            antiguo = _formatear_valor_comparativo(fila.get(f'{columna}_antiguo'))
            nuevo = _formatear_valor_comparativo(fila.get(f'{columna}_nuevo'))
            cambios.append(f'{columna}: {antiguo} -> {nuevo}')

    return '; '.join(cambios) or 'Cambio operativo sin variacion numerica principal.'


def _clasificar_resultado_comparativo(fila):
    estado = fila.get('estado comparativo')
    if estado in {'NUEVO', 'RETIRADO', 'SIN CAMBIOS'}:
        return estado

    delta_afectados = _numero_comparativo(fila.get('delta usuarios afectados')) or 0
    if delta_afectados > 0:
        return 'EMPEORO'
    if delta_afectados < 0:
        return 'MEJORO'

    prioridad_antigua = fila.get('prioridad_antiguo')
    prioridad_nueva = fila.get('prioridad_nuevo')
    orden_antiguo = ORDEN_PRIORIDAD_PUERTO.get(prioridad_antigua, 99)
    orden_nuevo = ORDEN_PRIORIDAD_PUERTO.get(prioridad_nueva, 99)
    if orden_nuevo < orden_antiguo:
        return 'EMPEORO'
    if orden_nuevo > orden_antiguo:
        return 'MEJORO'

    return 'CAMBIO OPERATIVO'


def _preparar_board_port_comparativo(board_port):
    trabajo = board_port.copy()
    for columna in ['olt', 'board', 'port']:
        trabajo[f'__clave_{columna}'] = trabajo[columna].apply(_normalizar_clave_comparativo)
    return trabajo


def _renombrar_columna_comparativo(columna):
    if columna.endswith('_antiguo'):
        return f"{columna[:-8]} antiguo"
    if columna.endswith('_nuevo'):
        return f"{columna[:-6]} nuevo"
    return columna


def _seleccionar_columnas_comparativo(comparativo):
    columnas_salida = [
        'olt',
        'board',
        'port',
        'estado comparativo',
        'resultado cambio',
        'detalle cambios',
        'referencia direccion_antiguo',
        'referencia direccion_nuevo',
        'prioridad_antiguo',
        'prioridad_nuevo',
        'estado arpon_antiguo',
        'estado arpon_nuevo',
        'danio dominante_antiguo',
        'danio dominante_nuevo',
        'cantidad usuarios_antiguo',
        'cantidad usuarios_nuevo',
        'delta cantidad usuarios',
        'usuarios afectados_antiguo',
        'usuarios afectados_nuevo',
        'delta usuarios afectados',
        'porcentaje afectados_antiguo',
        'porcentaje afectados_nuevo',
        'delta porcentaje afectados',
        'porcentaje online_antiguo',
        'porcentaje online_nuevo',
        'delta porcentaje online',
        'online_antiguo',
        'online_nuevo',
        'delta online',
        'log in_antiguo',
        'log in_nuevo',
        'delta log in',
        'los_antiguo',
        'los_nuevo',
        'delta los',
        'offline_antiguo',
        'offline_nuevo',
        'delta offline',
        'power fail_antiguo',
        'power fail_nuevo',
        'delta power fail',
        'sync mib_antiguo',
        'sync mib_nuevo',
        'delta sync mib',
        'otros estados_antiguo',
        'otros estados_nuevo',
        'delta otros estados',
        'administrative status enable_antiguo',
        'administrative status enable_nuevo',
        'delta administrative status enable',
        'administrative status disable_antiguo',
        'administrative status disable_nuevo',
        'delta administrative status disable',
        'promedio signal 1310_antiguo',
        'promedio signal 1310_nuevo',
        'delta promedio signal 1310',
        'promedio signal 1490_antiguo',
        'promedio signal 1490_nuevo',
        'delta promedio signal 1490',
    ]
    columnas_salida = [columna for columna in columnas_salida if columna in comparativo.columns]
    salida = comparativo[columnas_salida].copy()
    salida = salida.rename(columns=_renombrar_columna_comparativo)
    return _redondear_columnas(
        salida,
        [columna for columna in salida.columns if columna.startswith('delta ') or 'porcentaje' in columna or 'promedio signal' in columna],
    )


def _construir_comparativo_board_port(board_port_antiguo, board_port_nuevo):
    antiguo = _preparar_board_port_comparativo(board_port_antiguo)
    nuevo = _preparar_board_port_comparativo(board_port_nuevo)

    comparativo = antiguo.merge(
        nuevo,
        on=['__clave_olt', '__clave_board', '__clave_port'],
        how='outer',
        suffixes=('_antiguo', '_nuevo'),
        indicator=True,
    )

    for columna in ['olt', 'board', 'port']:
        comparativo[columna] = (
            comparativo[f'{columna}_nuevo']
            .combine_first(comparativo[f'{columna}_antiguo'])
            .fillna('Sin dato')
        )

    ambas_versiones = comparativo['_merge'].eq('both')
    columnas_con_cambio = pd.Series(False, index=comparativo.index)

    for columna in COLUMNAS_COMPARATIVO_NUMERICAS:
        columna_antigua = f'{columna}_antiguo'
        columna_nueva = f'{columna}_nuevo'
        if columna_antigua not in comparativo.columns or columna_nueva not in comparativo.columns:
            continue

        valores_antiguos = pd.to_numeric(comparativo[columna_antigua], errors='coerce')
        valores_nuevos = pd.to_numeric(comparativo[columna_nueva], errors='coerce')
        comparativo[f'delta {columna}'] = (valores_nuevos.fillna(0) - valores_antiguos.fillna(0)).round(2)
        cambio_columna = (
            valores_nuevos.sub(valores_antiguos).abs().gt(0.009) &
            ~(valores_antiguos.isna() & valores_nuevos.isna())
        )
        columnas_con_cambio = columnas_con_cambio | (ambas_versiones & cambio_columna)

    for columna in COLUMNAS_COMPARATIVO_TEXTO:
        columna_antigua = f'{columna}_antiguo'
        columna_nueva = f'{columna}_nuevo'
        if columna_antigua not in comparativo.columns or columna_nueva not in comparativo.columns:
            continue

        valores_antiguos = comparativo[columna_antigua].apply(_texto_comparativo)
        valores_nuevos = comparativo[columna_nueva].apply(_texto_comparativo)
        columnas_con_cambio = columnas_con_cambio | (ambas_versiones & valores_antiguos.ne(valores_nuevos))

    comparativo['estado comparativo'] = 'SIN CAMBIOS'
    comparativo.loc[comparativo['_merge'].eq('left_only'), 'estado comparativo'] = 'RETIRADO'
    comparativo.loc[comparativo['_merge'].eq('right_only'), 'estado comparativo'] = 'NUEVO'
    comparativo.loc[ambas_versiones & columnas_con_cambio, 'estado comparativo'] = 'CON CAMBIOS'
    comparativo['resultado cambio'] = comparativo.apply(_clasificar_resultado_comparativo, axis=1)
    comparativo['detalle cambios'] = comparativo.apply(_detalle_cambios_board_port, axis=1)
    comparativo['__orden_resultado'] = (
        comparativo['resultado cambio']
        .map(ORDEN_RESULTADO_CAMBIO)
        .fillna(99)
    )

    return comparativo


def _construir_resumen_comparativo_olt(comparativo):
    trabajo = comparativo.copy()
    for columna in ['cantidad usuarios', 'usuarios afectados', 'online', 'los', 'offline', 'power fail']:
        trabajo[f'__{columna}_antiguo'] = pd.to_numeric(
            trabajo.get(f'{columna}_antiguo', 0),
            errors='coerce',
        ).fillna(0)
        trabajo[f'__{columna}_nuevo'] = pd.to_numeric(
            trabajo.get(f'{columna}_nuevo', 0),
            errors='coerce',
        ).fillna(0)

    trabajo['__sin_cambios'] = trabajo['estado comparativo'].eq('SIN CAMBIOS')
    trabajo['__con_cambios'] = trabajo['estado comparativo'].ne('SIN CAMBIOS')
    trabajo['__nuevo'] = trabajo['estado comparativo'].eq('NUEVO')
    trabajo['__retirado'] = trabajo['estado comparativo'].eq('RETIRADO')
    trabajo['__empeoro'] = trabajo['resultado cambio'].eq('EMPEORO')
    trabajo['__mejoro'] = trabajo['resultado cambio'].eq('MEJORO')
    trabajo['__afectado_antiguo'] = trabajo['__usuarios afectados_antiguo'].gt(0)
    trabajo['__afectado_nuevo'] = trabajo['__usuarios afectados_nuevo'].gt(0)

    resumen = (
        trabajo.groupby('olt', dropna=False)
        .agg(**{
            'puertos comparados': ('estado comparativo', 'size'),
            'puertos sin cambios': ('__sin_cambios', 'sum'),
            'puertos con cambios': ('__con_cambios', 'sum'),
            'puertos nuevos': ('__nuevo', 'sum'),
            'puertos retirados': ('__retirado', 'sum'),
            'puertos que empeoraron': ('__empeoro', 'sum'),
            'puertos que mejoraron': ('__mejoro', 'sum'),
            'puertos afectados antiguo': ('__afectado_antiguo', 'sum'),
            'puertos afectados nuevo': ('__afectado_nuevo', 'sum'),
            'usuarios antiguo': ('__cantidad usuarios_antiguo', 'sum'),
            'usuarios nuevo': ('__cantidad usuarios_nuevo', 'sum'),
            'usuarios afectados antiguo': ('__usuarios afectados_antiguo', 'sum'),
            'usuarios afectados nuevo': ('__usuarios afectados_nuevo', 'sum'),
            'online antiguo': ('__online_antiguo', 'sum'),
            'online nuevo': ('__online_nuevo', 'sum'),
            'los antiguo': ('__los_antiguo', 'sum'),
            'los nuevo': ('__los_nuevo', 'sum'),
            'offline antiguo': ('__offline_antiguo', 'sum'),
            'offline nuevo': ('__offline_nuevo', 'sum'),
            'power fail antiguo': ('__power fail_antiguo', 'sum'),
            'power fail nuevo': ('__power fail_nuevo', 'sum'),
        })
        .reset_index()
        .sort_values('olt')
        .reset_index(drop=True)
    )

    for columna in ['usuarios', 'usuarios afectados', 'online', 'los', 'offline', 'power fail']:
        resumen[f'delta {columna}'] = resumen[f'{columna} nuevo'] - resumen[f'{columna} antiguo']

    columnas = [
        'olt',
        'puertos comparados',
        'puertos sin cambios',
        'puertos con cambios',
        'puertos nuevos',
        'puertos retirados',
        'puertos que empeoraron',
        'puertos que mejoraron',
        'puertos afectados antiguo',
        'puertos afectados nuevo',
        'usuarios antiguo',
        'usuarios nuevo',
        'delta usuarios',
        'usuarios afectados antiguo',
        'usuarios afectados nuevo',
        'delta usuarios afectados',
        'online antiguo',
        'online nuevo',
        'delta online',
        'los antiguo',
        'los nuevo',
        'delta los',
        'offline antiguo',
        'offline nuevo',
        'delta offline',
        'power fail antiguo',
        'power fail nuevo',
        'delta power fail',
    ]
    return resumen[columnas]


def _nombre_hoja_estadistico(nombre, prefijo):
    return f'{prefijo} {nombre}'[:31]


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


def procesar_comparativo_estadisticos_olt(olt_antiguo_file, olt_nuevo_file):
    df_antiguo = normalize_cols(read_csv(olt_antiguo_file))
    sheets_antiguo, _, summary_antiguo = job_estadisticos_olt(df_antiguo)
    board_port_antiguo = dict(sheets_antiguo)['Board Port']

    df_nuevo = normalize_cols(read_csv(olt_nuevo_file))
    sheets_nuevo, _, summary_nuevo = job_estadisticos_olt(df_nuevo)
    board_port_nuevo = dict(sheets_nuevo)['Board Port']

    comparativo = _construir_comparativo_board_port(board_port_antiguo, board_port_nuevo)
    resumen_comparativo = _construir_resumen_comparativo_olt(comparativo)

    comparativo_ordenado = comparativo.sort_values(['olt', 'board', 'port']).reset_index(drop=True)
    comparativo_board_port = _seleccionar_columnas_comparativo(comparativo_ordenado)

    cambios = (
        comparativo[comparativo['estado comparativo'].ne('SIN CAMBIOS')]
        .sort_values(['__orden_resultado', 'olt', 'board', 'port'])
        .reset_index(drop=True)
    )
    cambios_board_port = _seleccionar_columnas_comparativo(cambios)

    puertos_nuevos = _seleccionar_columnas_comparativo(
        comparativo[comparativo['estado comparativo'].eq('NUEVO')]
        .sort_values(['olt', 'board', 'port'])
        .reset_index(drop=True)
    )
    puertos_retirados = _seleccionar_columnas_comparativo(
        comparativo[comparativo['estado comparativo'].eq('RETIRADO')]
        .sort_values(['olt', 'board', 'port'])
        .reset_index(drop=True)
    )
    cambios_prioridad = _seleccionar_columnas_comparativo(
        comparativo[
            comparativo['_merge'].eq('both') &
            comparativo['prioridad_antiguo'].apply(_texto_comparativo).ne(
                comparativo['prioridad_nuevo'].apply(_texto_comparativo)
            )
        ]
        .sort_values(['__orden_resultado', 'olt', 'board', 'port'])
        .reset_index(drop=True)
    )

    total_cambios = int(comparativo['estado comparativo'].ne('SIN CAMBIOS').sum())
    total_nuevos = int(comparativo['estado comparativo'].eq('NUEVO').sum())
    total_retirados = int(comparativo['estado comparativo'].eq('RETIRADO').sum())
    total_empeoro = int(comparativo['resultado cambio'].eq('EMPEORO').sum())
    total_mejoro = int(comparativo['resultado cambio'].eq('MEJORO').sum())
    delta_afectados = int(pd.to_numeric(comparativo['delta usuarios afectados'], errors='coerce').fillna(0).sum())

    summary = {
        'Puertos comparados': int(comparativo.shape[0]),
        'Puertos con cambios': total_cambios,
        'Puertos nuevos': total_nuevos,
        'Puertos retirados': total_retirados,
        'Empeoraron': total_empeoro,
        'Mejoraron': total_mejoro,
        'Delta afectados': delta_afectados,
        'Abonados nuevo': summary_nuevo['Total abonados'],
        'Abonados antiguo': summary_antiguo['Total abonados'],
    }

    sheets = [
        ('Resumen Comparativo', resumen_comparativo),
        ('Cambios Board Port', cambios_board_port),
        ('Comparativo Board Port', comparativo_board_port),
        ('Puertos Nuevos', puertos_nuevos),
        ('Puertos Retirados', puertos_retirados),
        ('Cambios Prioridad', cambios_prioridad),
    ]
    sheets.extend(
        (_nombre_hoja_estadistico(nombre, 'Ant'), df)
        for nombre, df in sheets_antiguo
    )
    sheets.extend(
        (_nombre_hoja_estadistico(nombre, 'Nvo'), df)
        for nombre, df in sheets_nuevo
    )

    preview_df = cambios_board_port if not cambios_board_port.empty else comparativo_board_port
    return {
        'data': preview_df,
        'columns': list(preview_df.columns),
        'num_casos': total_cambios,
        'summary': summary,
        'excel': save_excel_sheets(sheets),
    }
