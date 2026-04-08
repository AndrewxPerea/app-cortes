import re
import unicodedata

import pandas as pd

from funciones import abrir_excel_seguro, procesar_archivo_csv_solo
from services.common import excel_desde_hojas, validar_columnas


ESTADOS_ALERTA = {'offline', 'power fail', 'los'}
COLUMNAS_UBICACION = ['barrio', 'direccion', 'ciudad']
TIPOS_VIA = {
    'AK', 'AUTOPISTA', 'AV', 'AVENIDA', 'BL', 'BLOQUE', 'C', 'CA', 'CALLE', 'CARRERA',
    'CASA', 'CL', 'CRA', 'CR', 'CS', 'DG', 'DIAGONAL', 'KR', 'MANZANA', 'MZ', 'PASAJE',
    'PJ', 'PORTAL', 'TORRE', 'TO', 'TR', 'TRANSV', 'TRANSVERSAL', 'TRV', 'TV'
}
COLUMNAS_RESUMEN = [
    'n° abonado',
    'nombre',
    'estatus saeplus',
    'status smartolt propio',
    'precinto',
    'ciudad',
    'barrio',
    'dirección',
    'referencia dirección',
    'cantidad alertas cercanas',
    'abonados alerta cercanos',
    'estatus smartolt cercanos',
    'criterios detectados',
]
COLUMNAS_DETALLE = [
    'criterio cercanía',
    'ciudad',
    'barrio',
    'referencia dirección',
    'n° abonado sin precinto',
    'nombre sin precinto',
    'estatus saeplus sin precinto',
    'status smartolt sin precinto',
    'precinto saeplus',
    'dirección sin precinto',
    'n° abonado alerta',
    'nombre alerta',
    'estatus saeplus alerta',
    'status smartolt alerta',
    'dirección alerta',
    'olt alerta',
    'sn alerta',
]


def obtener_stream(archivo):
    stream = getattr(archivo, 'stream', archivo)
    if hasattr(stream, 'seek'):
        stream.seek(0)
    return stream


def normalizar_encabezado(valor):
    texto = '' if valor is None else str(valor)
    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(char for char in texto if not unicodedata.combining(char))
    texto = texto.lower().strip()
    texto = texto.replace('º', 'o').replace('°', 'o')
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()

    alias = {
        'no abonado': 'n abonado',
        'n abonado': 'n abonado',
        'nro abonado': 'n abonado',
        'numero abonado': 'n abonado',
        'n o abonado': 'n abonado',
        'direccion': 'direccion',
        'equipo mac': 'equipo mac',
        'equipo maco': 'equipo maco',
    }
    return alias.get(texto, texto)


def normalizar_identificador(valor):
    if pd.isna(valor):
        return pd.NA

    if isinstance(valor, (int, float)):
        if float(valor).is_integer():
            return str(int(valor))
        return str(valor).strip()

    texto = str(valor).strip()
    if not texto:
        return pd.NA

    if re.fullmatch(r'\d+\.0', texto):
        return texto[:-2]

    return texto


def normalizar_equipo(valor):
    texto = normalizar_identificador(valor)
    if pd.isna(texto):
        return pd.NA
    return str(texto).upper()[-8:]


def valor_tiene_contenido(valor):
    if pd.isna(valor):
        return False
    return str(valor).strip() != ''


def serie_tiene_contenido(serie):
    if pd.api.types.is_object_dtype(serie) or pd.api.types.is_string_dtype(serie):
        return serie.notna() & serie.astype(str).str.strip().ne('')
    return serie.notna()


def limpiar_vacios_serie(serie):
    if pd.api.types.is_object_dtype(serie) or pd.api.types.is_string_dtype(serie):
        return serie.where(serie_tiene_contenido(serie), pd.NA)
    return serie


def consolidar_por_llave(df, llave, columnas):
    if df.empty:
        return pd.DataFrame(columns=columnas)

    trabajo = df[columnas].copy()
    trabajo = trabajo[serie_tiene_contenido(trabajo[llave])].copy()
    if trabajo.empty:
        return pd.DataFrame(columns=columnas)

    for columna in columnas:
        if columna == llave:
            continue
        trabajo[columna] = limpiar_vacios_serie(trabajo[columna])

    return (
        trabajo.groupby(llave, as_index=False, dropna=False, sort=False)
        .first()
        .reindex(columns=columnas)
    )


def preparar_hoja_saeplus(df):
    hoja = df.copy()
    hoja.columns = [normalizar_encabezado(columna) for columna in hoja.columns]

    if 'equipo maco' not in hoja.columns and 'equipo mac' in hoja.columns:
        hoja['equipo maco'] = hoja['equipo mac'].apply(normalizar_equipo)
    elif 'equipo maco' in hoja.columns:
        hoja['equipo maco'] = hoja['equipo maco'].apply(normalizar_equipo)

    if 'n abonado' in hoja.columns:
        hoja['n abonado'] = hoja['n abonado'].apply(normalizar_identificador)
    if 'documento' in hoja.columns:
        hoja['documento'] = hoja['documento'].apply(normalizar_identificador)

    base_columnas = ['equipo maco', 'n abonado', 'documento', 'nombre', 'estatus', 'precinto']
    ubicacion_llaves = [col for col in ['n abonado', 'documento', 'equipo maco'] if col in hoja.columns]

    base = None
    if set(['equipo maco', 'n abonado', 'documento', 'nombre', 'estatus', 'precinto']).issubset(hoja.columns):
        base = hoja[base_columnas + [col for col in COLUMNAS_UBICACION if col in hoja.columns]].copy()

    ubicacion = None
    if ubicacion_llaves and any(col in hoja.columns for col in COLUMNAS_UBICACION):
        columnas_ubicacion = [*ubicacion_llaves, *[col for col in COLUMNAS_UBICACION if col in hoja.columns]]
        ubicacion = hoja[columnas_ubicacion].copy()

    return base, ubicacion


def cargar_saeplus_con_ubicacion(archivo_excel, requerir_ubicacion=False):
    excel = abrir_excel_seguro(obtener_stream(archivo_excel))
    bases = []
    ubicaciones = []

    for nombre_hoja in excel.sheet_names:
        df = excel.parse(sheet_name=nombre_hoja)
        base, ubicacion = preparar_hoja_saeplus(df)
        if base is not None:
            bases.append(base)
        if ubicacion is not None:
            ubicaciones.append(ubicacion)

    if not bases:
        raise ValueError(
            "El archivo SAEPlus no contiene una hoja con las columnas requeridas: "
            "EQUIPO MAC, n° abonado, documento, nombre, estatus y precinto."
        )

    base = pd.concat(bases, ignore_index=True, sort=False)
    for columna in COLUMNAS_UBICACION:
        if columna not in base.columns:
            base[columna] = pd.NA

    base = consolidar_por_llave(
        base,
        'equipo maco',
        ['equipo maco', 'n abonado', 'documento', 'nombre', 'estatus', 'precinto', *COLUMNAS_UBICACION]
    )
    validar_columnas(
        base,
        ['equipo maco', 'n abonado', 'documento', 'nombre', 'estatus', 'precinto'],
        'SAEPlus'
    )

    if not ubicaciones:
        if not requerir_ubicacion:
            return base
        raise ValueError(
            "El archivo SAEPlus no contiene una hoja con columnas de ubicación como barrio, dirección o ciudad."
        )

    ubicacion = pd.concat(ubicaciones, ignore_index=True, sort=False)
    for columna in COLUMNAS_UBICACION:
        if columna not in ubicacion.columns:
            ubicacion[columna] = pd.NA

    for llave in ['n abonado', 'documento', 'equipo maco']:
        if llave not in ubicacion.columns or llave not in base.columns:
            continue

        aux = consolidar_por_llave(
            ubicacion,
            llave,
            [llave, *COLUMNAS_UBICACION]
        )
        if aux.empty:
            continue

        renombres = {columna: f'{columna}_ubicacion' for columna in COLUMNAS_UBICACION}
        base = base.merge(aux.rename(columns=renombres), on=llave, how='left')

        for columna in COLUMNAS_UBICACION:
            columna_ubicacion = f'{columna}_ubicacion'
            if columna_ubicacion not in base.columns:
                continue

            mascara_vacia = ~serie_tiene_contenido(base[columna])
            base.loc[mascara_vacia, columna] = base.loc[mascara_vacia, columna_ubicacion]
            base = base.drop(columns=[columna_ubicacion])

    if not any(serie_tiene_contenido(base[columna]).any() for columna in COLUMNAS_UBICACION):
        if not requerir_ubicacion:
            return base
        raise ValueError(
            "No fue posible relacionar las columnas de barrio, dirección o ciudad con la hoja principal de SAEPlus."
        )

    return base


def normalizar_texto(valor):
    if pd.isna(valor):
        return ''
    texto = unicodedata.normalize('NFKD', str(valor))
    texto = ''.join(char for char in texto if not unicodedata.combining(char))
    texto = texto.upper().strip()
    texto = re.sub(r'[^A-Z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def precinto_vacio(valor):
    return not valor_tiene_contenido(valor)


def es_status_alerta(valor):
    return normalizar_texto(valor).lower() in ESTADOS_ALERTA


def referencia_direccion(valor):
    texto = normalizar_texto(valor)
    if not texto:
        return ''

    componentes = []
    for token in texto.split():
        if token in TIPOS_VIA:
            continue
        if not re.search(r'\d', token):
            continue

        coincidencia = re.match(r'(\d+)', token)
        if coincidencia is None:
            continue

        componente = coincidencia.group(1)
        if componente not in componentes:
            componentes.append(componente)

    if len(componentes) >= 2:
        principales = sorted(componentes[:2], key=lambda valor: (int(valor), valor))
        return ' '.join(principales)
    if componentes:
        return componentes[0]

    tokens = texto.split()
    return ' '.join(tokens[:3])


def construir_claves_ubicacion(df):
    trabajo = df.copy()
    trabajo['ciudad_norm'] = trabajo['ciudad'].apply(normalizar_texto)
    trabajo['barrio_norm'] = trabajo['barrio'].apply(normalizar_texto)
    trabajo['direccion_ref'] = trabajo['direccion'].apply(referencia_direccion)
    trabajo['direccion_ref_norm'] = trabajo['direccion_ref'].apply(normalizar_texto)

    trabajo['clave_barrio'] = trabajo.apply(
        lambda fila: ' | '.join(
            parte for parte in [fila['ciudad_norm'], fila['barrio_norm']] if parte
        ) if fila['barrio_norm'] else '',
        axis=1
    )
    trabajo['clave_direccion'] = trabajo.apply(
        lambda fila: ' | '.join(
            parte for parte in [fila['ciudad_norm'], fila['direccion_ref_norm']] if parte
        ) if fila['direccion_ref_norm'] else '',
        axis=1
    )
    trabajo['clave_barrio_direccion'] = trabajo.apply(
        lambda fila: ' | '.join(
            parte for parte in [fila['ciudad_norm'], fila['barrio_norm'], fila['direccion_ref_norm']] if parte
        ) if fila['barrio_norm'] and fila['direccion_ref_norm'] else '',
        axis=1
    )
    return trabajo


def preparar_lado_detalle(df, rol):
    if rol == 'sin precinto':
        columnas = {
            'n abonado': 'n° abonado sin precinto',
            'nombre': 'nombre sin precinto',
            'estatus': 'estatus saeplus sin precinto',
            'status': 'status smartolt sin precinto',
            'precinto': 'precinto saeplus',
            'direccion': 'dirección sin precinto',
            'barrio': 'barrio',
            'ciudad': 'ciudad',
            'direccion_ref': 'referencia dirección',
            'equipo maco': 'equipo maco sin precinto',
            'clave_barrio': 'clave_barrio',
            'clave_direccion': 'clave_direccion',
            'clave_barrio_direccion': 'clave_barrio_direccion',
        }
    else:
        columnas = {
            'n abonado': 'n° abonado alerta',
            'nombre': 'nombre alerta',
            'estatus': 'estatus saeplus alerta',
            'status': 'status smartolt alerta',
            'direccion': 'dirección alerta',
            'equipo maco': 'equipo maco alerta',
            'olt': 'olt alerta',
            'sn': 'sn alerta',
            'clave_barrio': 'clave_barrio',
            'clave_direccion': 'clave_direccion',
            'clave_barrio_direccion': 'clave_barrio_direccion',
        }

    disponibles = [columna for columna in columnas if columna in df.columns]
    trabajo = df[disponibles].copy()
    return trabajo.rename(columns={columna: columnas[columna] for columna in disponibles})


def combinar_criterios(series):
    vistos = []
    for valor in series:
        if valor and valor not in vistos:
            vistos.append(valor)
    return ', '.join(vistos)


def unir_valores_unicos(series):
    vistos = []
    for valor in series:
        if not valor_tiene_contenido(valor):
            continue
        texto = str(valor).strip()
        if texto not in vistos:
            vistos.append(texto)
    return ', '.join(vistos)


def construir_detalle_cercania(df):
    sin_precinto = df[df['precinto'].apply(precinto_vacio)].copy()
    alertas = df[df['status'].apply(es_status_alerta)].copy()

    detalle_base = pd.DataFrame(columns=COLUMNAS_DETALLE)
    if sin_precinto.empty or alertas.empty:
        return detalle_base

    izquierda = preparar_lado_detalle(sin_precinto, 'sin precinto')
    derecha = preparar_lado_detalle(alertas, 'alerta')

    pares = []

    criterios = [
        ('Mismo abonado', ['n° abonado sin precinto', 'n° abonado alerta']),
        ('Mismo barrio y dirección aproximada', ['clave_barrio_direccion']),
        ('Misma dirección aproximada', ['clave_direccion']),
        ('Mismo barrio', ['clave_barrio']),
    ]

    for criterio, llaves in criterios:
        if criterio == 'Mismo abonado':
            merge = izquierda.merge(
                derecha,
                left_on='n° abonado sin precinto',
                right_on='n° abonado alerta',
                how='inner'
            )
        else:
            llave = llaves[0]
            izquierda_filtrada = izquierda[izquierda[llave].apply(valor_tiene_contenido)]
            derecha_filtrada = derecha[derecha[llave].apply(valor_tiene_contenido)]
            if izquierda_filtrada.empty or derecha_filtrada.empty:
                continue
            merge = izquierda_filtrada.merge(derecha_filtrada, on=llave, how='inner')

        if merge.empty:
            continue

        merge['criterio cercanía'] = criterio
        pares.append(merge)

    if not pares:
        return detalle_base

    detalle = pd.concat(pares, ignore_index=True, sort=False)

    prioridad = {
        'Mismo abonado': 0,
        'Mismo barrio y dirección aproximada': 1,
        'Misma dirección aproximada': 2,
        'Mismo barrio': 3,
    }
    detalle['orden_criterio'] = detalle['criterio cercanía'].map(prioridad).fillna(99)
    detalle = detalle.sort_values(
        by=['orden_criterio', 'n° abonado sin precinto', 'n° abonado alerta']
    )
    detalle = detalle.drop_duplicates(
        subset=['n° abonado sin precinto', 'n° abonado alerta'],
        keep='first'
    )

    if 'ciudad sin precinto' in detalle.columns:
        detalle['ciudad'] = detalle['ciudad sin precinto']
    else:
        detalle['ciudad'] = pd.NA

    if 'barrio sin precinto' in detalle.columns:
        detalle['barrio'] = detalle['barrio sin precinto']
    else:
        detalle['barrio'] = pd.NA

    columnas_finales = [
        'criterio cercanía',
        'ciudad',
        'barrio',
        'referencia dirección',
        'n° abonado sin precinto',
        'nombre sin precinto',
        'estatus saeplus sin precinto',
        'status smartolt sin precinto',
        'precinto saeplus',
        'dirección sin precinto',
        'n° abonado alerta',
        'nombre alerta',
        'estatus saeplus alerta',
        'status smartolt alerta',
        'dirección alerta',
        'olt alerta',
        'sn alerta',
    ]

    detalle = detalle[columnas_finales + ['orden_criterio']].copy()
    detalle = detalle.sort_values(
        by=['orden_criterio', 'ciudad', 'barrio', 'referencia dirección', 'n° abonado sin precinto', 'n° abonado alerta'],
        na_position='last'
    ).drop(columns=['orden_criterio'])
    return detalle.reset_index(drop=True)


def construir_resumen_sugerido(df, detalle):
    resumen_base = pd.DataFrame(columns=COLUMNAS_RESUMEN)
    if detalle.empty:
        return resumen_base

    sin_precinto = df[df['precinto'].apply(precinto_vacio)].copy()
    sin_precinto['n° abonado'] = sin_precinto['n abonado']
    sin_precinto['estatus saeplus'] = sin_precinto['estatus']
    sin_precinto['status smartolt propio'] = sin_precinto['status']
    sin_precinto['dirección'] = sin_precinto['direccion']
    sin_precinto['referencia dirección'] = sin_precinto['direccion_ref']

    base = sin_precinto[
        ['n° abonado', 'nombre', 'estatus saeplus', 'status smartolt propio', 'precinto', 'ciudad', 'barrio', 'dirección', 'referencia dirección']
    ].copy()

    agrupado = (
        detalle.groupby('n° abonado sin precinto', as_index=False)
        .agg({
            'n° abonado alerta': unir_valores_unicos,
            'status smartolt alerta': unir_valores_unicos,
            'criterio cercanía': combinar_criterios,
        })
        .rename(columns={
            'n° abonado sin precinto': 'n° abonado',
            'n° abonado alerta': 'abonados alerta cercanos',
            'status smartolt alerta': 'estatus smartolt cercanos',
            'criterio cercanía': 'criterios detectados',
        })
    )
    agrupado['cantidad alertas cercanas'] = agrupado['abonados alerta cercanos'].apply(
        lambda valor: len([parte for parte in str(valor).split(',') if parte.strip()]) if valor_tiene_contenido(valor) else 0
    )

    resumen = base.merge(agrupado, on='n° abonado', how='inner')
    return resumen[COLUMNAS_RESUMEN].sort_values(
        by=['cantidad alertas cercanas', 'ciudad', 'barrio', 'referencia dirección', 'n° abonado'],
        ascending=[False, True, True, True, True],
        na_position='last'
    ).reset_index(drop=True)


def construir_analisis_precintos_cercanos_desde_coincidencias(coincidencias):
    coincidencias = construir_claves_ubicacion(coincidencias)
    detalle = construir_detalle_cercania(coincidencias)
    resumen = construir_resumen_sugerido(coincidencias, detalle)
    return resumen, detalle


def procesar_precintos_cercanos(saeplus_file, olt_file):
    saeplus = cargar_saeplus_con_ubicacion(saeplus_file, requerir_ubicacion=False)
    olt = procesar_archivo_csv_solo(olt_file)
    olt.columns = olt.columns.str.lower()
    validar_columnas(olt, ['nsn', 'name', 'status', 'sn', 'olt'], 'SmartOLT')

    saeplus['equipo maco'] = saeplus['equipo maco'].apply(normalizar_equipo)
    olt['nsn'] = olt['nsn'].apply(normalizar_equipo)

    coincidencias = pd.merge(
        saeplus,
        olt[['nsn', 'name', 'status', 'sn', 'olt']],
        how='inner',
        left_on='equipo maco',
        right_on='nsn'
    )
    resumen, detalle = construir_analisis_precintos_cercanos_desde_coincidencias(coincidencias)

    return {
        'data': resumen,
        'columns': resumen.columns.tolist(),
        'num_casos': int(resumen.shape[0]),
        'excel': excel_desde_hojas([
            ('Zonas sugeridas', resumen),
            ('Detalle cruce', detalle),
        ]),
    }
