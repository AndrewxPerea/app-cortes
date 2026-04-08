import io
import re
import unicodedata

import pandas as pd

from funciones import procesar_archivo_csv_solo
from services.common import validar_columnas
from services.precintos_cercanos import (
    cargar_saeplus_con_ubicacion,
    normalizar_texto,
    referencia_direccion,
    valor_tiene_contenido,
)


def serie_o_vacia(df, columna):
    if isinstance(columna, (list, tuple)):
        for nombre in columna:
            if nombre in df.columns:
                return df[nombre]
    elif columna in df.columns:
        return df[columna]
    return pd.Series([None] * len(df), index=df.index)


def precinto_vacio(valor):
    if pd.isna(valor):
        return True
    return str(valor).strip() == ''


def valor_ordenado(serie):
    return serie.fillna('').astype(str).str.strip().str.upper()


def es_status_alerta(valor):
    return normalizar_texto(valor) in {'OFFLINE', 'POWER FAIL', 'LOS'}


def es_estatus_permitido(valor):
    return normalizar_texto(valor) in {'ACTIVO', 'CORTADO'}


def prioridad_status(valor):
    orden = {
        'ONLINE': -1,
        'LOS': 0,
        'POWER FAIL': 1,
        'OFFLINE': 2,
    }
    return orden.get(normalizar_texto(valor), 99)


def prioridad_revision(valor):
    estado = normalizar_texto(valor)
    if estado == 'ONLINE':
        return 'Alta'
    if estado == 'LOS':
        return 'Alta'
    if estado == 'POWER FAIL':
        return 'Media'
    if estado == 'OFFLINE':
        return 'Baja'
    return 'Por revisar'


def hallazgo_precinto(valor):
    estado = normalizar_texto(valor)
    if not estado:
        return 'Abonado sin precinto en SAEPlus'
    return f'Abonado sin precinto en SAEPlus y con {estado} en SmartOLT'


def normalizar_precinto(valor):
    if pd.isna(valor):
        return ''

    if isinstance(valor, (int, float)):
        if float(valor).is_integer():
            texto = str(int(valor))
        else:
            texto = str(valor).strip()
    else:
        texto = str(valor).strip()

    if not texto:
        return ''

    if re.fullmatch(r'\d+\.0+', texto):
        texto = texto.split('.', 1)[0]

    texto = unicodedata.normalize('NFKD', texto)
    texto = ''.join(char for char in texto if not unicodedata.combining(char))
    texto = texto.upper().strip()
    texto = re.sub(r'[^A-Z0-9]+', '', texto)

    if texto.isdigit():
        texto = texto.lstrip('0') or '0'

    return texto


def construir_zona_operativa(df):
    zona = serie_o_vacia(df, ['zone', 'zona']).fillna('').astype(str).str.strip()
    board = serie_o_vacia(df, 'board').fillna('').astype(str).str.strip()
    port = serie_o_vacia(df, 'port').fillna('').astype(str).str.strip()

    zona_board_port = pd.Series([''] * len(df), index=df.index)
    mascara_board_port = board.ne('') & port.ne('')
    zona_board_port.loc[mascara_board_port] = (
        'BOARD ' + board.loc[mascara_board_port] + ' PORT ' + port.loc[mascara_board_port]
    )

    mascara_solo_board = board.ne('') & port.eq('')
    zona_board_port.loc[mascara_solo_board] = 'BOARD ' + board.loc[mascara_solo_board]

    mascara_solo_port = port.ne('') & board.eq('')
    zona_board_port.loc[mascara_solo_port] = 'PORT ' + port.loc[mascara_solo_port]

    return zona.where(zona.ne(''), zona_board_port)


def construir_analisis_status_smartolt(coincidencias, solo_alertas=True):
    tabla = pd.DataFrame({
        'status': serie_o_vacia(coincidencias, 'status'),
        'n° abonado': serie_o_vacia(coincidencias, ['n° abonado', 'n abonado']),
        'documento': serie_o_vacia(coincidencias, 'documento'),
        'nombre': serie_o_vacia(coincidencias, 'nombre'),
        'estatus': serie_o_vacia(coincidencias, 'estatus'),
        'ciudad': serie_o_vacia(coincidencias, 'ciudad'),
        'barrio': serie_o_vacia(coincidencias, 'barrio'),
        'dirección': serie_o_vacia(coincidencias, ['dirección', 'direccion']),
        'precinto': serie_o_vacia(coincidencias, 'precinto'),
        'equipo maco': serie_o_vacia(coincidencias, 'equipo maco'),
        'sn': serie_o_vacia(coincidencias, 'sn'),
        'olt': serie_o_vacia(coincidencias, 'olt'),
    })

    tabla = tabla[
        tabla['precinto'].apply(precinto_vacio) &
        tabla['estatus'].apply(es_estatus_permitido)
    ].copy()

    if solo_alertas:
        tabla = tabla[tabla['status'].apply(es_status_alerta)].copy()

    if tabla.empty:
        return tabla

    tabla.insert(0, 'prioridad', tabla['status'].apply(prioridad_revision))
    tabla.insert(1, 'hallazgo', tabla['status'].apply(hallazgo_precinto))
    tabla['referencia dirección'] = tabla['dirección'].apply(referencia_direccion)
    tabla['orden_status'] = tabla['status'].apply(prioridad_status)
    tabla['status_ordenado'] = valor_ordenado(tabla['status'])
    tabla = tabla.sort_values(
        by=['orden_status', 'status_ordenado', 'ciudad', 'barrio', 'referencia dirección', 'dirección', 'n° abonado'],
        ascending=[True, True, True, True, True, True, True],
        na_position='last'
    ).drop(columns=['orden_status', 'status_ordenado'])
    tabla = tabla[
        [
            'prioridad',
            'hallazgo',
            'n° abonado',
            'documento',
            'nombre',
            'estatus',
            'status',
            'ciudad',
            'barrio',
            'referencia dirección',
            'dirección',
            'precinto',
            'equipo maco',
            'sn',
            'olt',
        ]
    ]
    return tabla.reset_index(drop=True)


def construir_posibles_precintos_perdidos(coincidencias):
    return construir_analisis_status_smartolt(coincidencias, solo_alertas=True)


def construir_todos_los_estados_smartolt(coincidencias):
    return construir_analisis_status_smartolt(coincidencias, solo_alertas=False)


def unir_textos_unicos(serie):
    valores = []
    for valor in serie:
        if not valor_tiene_contenido(valor):
            continue
        texto = str(valor).strip()
        if texto not in valores:
            valores.append(texto)
    return ' | '.join(valores)


def primer_texto_contenido(serie):
    for valor in serie:
        if valor_tiene_contenido(valor):
            return str(valor).strip()
    return ''


def cargar_precintos_referencia(texto_precintos):
    if texto_precintos is None:
        return None

    texto = str(texto_precintos or '')
    lineas = [linea.strip() for linea in texto.splitlines() if linea.strip()]

    if len(lineas) > 1:
        valores = []
        for linea in lineas:
            partes = [parte.strip() for parte in re.split(r'[\t,;]+', linea) if parte.strip()]
            if partes:
                valores.append(partes[0])
                continue

            tokens = [token.strip() for token in linea.split() if token.strip()]
            if tokens:
                valores.append(tokens[0])
    else:
        valores = [
            valor.strip()
            for valor in re.split(r'[\s,;]+', texto)
            if valor.strip()
        ]

    if not valores:
        return None

    precintos = pd.DataFrame({
        'origen captura': ['Formulario web'] * len(valores),
        'precinto cargado': valores,
    })
    precintos['precinto normalizado'] = precintos['precinto cargado'].apply(normalizar_precinto)
    precintos = precintos[precintos['precinto normalizado'].apply(valor_tiene_contenido)].copy()
    if precintos.empty:
        return None

    precintos.insert(0, 'orden carga', range(1, len(precintos) + 1))
    return precintos


def construir_comparacion_precintos_cargados(saeplus, texto_precintos):
    precintos_cargados = cargar_precintos_referencia(texto_precintos)
    if precintos_cargados is None:
        return None

    sae_precintos = saeplus[saeplus['precinto'].apply(valor_tiene_contenido)].copy()
    sae_precintos['precinto normalizado'] = sae_precintos['precinto'].apply(normalizar_precinto)
    sae_precintos['n° abonado'] = sae_precintos['n abonado']
    sae_precintos['dirección'] = sae_precintos['direccion']

    comparacion = precintos_cargados.merge(
        sae_precintos[
            [
                'precinto normalizado',
                'precinto',
                'n° abonado',
                'documento',
                'nombre',
                'estatus',
                'ciudad',
                'barrio',
                'dirección',
                'equipo maco',
            ]
        ],
        on='precinto normalizado',
        how='left'
    )
    comparacion['coincide en saeplus'] = comparacion['n° abonado'].notna().map({
        True: 'Si',
        False: 'No',
    })

    comparacion = comparacion.rename(columns={
        'precinto': 'precinto saeplus',
    })
    comparacion = comparacion[
        [
            'precinto cargado',
            'coincide en saeplus',
            'precinto saeplus',
            'n° abonado',
            'documento',
            'nombre',
            'estatus',
            'ciudad',
            'barrio',
            'dirección',
            'orden carga',
    ]
    ].sort_values(by=['orden carga', 'coincide en saeplus', 'n° abonado'], ascending=[True, False, True])

    return comparacion.drop(columns=['orden carga']).reset_index(drop=True)


def construir_contexto_tecnico(coincidencias):
    columnas = ['abonado_key', 'status', 'sn', 'olt', 'zona operativa']
    if coincidencias is None or coincidencias.empty:
        return pd.DataFrame(columns=columnas)

    contexto = pd.DataFrame({
        'n° abonado': serie_o_vacia(coincidencias, ['n abonado', 'n° abonado']),
        'status': serie_o_vacia(coincidencias, 'status'),
        'sn': serie_o_vacia(coincidencias, 'sn'),
        'olt': serie_o_vacia(coincidencias, 'olt'),
        'zona operativa': construir_zona_operativa(coincidencias),
    })
    contexto['abonado_key'] = contexto['n° abonado'].apply(normalizar_precinto)
    contexto = contexto[contexto['abonado_key'].apply(valor_tiene_contenido)].copy()

    if contexto.empty:
        return pd.DataFrame(columns=columnas)

    contexto = (
        contexto.groupby('abonado_key', as_index=False)
        .agg({
            'status': primer_texto_contenido,
            'sn': primer_texto_contenido,
            'olt': primer_texto_contenido,
            'zona operativa': primer_texto_contenido,
        })
    )
    return contexto[columnas]


def construir_referencias_precintos(comparacion_precintos, coincidencias=None):
    if comparacion_precintos is None:
        return pd.DataFrame()

    referencias = comparacion_precintos[
        comparacion_precintos['coincide en saeplus'].astype(str).str.strip().str.upper() == 'SI'
    ].copy()
    if referencias.empty:
        return referencias

    contexto_tecnico = construir_contexto_tecnico(coincidencias)
    referencias['abonado_key'] = referencias['n° abonado'].apply(normalizar_precinto)
    if not contexto_tecnico.empty:
        referencias = referencias.merge(contexto_tecnico, on='abonado_key', how='left')

    referencias['referencia dirección'] = referencias['dirección'].apply(referencia_direccion)
    referencias['ciudad_norm'] = referencias['ciudad'].apply(normalizar_texto)
    referencias['barrio_norm'] = referencias['barrio'].apply(normalizar_texto)
    referencias['direccion_ref_norm'] = referencias['referencia dirección'].apply(normalizar_texto)
    referencias['olt_norm'] = serie_o_vacia(referencias, 'olt').apply(normalizar_texto)
    referencias['zona_norm'] = serie_o_vacia(referencias, 'zona operativa').apply(normalizar_texto)
    return referencias


def detectar_criterio_ubicacion(fila, referencia):
    misma_ciudad = fila['ciudad_norm'] == referencia['ciudad_norm'] and bool(referencia['ciudad_norm'])
    mismo_barrio = fila['barrio_norm'] == referencia['barrio_norm'] and bool(referencia['barrio_norm'])
    misma_direccion = fila['direccion_ref_norm'] == referencia['direccion_ref_norm'] and bool(referencia['direccion_ref_norm'])
    mismo_olt = fila.get('olt_norm', '') == referencia.get('olt_norm', '') and bool(referencia.get('olt_norm', ''))
    misma_zona = fila.get('zona_norm', '') == referencia.get('zona_norm', '') and bool(referencia.get('zona_norm', ''))
    ciudades_compatibles = misma_ciudad or not fila['ciudad_norm'] or not referencia['ciudad_norm']

    if misma_ciudad and mismo_barrio and misma_direccion:
        return 'Mismo barrio y dirección aproximada'
    if ciudades_compatibles and mismo_barrio and mismo_olt and misma_zona:
        return 'Mismo OLT, zona y barrio'
    return None


def filtrar_posibles_perdidos_por_precintos(posibles_perdidos, comparacion_precintos, coincidencias=None):
    columnas_contexto = [
        'precinto cargado relacionado',
        'precinto saeplus relacionado',
        'n° abonado referencia',
        'criterio ubicación',
    ]
    if comparacion_precintos is None:
        return posibles_perdidos

    referencias = construir_referencias_precintos(comparacion_precintos, coincidencias)
    if referencias.empty:
        return pd.DataFrame(columns=[*columnas_contexto, *posibles_perdidos.columns.tolist()])

    posibles = posibles_perdidos.copy()
    if posibles.empty:
        return pd.DataFrame(columns=[*columnas_contexto, *posibles.columns.tolist()])

    contexto_tecnico = construir_contexto_tecnico(coincidencias)
    posibles['abonado_key'] = posibles['n° abonado'].apply(normalizar_precinto)
    if not contexto_tecnico.empty:
        posibles = posibles.merge(
            contexto_tecnico[['abonado_key', 'zona operativa']],
            on='abonado_key',
            how='left'
        )

    posibles['ciudad_norm'] = posibles['ciudad'].apply(normalizar_texto)
    posibles['barrio_norm'] = posibles['barrio'].apply(normalizar_texto)
    posibles['direccion_ref_norm'] = posibles['referencia dirección'].apply(normalizar_texto)
    posibles['olt_norm'] = posibles['olt'].apply(normalizar_texto)
    posibles['zona_norm'] = serie_o_vacia(posibles, 'zona operativa').apply(normalizar_texto)

    relaciones = []
    for indice_posible, posible in posibles.iterrows():
        for _, referencia in referencias.iterrows():
            criterio = detectar_criterio_ubicacion(posible, referencia)
            if not criterio:
                continue
            relaciones.append({
                'indice_posible': indice_posible,
                'precinto cargado relacionado': referencia['precinto cargado'],
                'precinto saeplus relacionado': referencia['precinto saeplus'],
                'n° abonado referencia': referencia['n° abonado'],
                'criterio ubicación': criterio,
            })

    if not relaciones:
        return pd.DataFrame(columns=[*columnas_contexto, *posibles_perdidos.columns.tolist()])

    contexto = pd.DataFrame(relaciones)
    prioridad_criterio = {
        'Mismo barrio y dirección aproximada': 0,
        'Mismo OLT, zona y barrio': 1,
    }
    contexto['orden criterio'] = contexto['criterio ubicación'].map(prioridad_criterio).fillna(99)
    contexto = contexto.sort_values(
        by=['indice_posible', 'orden criterio', 'precinto cargado relacionado', 'n° abonado referencia']
    )
    contexto_agrupado = (
        contexto.groupby('indice_posible', as_index=False)
        .agg({
            'precinto cargado relacionado': unir_textos_unicos,
            'precinto saeplus relacionado': unir_textos_unicos,
            'n° abonado referencia': unir_textos_unicos,
            'criterio ubicación': unir_textos_unicos,
        })
        .set_index('indice_posible')
    )

    posibles = posibles.join(contexto_agrupado, how='inner')
    columnas_auxiliares = [
        'abonado_key',
        'zona operativa',
        'ciudad_norm',
        'barrio_norm',
        'direccion_ref_norm',
        'olt_norm',
        'zona_norm',
    ]
    posibles = posibles.drop(columns=[columna for columna in columnas_auxiliares if columna in posibles.columns])
    return posibles[
        [
            'precinto cargado relacionado',
            'precinto saeplus relacionado',
            'n° abonado referencia',
            'criterio ubicación',
            *posibles_perdidos.columns.tolist(),
        ]
    ].reset_index(drop=True)


def construir_ubicaciones_sugeridas_precintos(saeplus, comparacion_precintos, coincidencias=None):
    columnas = [
        'precinto cargado',
        'precinto saeplus',
        'n° abonado referencia',
        'nombre referencia',
        'ciudad referencia',
        'barrio referencia',
        'dirección referencia',
        'criterio ubicación',
        'n° abonado posible',
        'documento posible',
        'nombre posible',
        'estatus posible',
        'status smartolt posible',
        'ciudad posible',
        'barrio posible',
        'dirección posible',
        'precinto posible saeplus',
    ]
    if comparacion_precintos is None:
        return None

    referencias = construir_referencias_precintos(comparacion_precintos, coincidencias)
    if referencias.empty:
        return pd.DataFrame(columns=columnas)

    candidatos = saeplus.copy()
    candidatos['n° abonado'] = candidatos['n abonado']
    candidatos['dirección'] = candidatos['direccion']
    candidatos['referencia dirección'] = candidatos['dirección'].apply(referencia_direccion)
    candidatos['ciudad_norm'] = candidatos['ciudad'].apply(normalizar_texto)
    candidatos['barrio_norm'] = candidatos['barrio'].apply(normalizar_texto)
    candidatos['direccion_ref_norm'] = candidatos['referencia dirección'].apply(normalizar_texto)
    candidatos['abonado_key'] = candidatos['n° abonado'].apply(normalizar_precinto)

    contexto_tecnico = construir_contexto_tecnico(coincidencias)
    if not contexto_tecnico.empty:
        candidatos = candidatos.merge(
            contexto_tecnico[['abonado_key', 'status']],
            on='abonado_key',
            how='left'
        )
    else:
        candidatos['status'] = pd.NA

    candidatos = candidatos[
        candidatos['precinto'].apply(precinto_vacio) &
        candidatos['estatus'].apply(es_estatus_permitido)
    ].copy()

    if candidatos.empty:
        return pd.DataFrame(columns=columnas)

    registros = []
    vistos = set()
    for _, referencia in referencias.iterrows():
        candidatos_ref = candidatos[candidatos['n° abonado'] != referencia['n° abonado']]
        for _, candidato in candidatos_ref.iterrows():
            llave = (
                str(referencia['precinto cargado']),
                str(referencia['n° abonado']),
                str(candidato['n° abonado']),
            )
            if llave in vistos:
                continue

            criterio = detectar_criterio_ubicacion(candidato, referencia)
            if not criterio:
                continue

            vistos.add(llave)
            registros.append({
                'precinto cargado': referencia['precinto cargado'],
                'precinto saeplus': referencia['precinto saeplus'],
                'n° abonado referencia': referencia['n° abonado'],
                'nombre referencia': referencia['nombre'],
                'ciudad referencia': referencia['ciudad'],
                'barrio referencia': referencia['barrio'],
                'dirección referencia': referencia['dirección'],
                'criterio ubicación': criterio,
                'n° abonado posible': candidato['n° abonado'],
                'documento posible': candidato['documento'],
                'nombre posible': candidato['nombre'],
                'estatus posible': candidato['estatus'],
                'status smartolt posible': candidato['status'],
                'ciudad posible': candidato['ciudad'],
                'barrio posible': candidato['barrio'],
                'dirección posible': candidato['dirección'],
                'precinto posible saeplus': candidato['precinto'],
            })

    if not registros:
        return pd.DataFrame(columns=columnas)

    ubicaciones = pd.DataFrame(registros)
    orden_criterio = {
        'Mismo barrio y dirección aproximada': 0,
        'Mismo OLT, zona y barrio': 1,
    }
    ubicaciones['orden criterio'] = ubicaciones['criterio ubicación'].map(orden_criterio).fillna(99)
    ubicaciones = ubicaciones.sort_values(
        by=[
            'precinto cargado',
            'orden criterio',
            'ciudad referencia',
            'barrio referencia',
            'dirección referencia',
            'n° abonado posible',
        ],
        ascending=[True, True, True, True, True, True],
        na_position='last'
    ).drop(columns=['orden criterio'])
    return ubicaciones[columnas].reset_index(drop=True)


def escribir_hoja(writer, sheet_name, df):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    alerta_falta_precinto = workbook.add_format({'bg_color': '#FDE9D9'})
    alerta_alta = workbook.add_format({'bg_color': '#F4CCCC'})
    alerta_media = workbook.add_format({'bg_color': '#FCE5CD'})
    alerta_baja = workbook.add_format({'bg_color': '#FFF2CC'})

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

    for indice_columna, nombre_columna in enumerate(df.columns):
        ancho_datos = 0 if df.empty else df[nombre_columna].fillna('').astype(str).map(len).max()
        ancho = min(max(len(str(nombre_columna)), int(ancho_datos)) + 2, 38)
        worksheet.set_column(indice_columna, indice_columna, max(ancho, 12))

    if 'coincide en saeplus' in df.columns:
        for fila_excel, coincide in enumerate(df['coincide en saeplus'], start=1):
            if str(coincide).strip().upper() == 'NO':
                worksheet.set_row(fila_excel, None, alerta_falta_precinto)
        return

    if 'prioridad' not in df.columns:
        return

    formato_por_prioridad = {
        'ALTA': alerta_alta,
        'MEDIA': alerta_media,
        'BAJA': alerta_baja,
    }
    for fila_excel, prioridad in enumerate(df['prioridad'], start=1):
        formato = formato_por_prioridad.get(normalizar_texto(prioridad))
        if formato is not None:
            worksheet.set_row(fila_excel, None, formato)


def generar_excel_comparativo(
    posibles_perdidos,
    comparacion_precintos=None,
    ubicaciones_precintos=None,
    todos_los_estados=None
):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if comparacion_precintos is not None:
            escribir_hoja(writer, 'Precintos cargados', comparacion_precintos)
        if ubicaciones_precintos is not None:
            escribir_hoja(writer, 'Ubicaciones sugeridas', ubicaciones_precintos)
        escribir_hoja(writer, 'Posibles precintos perdidos', posibles_perdidos)
        if todos_los_estados is not None:
            escribir_hoja(writer, 'Todos los estados SmartOLT', todos_los_estados)
    output.seek(0)
    return output


def procesar_comparativo_precintos(saeplus_file, olt_file, texto_precintos=''):
    saeplus = cargar_saeplus_con_ubicacion(saeplus_file, requerir_ubicacion=False)
    olt = procesar_archivo_csv_solo(olt_file)

    olt.columns = olt.columns.str.lower()
    validar_columnas(
        saeplus,
        ['equipo maco', 'n abonado', 'documento', 'nombre', 'estatus', 'precinto'],
        'SAEPlus'
    )
    validar_columnas(
        olt,
        ['nsn', 'name', 'status', 'sn', 'olt'],
        'SmartOLT'
    )

    columnas_olt = ['nsn', 'name', 'status', 'sn', 'olt']
    for columna_extra in ['zone', 'zona', 'board', 'port']:
        if columna_extra in olt.columns and columna_extra not in columnas_olt:
            columnas_olt.append(columna_extra)

    coincidencias_raw = pd.merge(
        saeplus,
        olt[columnas_olt],
        how='inner',
        left_on='equipo maco',
        right_on='nsn'
    )

    comparacion_precintos = construir_comparacion_precintos_cargados(saeplus, texto_precintos)
    posibles_perdidos = construir_posibles_precintos_perdidos(coincidencias_raw)
    todos_los_estados = construir_todos_los_estados_smartolt(coincidencias_raw)
    posibles_perdidos = filtrar_posibles_perdidos_por_precintos(
        posibles_perdidos,
        comparacion_precintos,
        coincidencias_raw
    )
    todos_los_estados = filtrar_posibles_perdidos_por_precintos(
        todos_los_estados,
        comparacion_precintos,
        coincidencias_raw
    )
    ubicaciones_precintos = construir_ubicaciones_sugeridas_precintos(
        saeplus,
        comparacion_precintos,
        coincidencias_raw
    )

    if comparacion_precintos is not None:
        data = comparacion_precintos
        columns = comparacion_precintos.columns.tolist()
        num_casos = int(comparacion_precintos.shape[0])
    else:
        data = posibles_perdidos
        columns = posibles_perdidos.columns.tolist()
        num_casos = int(posibles_perdidos.shape[0])

    return {
        'data': data,
        'columns': columns,
        'num_casos': num_casos,
        'excel': generar_excel_comparativo(
            posibles_perdidos,
            comparacion_precintos,
            ubicaciones_precintos,
            todos_los_estados
        ),
    }
