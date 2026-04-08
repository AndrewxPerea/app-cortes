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
        'LOS': 0,
        'POWER FAIL': 1,
        'OFFLINE': 2,
    }
    return orden.get(normalizar_texto(valor), 99)


def prioridad_revision(valor):
    estado = normalizar_texto(valor)
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


def construir_posibles_precintos_perdidos(coincidencias):
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
        tabla['status'].apply(es_status_alerta) &
        tabla['estatus'].apply(es_estatus_permitido)
    ].copy()

    if tabla.empty:
        return tabla

    tabla.insert(0, 'prioridad', tabla['status'].apply(prioridad_revision))
    tabla.insert(1, 'hallazgo', tabla['status'].apply(hallazgo_precinto))
    tabla['referencia dirección'] = tabla['dirección'].apply(referencia_direccion)
    tabla['orden_status'] = tabla['status'].apply(prioridad_status)
    tabla = tabla.sort_values(
        by=['orden_status', 'ciudad', 'barrio', 'referencia dirección', 'dirección', 'n° abonado'],
        ascending=[True, True, True, True, True, True],
        na_position='last'
    ).drop(columns=['orden_status'])
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


def construir_ubicaciones_sugeridas_precintos(saeplus, comparacion_precintos):
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
        'ciudad posible',
        'barrio posible',
        'dirección posible',
        'precinto posible saeplus',
    ]
    if comparacion_precintos is None:
        return None

    referencias = comparacion_precintos[
        comparacion_precintos['coincide en saeplus'].astype(str).str.strip().str.upper() == 'SI'
    ].copy()
    if referencias.empty:
        return pd.DataFrame(columns=columnas)

    referencias['referencia dirección'] = referencias['dirección'].apply(referencia_direccion)
    referencias['ciudad_norm'] = referencias['ciudad'].apply(normalizar_texto)
    referencias['barrio_norm'] = referencias['barrio'].apply(normalizar_texto)
    referencias['direccion_ref_norm'] = referencias['referencia dirección'].apply(normalizar_texto)

    candidatos = saeplus.copy()
    candidatos['n° abonado'] = candidatos['n abonado']
    candidatos['dirección'] = candidatos['direccion']
    candidatos['referencia dirección'] = candidatos['dirección'].apply(referencia_direccion)
    candidatos['ciudad_norm'] = candidatos['ciudad'].apply(normalizar_texto)
    candidatos['barrio_norm'] = candidatos['barrio'].apply(normalizar_texto)
    candidatos['direccion_ref_norm'] = candidatos['referencia dirección'].apply(normalizar_texto)
    candidatos = candidatos[
        candidatos['precinto'].apply(precinto_vacio) &
        candidatos['estatus'].apply(es_estatus_permitido)
    ].copy()

    if candidatos.empty:
        return pd.DataFrame(columns=columnas)

    criterios = [
        ('Mismo barrio y dirección aproximada', lambda fila, ref: fila['ciudad_norm'] == ref['ciudad_norm'] and fila['barrio_norm'] == ref['barrio_norm'] and fila['direccion_ref_norm'] == ref['direccion_ref_norm'] and bool(ref['barrio_norm']) and bool(ref['direccion_ref_norm'])),
        ('Misma dirección aproximada', lambda fila, ref: fila['ciudad_norm'] == ref['ciudad_norm'] and fila['direccion_ref_norm'] == ref['direccion_ref_norm'] and bool(ref['direccion_ref_norm'])),
        ('Mismo barrio', lambda fila, ref: fila['ciudad_norm'] == ref['ciudad_norm'] and fila['barrio_norm'] == ref['barrio_norm'] and bool(ref['barrio_norm'])),
    ]

    registros = []
    vistos = set()
    for _, referencia in referencias.iterrows():
        candidatos_ref = candidatos[candidatos['n° abonado'] != referencia['n° abonado']]
        for criterio, comparador in criterios:
            for _, candidato in candidatos_ref.iterrows():
                llave = (
                    str(referencia['precinto cargado']),
                    str(referencia['n° abonado']),
                    str(candidato['n° abonado']),
                )
                if llave in vistos:
                    continue
                if not comparador(candidato, referencia):
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
        'Misma dirección aproximada': 1,
        'Mismo barrio': 2,
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


def generar_excel_comparativo(posibles_perdidos, comparacion_precintos=None, ubicaciones_precintos=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if comparacion_precintos is not None:
            escribir_hoja(writer, 'Precintos cargados', comparacion_precintos)
        if ubicaciones_precintos is not None:
            escribir_hoja(writer, 'Ubicaciones sugeridas', ubicaciones_precintos)
        escribir_hoja(writer, 'Posibles precintos perdidos', posibles_perdidos)
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

    coincidencias_raw = pd.merge(
        saeplus,
        olt[['nsn', 'name', 'status', 'sn', 'olt']],
        how='inner',
        left_on='equipo maco',
        right_on='nsn'
    )

    posibles_perdidos = construir_posibles_precintos_perdidos(coincidencias_raw)
    comparacion_precintos = construir_comparacion_precintos_cargados(saeplus, texto_precintos)
    ubicaciones_precintos = construir_ubicaciones_sugeridas_precintos(saeplus, comparacion_precintos)

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
            ubicaciones_precintos
        ),
    }
