import io
import re

import pandas as pd

from funciones import procesar_archivo_csv_solo
from services.common import validar_columnas
from services.precintos_cercanos import (
    construir_analisis_precintos_cercanos_desde_coincidencias,
    cargar_saeplus_con_ubicacion,
    normalizar_texto,
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


def construir_tabla_comparativa(resultado):
    resultado = resultado[resultado['_merge'] == 'both'].copy()
    tabla = pd.DataFrame({
        'observacion_precinto': serie_o_vacia(resultado, 'precinto').apply(
            lambda valor: 'Precinto vacío en SAEPlus' if precinto_vacio(valor) else ''
        ),
        'n° abonado': serie_o_vacia(resultado, ['n° abonado', 'n abonado']),
        'documento': serie_o_vacia(resultado, 'documento'),
        'nombre': serie_o_vacia(resultado, 'nombre'),
        'estatus': serie_o_vacia(resultado, 'estatus'),
        'barrio': serie_o_vacia(resultado, 'barrio'),
        'dirección': serie_o_vacia(resultado, ['dirección', 'direccion']),
        'precinto': serie_o_vacia(resultado, 'precinto'),
        'equipo maco': serie_o_vacia(resultado, 'equipo maco'),
        'status': serie_o_vacia(resultado, 'status'),
        'sn': serie_o_vacia(resultado, 'sn'),
        'olt': serie_o_vacia(resultado, 'olt'),
    })

    tabla['orden_estatus'] = valor_ordenado(tabla['estatus'])
    tabla['orden_precinto'] = valor_ordenado(tabla['precinto'])
    tabla = tabla.sort_values(
        by=['orden_estatus', 'orden_precinto', 'n° abonado'],
        ascending=[True, True, True],
        na_position='last'
    ).drop(columns=['orden_estatus', 'orden_precinto'])
    return tabla


def cargar_precintos_referencia(texto_precintos):
    if texto_precintos is None:
        return None

    valores = [
        valor.strip()
        for valor in re.split(r'[\s,;]+', str(texto_precintos))
        if valor.strip()
    ]
    if not valores:
        return None

    precintos = pd.DataFrame({
        'origen captura': ['Formulario web'] * len(valores),
        'precinto cargado': valores,
    })
    precintos['precinto normalizado'] = precintos['precinto cargado'].apply(normalizar_texto)
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
    sae_precintos['precinto normalizado'] = sae_precintos['precinto'].apply(normalizar_texto)
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


def escribir_hoja(writer, sheet_name, df):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    destacado = workbook.add_format({'bg_color': '#FFF2CC'})

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

    if 'observacion_precinto' not in df.columns:
        return

    for fila_excel, observacion in enumerate(df['observacion_precinto'], start=1):
        if observacion:
            worksheet.set_row(fila_excel, None, destacado)


def generar_excel_comparativo(coinciden, detalle_cruce, comparacion_precintos=None):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if comparacion_precintos is not None:
            escribir_hoja(writer, 'Precintos cargados', comparacion_precintos)
        escribir_hoja(writer, 'Coinciden', coinciden)
        escribir_hoja(writer, 'Detalle cruce', detalle_cruce)
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

    resultado = pd.merge(
        saeplus,
        olt,
        how='outer',
        left_on='equipo maco',
        right_on='nsn',
        indicator=True,
        suffixes=('_saeplus', '_smartolt')
    )
    coincidencias_raw = pd.merge(
        saeplus,
        olt[['nsn', 'name', 'status', 'sn', 'olt']],
        how='inner',
        left_on='equipo maco',
        right_on='nsn'
    )

    tabla = construir_tabla_comparativa(resultado)
    _, detalle_cruce = construir_analisis_precintos_cercanos_desde_coincidencias(coincidencias_raw)
    comparacion_precintos = construir_comparacion_precintos_cargados(saeplus, texto_precintos)

    return {
        'data': tabla,
        'columns': tabla.columns.tolist(),
        'num_casos': int(tabla.shape[0]),
        'excel': generar_excel_comparativo(
            tabla,
            detalle_cruce,
            comparacion_precintos
        ),
    }
