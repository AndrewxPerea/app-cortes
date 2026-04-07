import io

import pandas as pd

from funciones import procesar_archivo_csv_solo, procesar_archivo_excel_solo
from services.common import validar_columnas


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
    tabla = pd.DataFrame({
        'resultado_comparativo': resultado['_merge'].map({
            'both': 'Coincide',
            'left_only': 'Solo en SAEPlus',
            'right_only': 'Solo en SmartOLT',
        }),
        'observacion_precinto': serie_o_vacia(resultado, 'precinto').apply(
            lambda valor: 'Precinto vacío en SAEPlus' if precinto_vacio(valor) else ''
        ),
        'n° abonado': serie_o_vacia(resultado, 'n° abonado'),
        'documento': serie_o_vacia(resultado, 'documento'),
        'nombre': serie_o_vacia(resultado, 'nombre'),
        'estatus': serie_o_vacia(resultado, 'estatus'),
        'barrio': serie_o_vacia(resultado, 'barrio'),
        'dirección': serie_o_vacia(resultado, ['dirección', 'direccion']),
        'ciudad': serie_o_vacia(resultado, 'ciudad'),
        'precinto': serie_o_vacia(resultado, 'precinto'),
        'equipo maco': serie_o_vacia(resultado, 'equipo maco'),
        'name': serie_o_vacia(resultado, 'name'),
        'status': serie_o_vacia(resultado, 'status'),
        'sn': serie_o_vacia(resultado, 'sn'),
        'olt': serie_o_vacia(resultado, 'olt'),
    })

    tabla = tabla[tabla['resultado_comparativo'] == 'Coincide'].copy()
    tabla['orden_estatus'] = valor_ordenado(tabla['estatus'])
    tabla['orden_precinto'] = valor_ordenado(tabla['precinto'])
    tabla = tabla.sort_values(
        by=['orden_estatus', 'orden_precinto', 'n° abonado'],
        ascending=[True, True, True],
        na_position='last'
    ).drop(columns=['orden_estatus', 'orden_precinto'])
    return tabla


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


def generar_excel_comparativo(coinciden):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        escribir_hoja(writer, 'Coinciden', coinciden)
    output.seek(0)
    return output


def procesar_comparativo_precintos(saeplus_file, olt_file):
    saeplus = procesar_archivo_excel_solo(saeplus_file)
    olt = procesar_archivo_csv_solo(olt_file)

    saeplus.columns = saeplus.columns.str.lower()
    olt.columns = olt.columns.str.lower()

    validar_columnas(
        saeplus,
        ['equipo maco', 'n° abonado', 'documento', 'nombre', 'estatus', 'precinto'],
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

    tabla = construir_tabla_comparativa(resultado)
    return {
        'data': tabla,
        'columns': tabla.columns.tolist(),
        'num_casos': int(tabla.shape[0]),
        'excel': generar_excel_comparativo(tabla),
    }
