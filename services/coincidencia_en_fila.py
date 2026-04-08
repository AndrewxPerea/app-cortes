import re
import unicodedata

import pandas as pd

from funciones import abrir_excel_seguro
from services.common import excel_desde_dataframe, excel_desde_hojas


COLUMNAS_ABONADO = {"abonado"}
COLUMNAS_NUMERO_ABONADO = {
    "no abonado",
    "nro abonado",
    "numero abonado",
    "n abonado",
    "n.o abonado",
}


def normalizar_encabezado(valor):
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    texto = texto.lower().strip()
    texto = texto.replace("º", "o").replace("°", "o")
    return re.sub(r"\s+", " ", texto)


def normalizar_valor(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        if float(valor).is_integer():
            return str(int(valor))
        return str(valor).strip()

    texto = str(valor).strip()
    if not texto:
        return None

    if re.fullmatch(r"\d+", texto):
        return str(int(texto))

    if re.fullmatch(r"\d+\.0", texto):
        return texto[:-2]

    return texto


def buscar_columna(columnas, nombres_validos):
    for columna in columnas:
        if normalizar_encabezado(columna) in nombres_validos:
            return columna
    return None


def filtrar_coincidencias(df):
    columna_abonado = buscar_columna(df.columns, COLUMNAS_ABONADO)
    columna_numero_abonado = buscar_columna(df.columns, COLUMNAS_NUMERO_ABONADO)

    if columna_abonado is None or columna_numero_abonado is None:
        raise ValueError(
            "No se encontraron las columnas requeridas: 'ABONADO' y 'n° abonado'."
        )

    valores_abonado = df[columna_abonado].map(normalizar_valor)
    valores_numero_abonado = df[columna_numero_abonado].map(normalizar_valor)
    mascara = valores_abonado.notna() & (valores_abonado == valores_numero_abonado)
    return df.loc[mascara].copy()


def obtener_stream(archivo):
    stream = getattr(archivo, 'stream', archivo)
    if hasattr(stream, 'seek'):
        stream.seek(0)
    return stream


def procesar_coincidencia_en_fila(archivo_excel):
    excel_data = abrir_excel_seguro(obtener_stream(archivo_excel))
    hojas_resultado = []
    vistas_previas = []
    columnas_vista = None
    encontro_columnas = False

    for nombre_hoja in excel_data.sheet_names:
        df = excel_data.parse(sheet_name=nombre_hoja)

        try:
            coincidencias = filtrar_coincidencias(df)
        except ValueError:
            continue

        encontro_columnas = True

        if columnas_vista is None:
            columnas_vista = ['hoja origen', *df.columns.tolist()]

        if coincidencias.empty:
            continue

        hoja_vista = coincidencias.copy()
        hoja_vista.insert(0, 'hoja origen', nombre_hoja)
        vistas_previas.append(hoja_vista)
        hojas_resultado.append((nombre_hoja[:31], coincidencias))

    if not encontro_columnas:
        raise ValueError(
            "El archivo no contiene las columnas requeridas 'ABONADO' y 'n° abonado' en ninguna hoja."
        )

    if vistas_previas:
        data = pd.concat(vistas_previas, ignore_index=True, sort=False)
        excel = excel_desde_hojas(hojas_resultado)
    else:
        data = pd.DataFrame(columns=columnas_vista or ['hoja origen'])
        excel = excel_desde_dataframe(
            pd.DataFrame(columns=(columnas_vista or ['hoja origen'])[1:]),
            'Coincidencias'
        )

    return {
        'data': data,
        'columns': data.columns.tolist(),
        'num_casos': int(data.shape[0]),
        'excel': excel,
    }
