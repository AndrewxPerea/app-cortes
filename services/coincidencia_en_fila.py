import re
import unicodedata

import pandas as pd

from funciones import abrir_excel_seguro
from services.common import excel_desde_hojas


COLUMNAS_ABONADO = {"abonado"}
COLUMNAS_NUMERO_ABONADO = {
    "no abonado",
    "nro abonado",
    "numero abonado",
    "n abonado",
    "n.o abonado",
}
COLUMNAS_METADATOS = [
    "hoja origen",
    "fila n° abonado",
    "valor comparado",
    "valor en ABONADO",
    "filas ABONADO",
    "cantidad en ABONADO",
]


def normalizar_encabezado(valor):
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    texto = texto.lower().strip()
    texto = texto.replace("º", "o").replace("°", "o")
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    alias = {
        "n o abonado": "n abonado",
        "n de abonado": "n abonado",
        "nro abonado": "nro abonado",
        "numero abonado": "numero abonado",
        "no abonado": "no abonado",
    }
    return alias.get(texto, texto)


def normalizar_valor(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        if float(valor).is_integer():
            return str(int(valor))
        return str(valor).strip()

    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    texto = texto.strip().strip("'").strip('"')
    if not texto:
        return None

    texto = re.sub(r"\s+", "", texto)
    texto = texto.replace(",", "")
    if not texto:
        return None

    if re.fullmatch(r"\d+(\.0+)?", texto):
        return str(int(float(texto)))

    return texto.upper()


def buscar_columna(columnas, nombres_validos):
    for columna in columnas:
        if normalizar_encabezado(columna) in nombres_validos:
            return columna
    return None


def unir_valores_originales(serie):
    valores = []
    for valor in serie:
        if pd.isna(valor):
            continue
        texto = str(valor).strip()
        if texto and texto not in valores:
            valores.append(texto)
    return " | ".join(valores)


def unir_filas(serie):
    filas = []
    for valor in serie:
        if pd.isna(valor):
            continue
        numero = str(int(valor))
        if numero not in filas:
            filas.append(numero)
    return ", ".join(filas)


def construir_resumen_columna(df, columna, etiqueta):
    trabajo = pd.DataFrame({
        "valor original": df[columna],
        "valor comparado": df[columna].map(normalizar_valor),
        "fila excel": range(2, len(df) + 2),
    })
    trabajo = trabajo[trabajo["valor comparado"].notna()].copy()
    if trabajo.empty:
        return pd.DataFrame(
            columns=[
                "valor comparado",
                f"valor en {etiqueta}",
                f"filas {etiqueta}",
                f"cantidad en {etiqueta}",
            ]
        )

    return (
        trabajo.groupby("valor comparado", as_index=False)
        .agg({
            "valor original": unir_valores_originales,
            "fila excel": unir_filas,
        })
        .rename(columns={
            "valor original": f"valor en {etiqueta}",
            "fila excel": f"filas {etiqueta}",
        })
        .assign(**{f"cantidad en {etiqueta}": trabajo.groupby("valor comparado").size().values})
    )


def ordenar_resultado(df):
    if df.empty or "fila n° abonado" not in df.columns:
        return df

    trabajo = df.copy()
    trabajo["_fila_orden"] = pd.to_numeric(trabajo["fila n° abonado"], errors="coerce")
    trabajo = trabajo.sort_values(
        by=["hoja origen", "_fila_orden", "valor comparado"],
        ascending=[True, True, True],
        na_position="last"
    )
    return trabajo.drop(columns=["_fila_orden"]).reset_index(drop=True)


def construir_columnas_resultado(columnas_originales):
    columnas_base = []
    for columna in COLUMNAS_METADATOS:
        if columna not in columnas_base:
            columnas_base.append(columna)

    for columna in columnas_originales:
        if columna not in columnas_base:
            columnas_base.append(columna)

    return columnas_base


def construir_resultados_hoja(df, nombre_hoja):
    columna_abonado = buscar_columna(df.columns, COLUMNAS_ABONADO)
    columna_numero_abonado = buscar_columna(df.columns, COLUMNAS_NUMERO_ABONADO)

    if columna_abonado is None or columna_numero_abonado is None:
        raise ValueError(
            "No se encontraron las columnas requeridas: 'ABONADO' y 'n° abonado'."
        )

    resumen_abonado = construir_resumen_columna(df, columna_abonado, "ABONADO")
    trabajo = df.copy()
    trabajo.insert(0, "hoja origen", nombre_hoja)
    trabajo.insert(1, "fila n° abonado", range(2, len(df) + 2))
    trabajo["valor comparado"] = trabajo[columna_numero_abonado].map(normalizar_valor)

    trabajo = trabajo.merge(resumen_abonado, on="valor comparado", how="left")
    trabajo = trabajo[trabajo["valor comparado"].notna()].copy()

    columnas_resultado = construir_columnas_resultado(df.columns.tolist())
    trabajo = trabajo[columnas_resultado]

    coincidencias = trabajo[trabajo["valor en ABONADO"].notna()].copy()
    no_coinciden = trabajo[trabajo["valor en ABONADO"].isna()].copy()

    return (
        ordenar_resultado(coincidencias),
        ordenar_resultado(no_coinciden),
        columnas_resultado,
    )


def obtener_stream(archivo):
    stream = getattr(archivo, "stream", archivo)
    if hasattr(stream, "seek"):
        stream.seek(0)
    return stream


def procesar_coincidencia_en_fila(archivo_excel):
    excel_data = abrir_excel_seguro(obtener_stream(archivo_excel))
    coincidencias_hojas = []
    no_coincidencias_hojas = []
    columnas_resultado = None
    encontro_columnas = False

    for nombre_hoja in excel_data.sheet_names:
        df = excel_data.parse(sheet_name=nombre_hoja, dtype=str)

        try:
            coincidencias, no_coinciden, columnas_hoja = construir_resultados_hoja(df, nombre_hoja)
        except ValueError:
            continue

        encontro_columnas = True
        if columnas_resultado is None:
            columnas_resultado = columnas_hoja

        if not coincidencias.empty:
            coincidencias_hojas.append(coincidencias)
        if not no_coinciden.empty:
            no_coincidencias_hojas.append(no_coinciden)

    if not encontro_columnas:
        raise ValueError(
            "El archivo no contiene las columnas requeridas 'ABONADO' y 'n° abonado' en ninguna hoja."
        )

    columnas_resultado = columnas_resultado or construir_columnas_resultado([])
    coincidencias_df = (
        pd.concat(coincidencias_hojas, ignore_index=True, sort=False)
        if coincidencias_hojas
        else pd.DataFrame(columns=columnas_resultado)
    )
    no_coinciden_df = (
        pd.concat(no_coincidencias_hojas, ignore_index=True, sort=False)
        if no_coincidencias_hojas
        else pd.DataFrame(columns=columnas_resultado)
    )

    return {
        "data": coincidencias_df,
        "columns": coincidencias_df.columns.tolist(),
        "num_casos": int(coincidencias_df.shape[0]),
        "excel": excel_desde_hojas(
            [
                ("Coinciden", coincidencias_df),
                ("No coinciden", no_coinciden_df),
            ]
        ),
    }
