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


def filtrar_coincidencias(df):
    columna_abonado = buscar_columna(df.columns, COLUMNAS_ABONADO)
    columna_numero_abonado = buscar_columna(df.columns, COLUMNAS_NUMERO_ABONADO)

    if columna_abonado is None or columna_numero_abonado is None:
        raise ValueError(
            "No se encontraron las columnas requeridas: 'ABONADO' y 'n° abonado'."
        )

    resumen_abonado = construir_resumen_columna(df, columna_abonado, "ABONADO")
    resumen_numero_abonado = construir_resumen_columna(df, columna_numero_abonado, "n° abonado")
    coincidencias = resumen_abonado.merge(
        resumen_numero_abonado,
        on="valor comparado",
        how="inner"
    )

    if coincidencias.empty:
        return pd.DataFrame(columns=[
            "valor comparado",
            "valor en ABONADO",
            "filas ABONADO",
            "cantidad en ABONADO",
            "valor en n° abonado",
            "filas n° abonado",
            "cantidad en n° abonado",
        ])

    return coincidencias[
        [
            "valor comparado",
            "valor en ABONADO",
            "filas ABONADO",
            "cantidad en ABONADO",
            "valor en n° abonado",
            "filas n° abonado",
            "cantidad en n° abonado",
        ]
    ].sort_values(by="valor comparado").reset_index(drop=True)


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
        df = excel_data.parse(sheet_name=nombre_hoja, dtype=str)

        try:
            coincidencias = filtrar_coincidencias(df)
        except ValueError:
            continue

        encontro_columnas = True

        if columnas_vista is None:
            columnas_vista = [
                'hoja origen',
                'valor comparado',
                'valor en ABONADO',
                'filas ABONADO',
                'cantidad en ABONADO',
                'valor en n° abonado',
                'filas n° abonado',
                'cantidad en n° abonado',
            ]

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
