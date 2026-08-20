import pandas as pd

from funciones import leer_excel_seguro
from services.common import excel_desde_dataframe


COL_ABONADO = "N\u00b0 Abonado"
COL_EQUIPO_MAC = "EQUIPO MAC"
COL_EQUIPO_KEY = "EQUIPO MACO"
COL_NSN = "NSN"


def _normalizar_nombre_columna(columna):
    texto = str(columna).replace("\u00c2", "").replace("\u00ba", "\u00b0")
    return " ".join(texto.strip().casefold().split())


def _columnas_por_nombre_normalizado(df, nombres):
    nombres = {_normalizar_nombre_columna(nombre) for nombre in nombres}
    return [
        columna
        for columna in df.columns
        if _normalizar_nombre_columna(columna) in nombres
    ]


def _normalizar_columna_estatus_saeplus(df):
    if "estatus" in df.columns:
        return df

    columnas_estatus = _columnas_por_nombre_normalizado(
        df,
        ["estatus", "estatus sae", "estatus saeplus"],
    )
    if not columnas_estatus:
        return df

    return df.rename(columns={columnas_estatus[0]: "estatus"})


def _columnas_esperadas(*grupos):
    columnas = []
    for grupo in grupos:
        columnas.extend(grupo)
    return {_normalizar_nombre_columna(columna) for columna in columnas}


def _leer_excel_columnas(archivo, columnas):
    esperadas = _columnas_esperadas(columnas)
    return leer_excel_seguro(
        archivo,
        usecols=lambda columna: _normalizar_nombre_columna(columna) in esperadas,
    )


def _leer_csv_columnas(archivo, columnas):
    esperadas = _columnas_esperadas(columnas)
    return pd.read_csv(
        archivo,
        dtype=object,
        low_memory=False,
        usecols=lambda columna: _normalizar_nombre_columna(columna) in esperadas,
    )


def _seleccionar_columnas(df, esquema, nombre_archivo):
    columnas = {}
    faltantes = []

    for destino, aliases in esquema.items():
        nombres = [destino, *aliases]
        encontradas = _columnas_por_nombre_normalizado(df, nombres)
        if not encontradas:
            faltantes.append(destino)
            continue
        columnas[destino] = df[encontradas[0]]

    if faltantes:
        raise ValueError(
            "El archivo "
            f"{nombre_archivo} no contiene las columnas requeridas: {', '.join(faltantes)}"
        )

    return pd.DataFrame(columnas, index=df.index)


def _ultimos_ocho_caracteres(serie):
    clave = serie.astype(str).str.strip().str[-8:]
    return clave.mask(serie.isna() | clave.eq(""))


def _normalizar_texto_comparacion(serie):
    return serie.astype(str).str.strip()


def _leer_cortes(archivo):
    df = _leer_excel_columnas(
        archivo,
        [COL_ABONADO, COL_EQUIPO_MAC, "observaciones"],
    )
    df = _seleccionar_columnas(
        df,
        {
            COL_ABONADO: [],
            COL_EQUIPO_MAC: [],
            "observaciones": [],
        },
        "de corte de abonados",
    )
    df[COL_EQUIPO_KEY] = _ultimos_ocho_caracteres(df[COL_EQUIPO_MAC])
    df = df.drop(columns=[COL_EQUIPO_MAC])
    return df[
        df["observaciones"].isna()
        & df[COL_ABONADO].notna()
        & df[COL_EQUIPO_KEY].notna()
    ]


def _leer_saeplus(archivo):
    df = _leer_excel_columnas(
        archivo,
        [
            COL_ABONADO,
            "documento",
            "nombre",
            "apellido",
            "estatus",
            "Estatus",
            "estatus sae",
            "estatus saeplus",
        ],
    )
    df = _normalizar_columna_estatus_saeplus(df)
    df = _seleccionar_columnas(
        df,
        {
            COL_ABONADO: [],
            "documento": [],
            "nombre": [],
            "apellido": [],
            "estatus": ["Estatus", "estatus sae", "estatus saeplus"],
        },
        "de abonados SAEPlus",
    )
    estado = _normalizar_texto_comparacion(df["estatus"])
    return df[df["estatus"].notna() & (estado != "ACTIVO") & df[COL_ABONADO].notna()]


def _leer_olt_activos(archivo):
    df = _leer_csv_columnas(
        archivo,
        ["SN", "olt", "catv", "administrative status", "status"],
    )
    df = _seleccionar_columnas(
        df,
        {
            "sn": ["SN"],
            "olt": [],
            "catv": [],
            "administrative status": [],
            "status": [],
        },
        "de SmartOLT",
    )
    df[COL_NSN] = _ultimos_ocho_caracteres(df["sn"])

    catv = _normalizar_texto_comparacion(df["catv"])
    status = _normalizar_texto_comparacion(df["status"])
    administrative_status = _normalizar_texto_comparacion(df["administrative status"])
    activos = (
        (status == "Online")
        | (catv != "Disabled")
        | (administrative_status == "Enabled")
    )

    return df[activos & df[COL_NSN].notna()]


def procesar_cortes(abonados_file, cortes_file, sae_file):
    df_cortes = _leer_cortes(abonados_file)
    df_saeplus = _leer_saeplus(sae_file)
    df_olt = _leer_olt_activos(cortes_file)

    resultado = pd.merge(df_cortes, df_saeplus, how="inner", on=COL_ABONADO)
    resultado = pd.merge(
        resultado,
        df_olt,
        how="inner",
        left_on=COL_EQUIPO_KEY,
        right_on=COL_NSN,
    )

    resultado_filtrado = resultado.rename(
        columns={
            COL_ABONADO: "n\u00b0 abonado",
            "documento": "documento_x",
            "nombre": "nombre_x",
            "apellido": "apellido_x",
        }
    )
    columnas_deseadas = [
        "n\u00b0 abonado",
        "documento_x",
        "nombre_x",
        "apellido_x",
        "estatus",
        "observaciones",
        "sn",
        "olt",
        "catv",
        "administrative status",
        "status",
    ]
    resultado_filtrado = resultado_filtrado[columnas_deseadas]

    return {
        "data": resultado_filtrado,
        "num_casos": resultado_filtrado.shape[0],
        "excel": excel_desde_dataframe(resultado_filtrado, "Resultado Filtrado"),
    }
