import io

import pandas as pd


def validar_columnas(df, columnas_requeridas, nombre_archivo):
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if faltantes:
        raise ValueError(
            f"El archivo {nombre_archivo} no contiene las columnas requeridas: {', '.join(faltantes)}"
        )


def excel_desde_dataframe(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


def excel_desde_hojas(hojas):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in hojas:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output
