import io
import unicodedata
from collections import Counter
from numbers import Real

import pandas as pd

from funciones import leer_excel_seguro


PROCESSING_YEAR = 2026
SOURCE_ROW_ORDER_COLUMN = "_orden_fila_origen"

REQUIRED_COLUMNS = (
    "DÍA",
    "MES",
    "SOLICITANTE",
    "ABONADO",
    "SERVICIO",
    "FALLA",
    "OBSERVACIONES",
    "ESTATUS",
    "DÍA2",
    "MES3",
    "ESTATUS 2",
    "INGENIERO",
    "OBSERVACIONES2",
)

OPTIONAL_COLUMNS = {"HORA4"}

CALCULATED_COLUMNS = (
    "Fecha",
    "Visita",
    "Es_Recurrencia",
    "Total_Visitas",
    "Primer_Ingeniero",
    "Ultimo_Ingeniero",
    "Cambio_Ingeniero",
    "Ingenieros_Diferentes",
    "Dias_Entre_Visitas",
    "Primera_Visita",
    "Ultima_Visita",
    "Recurrencia_Critica",
)

MONTH_ALIASES = {
    "ENERO": 1,
    "ENE": 1,
    "FEBRERO": 2,
    "FEB": 2,
    "MARZO": 3,
    "MAR": 3,
    "ABRIL": 4,
    "ABR": 4,
    "MAYO": 5,
    "MAY": 5,
    "JUNIO": 6,
    "JUN": 6,
    "JULIO": 7,
    "JUL": 7,
    "AGOSTO": 8,
    "AGO": 8,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "SEP": 9,
    "SEPT": 9,
    "SET": 9,
    "OCTUBRE": 10,
    "OCT": 10,
    "NOVIEMBRE": 11,
    "NOV": 11,
    "DICIEMBRE": 12,
    "DIC": 12,
}


class MissingColumnsError(ValueError):
    """Error raised when the uploaded file does not contain required columns."""

    def __init__(self, missing_columns: list[str]) -> None:
        self.missing_columns = missing_columns
        super().__init__(
            f"Faltan columnas obligatorias: {', '.join(missing_columns)}"
        )


def clean_column_name(column: object) -> str:
    """Trim extra spaces and normalize an Excel header to uppercase."""
    return " ".join(str(column).strip().split()).upper()


def remove_accents(value: str) -> str:
    """Return text without accent marks."""
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(
        character for character in normalized if not unicodedata.combining(character)
    )


def matching_key(value: str) -> str:
    """Build an accent-insensitive key used to match required headers."""
    return remove_accents(value).upper()


def find_duplicates(values: list[str]) -> list[str]:
    """Return duplicated values in deterministic order."""
    counter = Counter(values)
    return sorted(value for value, count in counter.items() if count > 1)


def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """Clean headers and map required columns to their canonical names."""
    normalized_df = df.copy()
    normalized_df.columns = [clean_column_name(column) for column in normalized_df.columns]

    duplicates = find_duplicates(list(normalized_df.columns))
    if duplicates:
        raise ValueError(
            "Hay columnas duplicadas después de limpiar encabezados: "
            f"{', '.join(duplicates)}"
        )

    required_by_key = {matching_key(column): column for column in REQUIRED_COLUMNS}
    rename_map = {
        column: required_by_key[matching_key(column)]
        for column in normalized_df.columns
        if matching_key(column) in required_by_key
    }
    normalized_df = normalized_df.rename(columns=rename_map)

    duplicates = find_duplicates(list(normalized_df.columns))
    if duplicates:
        raise ValueError(
            "Hay columnas duplicadas después de normalizar encabezados: "
            f"{', '.join(duplicates)}"
        )

    return normalized_df


def validate_required_columns(df: pd.DataFrame) -> None:
    """Validate that every required source column exists."""
    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in df.columns and column not in OPTIONAL_COLUMNS
    ]
    if missing_columns:
        raise MissingColumnsError(missing_columns)


def clean_and_validate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize headers and stop the process when required columns are missing."""
    normalized_df = normalize_column_names(df)
    validate_required_columns(normalized_df)
    return normalized_df


def clean_text_value(value: object) -> object:
    """Trim spaces and normalize text values to uppercase."""
    if pd.isna(value):
        return pd.NA

    if isinstance(value, Real) and not isinstance(value, bool):
        return str(int(value)) if float(value).is_integer() else str(value).strip()

    cleaned = " ".join(str(value).strip().split())
    if not cleaned:
        return pd.NA

    return cleaned.upper()


def clean_values(df: pd.DataFrame) -> pd.DataFrame:
    """Remove unnecessary spaces and normalize text-like values."""
    cleaned_df = df.copy()
    for column in cleaned_df.columns:
        if (
            pd.api.types.is_object_dtype(cleaned_df[column])
            or pd.api.types.is_string_dtype(cleaned_df[column])
            or column in {"ABONADO", "INGENIERO"}
        ):
            cleaned_df[column] = cleaned_df[column].map(clean_text_value)

    return cleaned_df


def parse_numeric_month(value: object) -> int | None:
    """Parse month values represented as numbers."""
    if pd.isna(value):
        return None

    if isinstance(value, Real) and not isinstance(value, bool):
        month = int(value)
        return month if 1 <= month <= 12 and float(value).is_integer() else None

    try:
        number = float(str(value).strip())
    except ValueError:
        return None

    month = int(number)
    return month if 1 <= month <= 12 and number.is_integer() else None


def parse_month(value: object) -> int | None:
    """Parse month values from numbers or Spanish month names."""
    numeric_month = parse_numeric_month(value)
    if numeric_month is not None:
        return numeric_month

    if pd.isna(value):
        return None

    normalized_text = remove_accents(" ".join(str(value).strip().split()).upper())
    return MONTH_ALIASES.get(normalized_text)


def build_fecha_column(df: pd.DataFrame, year: int) -> pd.Series:
    """Build a Fecha column from DÍA, MES and the configured year."""
    date_parts = pd.DataFrame(
        {
            "year": year,
            "month": df["MES"].map(parse_month),
            "day": pd.to_numeric(df["DÍA"], errors="coerce"),
        }
    )
    return pd.to_datetime(date_parts, errors="coerce")


def sort_for_visit_order(df: pd.DataFrame) -> pd.DataFrame:
    """Sort records by subscriber, visit date and original row order."""
    return (
        df.sort_values(
            by=["ABONADO", "Fecha", SOURCE_ROW_ORDER_COLUMN],
            kind="mergesort",
            na_position="last",
        )
        .reset_index(drop=True)
    )


def add_visit_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add all recurrence-related calculated columns."""
    result_df = df.copy()
    grouped = result_df.groupby("ABONADO", dropna=False, sort=False)

    result_df["Visita"] = grouped.cumcount().add(1).astype("int64")
    result_df["Total_Visitas"] = grouped["Fecha"].transform("size").astype("int64")
    result_df["Es_Recurrencia"] = "No"
    result_df.loc[result_df["Visita"] > 1, "Es_Recurrencia"] = "Sí"
    result_df["Primer_Ingeniero"] = grouped["INGENIERO"].transform(
        lambda values: values.iloc[0]
    )
    result_df["Ultimo_Ingeniero"] = grouped["INGENIERO"].transform(
        lambda values: values.iloc[-1]
    )

    previous_engineer = grouped["INGENIERO"].shift(1)
    changed_engineer = (
        (result_df["Visita"] > 1)
        & (result_df["INGENIERO"].fillna("") != previous_engineer.fillna(""))
    )
    result_df["Cambio_Ingeniero"] = "No"
    result_df.loc[changed_engineer, "Cambio_Ingeniero"] = "Sí"
    result_df["Ingenieros_Diferentes"] = grouped["INGENIERO"].transform(
        lambda values: values.dropna().nunique()
    ).astype("int64")

    previous_date = grouped["Fecha"].shift(1)
    result_df["Dias_Entre_Visitas"] = (
        (result_df["Fecha"] - previous_date).dt.days.fillna(0).astype("int64")
    )
    result_df["Primera_Visita"] = result_df["Visita"] == 1
    result_df["Ultima_Visita"] = result_df["Visita"] == result_df["Total_Visitas"]
    result_df["Recurrencia_Critica"] = result_df["Total_Visitas"] >= 3

    return result_df


def transform_orders(df: pd.DataFrame, year: int = PROCESSING_YEAR) -> pd.DataFrame:
    """Run the full transformation required for the recurrence report."""
    transformed_df = clean_values(df)
    original_columns = list(transformed_df.columns)
    transformed_df[SOURCE_ROW_ORDER_COLUMN] = range(len(transformed_df))
    transformed_df["Fecha"] = build_fecha_column(transformed_df, year)
    transformed_df = sort_for_visit_order(transformed_df)
    transformed_df = add_visit_columns(transformed_df)
    transformed_df = transformed_df.drop(columns=[SOURCE_ROW_ORDER_COLUMN])

    return transformed_df[original_columns + list(CALCULATED_COLUMNS)]


def calculate_summary(df: pd.DataFrame) -> dict[str, int]:
    """Build summary metrics for the result page."""
    return {
        "Total registros": int(df.shape[0]),
        "Abonados": int(df["ABONADO"].nunique(dropna=True)),
        "Visitas recurrentes": int((df["Visita"] > 1).sum()),
        "Abonados críticos": int(
            df.loc[df["Recurrencia_Critica"], "ABONADO"].nunique(dropna=True)
        ),
        "Cambios de ingeniero": int((df["Cambio_Ingeniero"] == "Sí").sum()),
        "Fechas inválidas": int(df["Fecha"].isna().sum()),
    }


def build_excel_output(df: pd.DataFrame) -> io.BytesIO:
    """Export the transformed DataFrame to an in-memory Excel workbook."""
    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        datetime_format="yyyy-mm-dd",
        date_format="yyyy-mm-dd",
    ) as writer:
        sheet_name = "Recurrencias"
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "fg_color": "#D9EAF7",
                "border": 1,
            }
        )
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})

        for column_index, column_name in enumerate(df.columns):
            worksheet.write(0, column_index, column_name, header_format)
            max_length = max(
                [len(str(column_name))]
                + df[column_name].fillna("").astype(str).map(len).tolist()
            )
            width = min(max(max_length + 2, 10), 45)
            worksheet.set_column(column_index, column_index, width)

        if "Fecha" in df.columns:
            fecha_index = df.columns.get_loc("Fecha")
            worksheet.set_column(fecha_index, fecha_index, 12, date_format)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df), max(len(df.columns) - 1, 0))

    output.seek(0)
    return output


def procesar_recurrencias(archivo_excel, year: int = PROCESSING_YEAR) -> dict[str, object]:
    """Process an uploaded service-order Excel file for Power BI consumption."""
    df = leer_excel_seguro(archivo_excel, dtype=object)
    validated_df = clean_and_validate_columns(df)
    transformed_df = transform_orders(validated_df, year=year)
    summary = calculate_summary(transformed_df)

    return {
        "data": transformed_df.head(200),
        "columns": transformed_df.columns.tolist(),
        "num_casos": summary["Total registros"],
        "summary": summary,
        "excel": build_excel_output(transformed_df),
    }
