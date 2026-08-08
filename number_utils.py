"""Normalize numeric values: comma decimals and trim unnecessary trailing zeros."""

from __future__ import annotations

import numbers

import pandas as pd

# Used for matrix cost columns where Excel locale formatting is acceptable.
EXCEL_NUMBER_FORMAT = "[$-409]0.##"
# Used when values must always display with a dot decimal separator.
EXCEL_TEXT_FORMAT = "@"


def _is_empty_value(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() == ""


def normalize_numeric_value(value: object) -> object:
    """
    Normalize numeric-like values:
      70,00 -> 70
      80,04 -> 80.04
      7.1900 -> 7.19
    Non-numeric values are returned unchanged.
    """
    if _is_empty_value(value):
        return None

    if isinstance(value, bool):
        return value

    if isinstance(value, numbers.Integral) and not isinstance(value, bool):
        return int(value)

    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        numeric = float(value)
    else:
        text = str(value).strip()
        if not text:
            return None

        cleaned = text.replace(" ", "")
        if "," in cleaned and "." in cleaned:
            cleaned = cleaned.replace(".", "").replace(",", ".")
        elif "," in cleaned:
            cleaned = cleaned.replace(",", ".")

        parsed = pd.to_numeric(cleaned, errors="coerce")
        if pd.isna(parsed):
            return value
        numeric = float(parsed)

    rounded = round(numeric, 10)
    if rounded == int(rounded):
        return int(rounded)
    return rounded


def format_number_as_dot_text(value: object) -> str | None:
    """
    Render a normalized number as text with dot decimals.

    Examples:
      44.96 -> "44.96"
      375   -> "375"
      "44,96" -> "44.96"
    """
    normalized = normalize_numeric_value(value)
    if normalized is None:
        return None
    if isinstance(normalized, int):
        return str(normalized)
    text = f"{float(normalized):.10f}".rstrip("0").rstrip(".")
    return text or "0"


def normalize_numeric_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    """Normalize selected columns and keep ints as ints via object dtype."""
    if df.empty:
        return df

    normalized = df.copy()
    for column in columns:
        if column not in normalized.columns:
            continue
        values = [normalize_numeric_value(value) for value in normalized[column]]
        normalized[column] = pd.Series(values, dtype=object)
    return normalized


def normalize_dataframe_numbers(df: pd.DataFrame) -> pd.DataFrame:
    """Apply numeric normalization to every cell in the dataframe."""
    if df.empty:
        return df

    normalized = df.copy()
    for column in normalized.columns:
        values = [normalize_numeric_value(value) for value in normalized[column]]
        normalized[column] = pd.Series(values, dtype=object)
    return normalized
