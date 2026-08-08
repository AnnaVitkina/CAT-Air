"""Build Accessorial cost tab from DELIVERY_CHARGES processing data."""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from number_utils import (
    EXCEL_TEXT_FORMAT,
    normalize_dataframe_numbers,
    normalize_numeric_columns,
    normalize_numeric_value,
    format_number_as_dot_text,
)

DELIVERY_CHARGES_SHEET = "DELIVERY_CHARGES"
ACCESSORIAL_SHEET_NAME = "Accessorial cost"
ACCESSORIAL_COST_NAME = "Delivery"

RATE_COLUMN = "Delivery Rate"
MIN_COLUMN = "Delivery Min"
MAX_COLUMN = "Delivery Max"
CURRENCY_COLUMN = "Delivery Currency"
ORIGIN_COUNTRY_COLUMN = "ORIGIN_COUNTRY__C"
DESTINATION_COUNTRY_COLUMN = "DESTINATION_COUNTRY__C"
ORIGIN_LOCATION_COLUMN = "ORIGIN_LOCATION_NAME__C"
DESTINATION_LOCATION_COLUMN = "DESTINATION_LOCATION_NAME__C"
VALID_FROM_COLUMN = "RATE_EFFECTIVE_DATE__C"
VALID_TO_COLUMN = "RATE_EXPIRATION_DATE__C"

ACCESSORIAL_COLUMNS = (
    "Cost Name",
    "Currency",
    "Min",
    "p/unit",
    "Max",
    "Applies if",
    "Valid from",
    "Valid to",
)

NUMERIC_ACCESSORIAL_COLUMNS = frozenset({"Min", "p/unit", "Max"})
COUNTRY_PREFIX_PATTERN = re.compile(r"^[A-Za-z]{2}(.+)$")


@dataclass(frozen=True)
class RateKey:
    rate: object
    minimum: object
    maximum: object


@dataclass(frozen=True)
class LaneIdentity:
    destination_country: str
    origin_country: str = ""
    origin_port: str = ""
    destination_port: str = ""


@dataclass
class AccessorialBuildResult:
    row_count: int
    sheet_name: str = ACCESSORIAL_SHEET_NAME


def _cell_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _is_empty_value(value: object) -> bool:
    return _cell_text(value) == ""


def trim_location_port(value: object) -> str:
    text = _cell_text(value)
    if not text or "," in text:
        return ""
    match = COUNTRY_PREFIX_PATTERN.match(text)
    if not match:
        return ""
    return match.group(1)


def format_date_dd_mm_yyyy(value: object) -> str | None:
    if _is_empty_value(value):
        return None

    if hasattr(value, "strftime"):
        return value.strftime("%d.%m.%Y")

    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return parsed.strftime("%d.%m.%Y")


def _rate_key_from_row(row: pd.Series) -> RateKey:
    return RateKey(
        rate=normalize_numeric_value(row.get(RATE_COLUMN)),
        minimum=normalize_numeric_value(row.get(MIN_COLUMN)),
        maximum=normalize_numeric_value(row.get(MAX_COLUMN)),
    )


def _lane_identity_from_row(row: pd.Series) -> LaneIdentity:
    return LaneIdentity(
        destination_country=_cell_text(row.get(DESTINATION_COUNTRY_COLUMN)),
        origin_country=_cell_text(row.get(ORIGIN_COUNTRY_COLUMN)),
        origin_port=trim_location_port(row.get(ORIGIN_LOCATION_COLUMN)),
        destination_port=trim_location_port(row.get(DESTINATION_LOCATION_COLUMN)),
    )


def _format_equals_clause(label: str, value: str) -> str:
    return f'{label} equals "{value}"'


def _build_lane_clause(lane: LaneIdentity, *, include_origin: bool, include_ports: bool) -> str:
    parts: list[str] = []

    if include_origin and lane.origin_country:
        parts.append(_format_equals_clause("Origin country", lane.origin_country))

    if lane.destination_country:
        parts.append(_format_equals_clause("Destination country", lane.destination_country))

    if include_ports:
        if lane.origin_port:
            parts.append(_format_equals_clause("Origin Port", lane.origin_port))
        if lane.destination_port:
            parts.append(_format_equals_clause("Destination Port", lane.destination_port))

    return " and ".join(parts)


@dataclass
class AppliesIfContext:
    destination_rate_keys: dict[str, set[RateKey]]
    origin_destination_rate_keys: dict[tuple[str, str], set[RateKey]]
    full_lane_rate_keys: dict[LaneIdentity, set[RateKey]]


def _build_applies_if_context(df: pd.DataFrame) -> AppliesIfContext:
    destination_rate_keys: dict[str, set[RateKey]] = {}
    origin_destination_rate_keys: dict[tuple[str, str], set[RateKey]] = {}
    full_lane_rate_keys: dict[LaneIdentity, set[RateKey]] = {}

    for _, row in df.iterrows():
        rate_key = _rate_key_from_row(row)
        lane = _lane_identity_from_row(row)

        if lane.destination_country:
            destination_rate_keys.setdefault(lane.destination_country, set()).add(rate_key)

        if lane.origin_country and lane.destination_country:
            origin_destination_rate_keys.setdefault(
                (lane.origin_country, lane.destination_country),
                set(),
            ).add(rate_key)

        if lane.destination_country:
            full_lane_rate_keys.setdefault(lane, set()).add(rate_key)

    return AppliesIfContext(
        destination_rate_keys=destination_rate_keys,
        origin_destination_rate_keys=origin_destination_rate_keys,
        full_lane_rate_keys=full_lane_rate_keys,
    )


def _clause_specificity_level(
    lane: LaneIdentity,
    rate_key: RateKey,
    context: AppliesIfContext,
) -> tuple[bool, bool]:
    destination_only = (
        bool(lane.destination_country)
        and len(context.destination_rate_keys.get(lane.destination_country, set())) == 1
    )
    if destination_only:
        return False, False

    origin_destination_key = (lane.origin_country, lane.destination_country)
    origin_destination_only = (
        bool(lane.origin_country)
        and bool(lane.destination_country)
        and len(context.origin_destination_rate_keys.get(origin_destination_key, set())) == 1
    )
    if origin_destination_only:
        return True, False

    return True, True


def _build_applies_if(
    group: pd.DataFrame,
    rate_key: RateKey,
    context: AppliesIfContext,
) -> str:
    lanes = sorted(
        {_lane_identity_from_row(row) for _, row in group.iterrows()},
        key=lambda lane: (
            lane.destination_country,
            lane.origin_country,
            lane.origin_port,
            lane.destination_port,
        ),
    )
    lanes = [lane for lane in lanes if lane.destination_country]
    if not lanes:
        return ""

    clauses: list[str] = []
    for lane in lanes:
        include_origin, include_ports = _clause_specificity_level(lane, rate_key, context)
        clause = _build_lane_clause(
            lane,
            include_origin=include_origin,
            include_ports=include_ports,
        )
        if clause and clause not in clauses:
            clauses.append(clause)

    return " or ".join(clauses)


def build_accessorial_costs_from_delivery_charges(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=ACCESSORIAL_COLUMNS)

    df = normalize_dataframe_numbers(df)

    required_columns = {
        RATE_COLUMN,
        CURRENCY_COLUMN,
        DESTINATION_COUNTRY_COLUMN,
        VALID_FROM_COLUMN,
        VALID_TO_COLUMN,
    }
    missing = sorted(required_columns - set(df.columns))
    if missing:
        raise ValueError(f"DELIVERY_CHARGES is missing required columns: {', '.join(missing)}")

    context = _build_applies_if_context(df)
    rows: list[dict[str, object]] = []

    grouped = df.groupby(
        [RATE_COLUMN, MIN_COLUMN, MAX_COLUMN],
        dropna=False,
        sort=True,
    )

    for (rate_value, min_value, max_value), group in grouped:
        first_row = group.iloc[0]
        rate_key = RateKey(
            rate=normalize_numeric_value(rate_value),
            minimum=normalize_numeric_value(min_value),
            maximum=normalize_numeric_value(max_value),
        )
        rows.append(
            {
                "Cost Name": ACCESSORIAL_COST_NAME,
                "Currency": first_row.get(CURRENCY_COLUMN),
                "Min": normalize_numeric_value(min_value),
                "p/unit": normalize_numeric_value(rate_value),
                "Max": normalize_numeric_value(max_value),
                "Applies if": _build_applies_if(group, rate_key, context),
                "Valid from": format_date_dd_mm_yyyy(first_row.get(VALID_FROM_COLUMN)),
                "Valid to": format_date_dd_mm_yyyy(first_row.get(VALID_TO_COLUMN)),
            }
        )

    result = pd.DataFrame(rows, columns=ACCESSORIAL_COLUMNS)
    return normalize_numeric_columns(result, list(NUMERIC_ACCESSORIAL_COLUMNS))
