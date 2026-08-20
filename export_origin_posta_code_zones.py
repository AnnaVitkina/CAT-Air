"""Export Origin Postal Code Zone catalog from matrix lane data."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

POSTAL_ZONE_COLUMNS = ("Name", "Country", "Postal Code", "Excluded")
POSTAL_CODE_ZONES_SUFFIX = "_origin_postal_code_zones.txt"

CARRIER_KEY_PATTERN = re.compile(r"KHNN|DHLC", re.IGNORECASE)
ZONE_CODE_PATTERN = re.compile(r"^([A-Za-z]{2})([\d][\d\-]*)$")
EXCLUSION_NAME_PATTERN = re.compile(r"^(.+?) excl\. (.+)$")
COUNTRY_PREFIX_PATTERN = re.compile(r"^[A-Za-z]{2}(.+)$")


@dataclass(frozen=True)
class OriginPostalCodeZone:
    name: str
    country: str
    postal_code: str
    excluded: str = ""

    def as_txt_row(self) -> str:
        return "\t".join([self.name, self.country, self.postal_code, self.excluded])


@dataclass
class OriginPostalCodeZoneExportResult:
    output_path: Path
    zone_count: int


def _cell_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _is_empty_value(value: object) -> bool:
    return _cell_text(value) == ""


def format_date_dd_mm_yyyy(value: object) -> str | None:
    if _is_empty_value(value):
        return None
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None
    return parsed.strftime("%d.%m.%Y")


def trim_location_port(value: object) -> str:
    text = _cell_text(value)
    if not text or "," in text:
        return ""
    match = COUNTRY_PREFIX_PATTERN.match(text)
    if not match:
        return ""
    return match.group(1)


def key_requires_origin_postal_code_zone(key: object) -> bool:
    text = _cell_text(key)
    if not text:
        return False
    return bool(CARRIER_KEY_PATTERN.search(text))


def base_origin_postal_code_zone_value(row: pd.Series) -> str | None:
    if not key_requires_origin_postal_code_zone(row.get("KEY")):
        return None
    value = row.get("ORIGIN_LOCATION_NAME__C")
    text = _cell_text(value)
    return text or None


def build_shipment_comparison_signature(row: pd.Series) -> tuple[object, ...]:
    """Shipment fields used to detect prefix-zone conflicts (matches matrix columns)."""
    return (
        _cell_text(row.get("COMMODITY")),
        _cell_text(row.get("ORIGIN_COUNTRY__C")),
        trim_location_port(row.get("DESTINATION_LOCATION_NAME__C")),
        _cell_text(row.get("DESTINATION_COUNTRY__C")),
        _cell_text(row.get("SERVICE_LEVEL")),
        _cell_text(row.get("SERVICE")),
        format_date_dd_mm_yyyy(row.get("RATE_EFFECTIVE_DATE__C")),
        format_date_dd_mm_yyyy(row.get("RATE_EXPIRATION_DATE__C")),
    )


def parse_origin_postal_code_zone(name: str) -> OriginPostalCodeZone | None:
    text = str(name).strip()
    if not text:
        return None

    exclusion_match = EXCLUSION_NAME_PATTERN.match(text)
    if exclusion_match:
        base_name = exclusion_match.group(1).strip()
        excluded = exclusion_match.group(2).strip()
        base_parsed = parse_origin_postal_code_zone(base_name)
        if base_parsed is None:
            return None
        return OriginPostalCodeZone(
            name=text,
            country=base_parsed.country,
            postal_code=base_parsed.postal_code,
            excluded=excluded,
        )

    if "," in text:
        city_part, country_part = text.rsplit(",", 1)
        country = country_part.strip().upper()
        if len(country) == 2 and city_part.strip():
            return OriginPostalCodeZone(name=text, country=country, postal_code="")
        return None

    match = ZONE_CODE_PATTERN.match(text)
    if not match:
        return None

    return OriginPostalCodeZone(
        name=text,
        country=match.group(1).upper(),
        postal_code=match.group(2),
    )


def find_prefix_zone_pairs(zone_names: list[str]) -> list[tuple[str, str]]:
    parsed = {name: parse_origin_postal_code_zone(name) for name in zone_names}
    pairs: list[tuple[str, str]] = []

    for short_name in zone_names:
        short_parsed = parsed.get(short_name)
        if short_parsed is None:
            continue
        for long_name in zone_names:
            if short_name == long_name:
                continue
            long_parsed = parsed.get(long_name)
            if long_parsed is None:
                continue
            if short_parsed.country != long_parsed.country:
                continue
            if long_name.startswith(short_name):
                pairs.append((short_name, long_name))

    return pairs


class OriginPostalCodeZoneResolver:
    """Resolve matrix zone names, applying prefix exclusions when shipment rows match."""

    def __init__(self, rate_card: pd.DataFrame) -> None:
        self._zones_by_index: dict[object, str | None] = {}
        self._build(rate_card)

    def resolve(self, row_index: object) -> str | None:
        return self._zones_by_index.get(row_index)

    def resolved_zone_names(self) -> list[str]:
        unique: list[str] = []
        seen: set[str] = set()
        for value in self._zones_by_index.values():
            if value is None:
                continue
            key = value.casefold()
            if key in seen:
                continue
            seen.add(key)
            unique.append(value)
        return sorted(unique, key=str.casefold)

    def _build(self, rate_card: pd.DataFrame) -> None:
        row_entries: list[tuple[object, str, tuple[object, ...]]] = []
        zone_signatures: dict[str, set[tuple[object, ...]]] = {}

        for row_index, row in rate_card.iterrows():
            base_zone = base_origin_postal_code_zone_value(row)
            if base_zone is None:
                self._zones_by_index[row_index] = None
                continue

            signature = build_shipment_comparison_signature(row)
            row_entries.append((row_index, base_zone, signature))
            zone_signatures.setdefault(base_zone, set()).add(signature)

        prefix_pairs = find_prefix_zone_pairs(list(zone_signatures.keys()))
        longer_zones_by_short: dict[str, list[str]] = {}
        for short_zone, long_zone in prefix_pairs:
            longer_zones_by_short.setdefault(short_zone, []).append(long_zone)
        for short_zone in longer_zones_by_short:
            longer_zones_by_short[short_zone].sort(key=lambda name: (-len(name), name.casefold()))

        for row_index, base_zone, signature in row_entries:
            resolved = base_zone
            matching_long_zones = [
                long_zone
                for long_zone in longer_zones_by_short.get(base_zone, [])
                if signature in zone_signatures.get(long_zone, set())
            ]
            if matching_long_zones:
                best_long_zone = max(matching_long_zones, key=len)
                long_parsed = parse_origin_postal_code_zone(best_long_zone)
                if long_parsed is not None:
                    resolved = f"{base_zone} excl. {long_parsed.postal_code}"

            self._zones_by_index[row_index] = resolved


def origin_postal_code_zone_value(row: pd.Series, resolver: OriginPostalCodeZoneResolver | None = None) -> str | None:
    if resolver is not None:
        return resolver.resolve(row.name)
    return base_origin_postal_code_zone_value(row)


def collect_origin_postal_code_zones(resolver: OriginPostalCodeZoneResolver) -> list[OriginPostalCodeZone]:
    zones: list[OriginPostalCodeZone] = []
    for name in resolver.resolved_zone_names():
        parsed = parse_origin_postal_code_zone(name)
        if parsed is not None:
            zones.append(parsed)
    return zones


def default_postal_zones_path(matrix_output_path: Path) -> Path:
    stem = matrix_output_path.stem
    if stem.endswith("_matrix"):
        stem = stem[: -len("_matrix")]
    return matrix_output_path.with_name(f"{stem}{POSTAL_CODE_ZONES_SUFFIX}")


def write_origin_postal_code_zones_txt(
    zones: list[OriginPostalCodeZone],
    output_path: Path,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["\t".join(POSTAL_ZONE_COLUMNS)]
    lines.extend(zone.as_txt_row() for zone in zones)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return output_path


def export_origin_postal_code_zones(
    rate_card: pd.DataFrame,
    output_path: Path,
    *,
    resolver: OriginPostalCodeZoneResolver | None = None,
) -> OriginPostalCodeZoneExportResult | None:
    zone_resolver = resolver or OriginPostalCodeZoneResolver(rate_card)
    zones = collect_origin_postal_code_zones(zone_resolver)
    if not zones:
        return None

    txt_path = write_origin_postal_code_zones_txt(zones, output_path)
    return OriginPostalCodeZoneExportResult(
        output_path=txt_path,
        zone_count=len(zones),
    )
