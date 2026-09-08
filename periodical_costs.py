"""Apply periodical cost columns to an RA matrix during update."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re

import pandas as pd

from build_matrix import CostBlock, discover_cost_blocks, resolve_carrier_workbook_sheet
from convert_to_processing import (
    _cell_text,
    convert_workbook_to_dataframes,
    save_dataframes,
)
from matrix_io import (
    MatrixLane,
    MatrixWorkbookData,
    RATE_CARD_SHEET_NAME,
    cell_text,
    find_cost_block_columns,
    format_date_dd_mm_yyyy,
    group_lanes_by_key,
    insert_cost_block_after,
    lane_value,
    list_excel_files,
    parse_date_dd_mm_yyyy,
    prompt_file_selection,
    read_matrix_workbook,
    write_matrix_workbook,
)
from project_paths import OUTPUT_DIR, PERIODICAL_DIR, PREVIOUS_RA_DIR, PROCESSING_DIR, ensure_workspace_dirs

PERIODICAL_SUB_HEADERS = ("Currency", "p/unit")
CARRIER_PATTERN = __import__("re").compile(r"\b(?:Air\s+)?(DHL|KN)\b", __import__("re").IGNORECASE)


@dataclass(frozen=True)
class PeriodicalCostKind:
    key: str
    anchor_cost_name: str
    display_base_name: str
    source_match: str

    def matches_source_block(self, block_name: str) -> bool:
        return self.source_match in block_name.lower()


CAPACITY_KIND = PeriodicalCostKind(
    key="capacity",
    anchor_cost_name="Base Rate",
    display_base_name="Base Rate Capacity",
    source_match="capacity",
)
ADDER_KIND = PeriodicalCostKind(
    key="adder",
    anchor_cost_name="Fuel Surcharge",
    display_base_name="Fuel Surcharge Adder",
    source_match="adder",
)
PERIODICAL_COST_KINDS = (CAPACITY_KIND, ADDER_KIND)


@dataclass(frozen=True)
class PeriodicalLaneCost:
    key: str
    valid_from: date
    valid_to: date
    kind: PeriodicalCostKind
    values: dict[str, object]


@dataclass(frozen=True)
class PeriodicalRaResult:
    previous_ra_file: Path
    periodical_file: Path
    output_file: Path
    matching_key_count: int
    updated_row_count: int
    updated_cell_count: int


def detect_carrier_from_filename(filename: str) -> str:
    match = CARRIER_PATTERN.search(filename)
    if not match:
        raise ValueError(f"Could not detect carrier (DHL/KN) from filename: {filename}")
    return match.group(1).upper()


def prompt_periodical_file_selection(
    files: list[Path],
    *,
    directory: Path,
    auto: bool,
    required: bool = False,
) -> Path | None:
    if not files:
        if required:
            raise FileNotFoundError(f"No periodical files found in: {directory}")
        print("\nNo periodical files found - skipping periodical costs.")
        return None

    if auto:
        print(f"\nAuto mode: using periodical file {files[0].name}")
        return files[0]

    print("\nSelect periodical costs file:")
    print(f"  Folder: {directory}")
    if not required:
        print("  0. Skip periodical costs")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")

    while True:
        prompt = "Enter file number"
        if required:
            prompt += f" (1-{len(files)})"
        else:
            prompt += " (0 to skip)"
        raw = input(f"{prompt}: ").strip()
        if not required and raw in {"", "0"}:
            return None
        try:
            choice = int(raw)
            if 1 <= choice <= len(files):
                return files[choice - 1]
        except ValueError:
            pass
        if required:
            print(f"Please enter a number between 1 and {len(files)}.")
        else:
            print(f"Please enter a number between 0 and {len(files)}.")


def _detect_periodical_header_row(file_path: Path, sheet_name: str) -> int:
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=2)
    if preview.shape[0] < 2:
        return 0

    row1_values = [_cell_text(value) for value in preview.iloc[1]]
    currency_headers = sum(1 for value in row1_values if value.endswith(" Currency"))
    if currency_headers >= 2:
        return 1
    return 0


def convert_periodical_file(periodical_file: Path, carrier: str) -> tuple[Path, str]:
    workbook = pd.ExcelFile(periodical_file)
    sheet_name = resolve_carrier_workbook_sheet(workbook.sheet_names, carrier)
    header_row = _detect_periodical_header_row(periodical_file, sheet_name)

    print(f"\nConverting periodical file '{periodical_file.name}' (tab: {sheet_name})...")
    dataframes = convert_workbook_to_dataframes(
        periodical_file,
        [sheet_name],
        header_row=header_row,
    )
    output_path = PROCESSING_DIR / f"{periodical_file.stem}_periodical.xlsx"
    save_dataframes(dataframes, output_path)
    print(f"Saved periodical processing workbook to: {output_path}")
    return output_path, sheet_name


def _parse_rate_card_date(value: object) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = cell_text(value)
    if not text:
        return None

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return datetime.strptime(text, "%Y-%m-%d").date()

    parsed = parse_date_dd_mm_yyyy(text)
    if parsed is not None:
        return parsed

    converted = pd.to_datetime(value, errors="coerce", dayfirst=False)
    if pd.isna(converted):
        return None
    return converted.date()


def _format_validity_label(valid_from: date, valid_to: date) -> str:
    return f"{format_date_dd_mm_yyyy(valid_from)}-{format_date_dd_mm_yyyy(valid_to)}"


def _build_display_cost_name(kind: PeriodicalCostKind, valid_from: date, valid_to: date) -> str:
    return f"{kind.display_base_name} ({_format_validity_label(valid_from, valid_to)})"


def _extract_block_values(row: pd.Series, block: CostBlock) -> dict[str, object]:
    values: dict[str, object] = {"Currency": row.get(block.currency_column)}
    for value_column in block.value_columns:
        values[value_column.header] = row.get(value_column.source_column)
    if values.get("Flat") is not None and values.get("p/unit") is None:
        values["p/unit"] = values["Flat"]
    return values


def _discover_periodical_blocks(columns: list[str]) -> dict[str, CostBlock]:
    discovered: dict[str, CostBlock] = {}
    for block in discover_cost_blocks(columns):
        for kind in PERIODICAL_COST_KINDS:
            if kind.matches_source_block(block.cost_name):
                discovered[kind.key] = block
    return discovered


def _periods_overlap(
    lane_start: date,
    lane_end: date,
    period_start: date,
    period_end: date,
) -> bool:
    return lane_start <= period_end and period_start <= lane_end


def _build_header_cells(
    matrix: MatrixWorkbookData,
    kind: PeriodicalCostKind,
    cost_name: str,
    valid_from: date,
    valid_to: date,
) -> list[list[object]]:
    validity_text = (
        f"Validity period: from {format_date_dd_mm_yyyy(valid_from)} "
        f"to {format_date_dd_mm_yyyy(valid_to)}\r\n\r\n\r\nApplies if invoiced by Carrier"
    )
    if kind.key == "capacity":
        rate_by = "Rate by: Weight/chargeable kg\r\nDirect rule"
        row6 = ["", "<= 10000"]
    else:
        rate_by = "Rate by: Weight/kg\r\nRegular rule"
        row6 = ["", ""]

    sub_count = len(PERIODICAL_SUB_HEADERS)
    rows: list[list[object]] = []
    for row_index in range(len(matrix.header_rows)):
        excel_row = row_index + 1
        if excel_row <= matrix.header_row - 5:
            rows.append([None] * sub_count)
        elif excel_row == matrix.header_row - 4:
            rows.append([cost_name, None])
        elif excel_row == matrix.header_row - 3:
            rows.append([validity_text, None])
        elif excel_row == matrix.header_row - 2:
            rows.append([rate_by, None])
        elif excel_row == matrix.header_row - 1:
            rows.append(list(row6))
        else:
            rows.append(list(PERIODICAL_SUB_HEADERS))
    return rows


def parse_periodical_lane_costs(rate_card: pd.DataFrame) -> list[PeriodicalLaneCost]:
    blocks = _discover_periodical_blocks(list(rate_card.columns))
    if not blocks:
        raise ValueError(
            "Periodical file does not contain Base Rate Capacity or Fuel Surcharge Adder costs."
        )

    entries: list[PeriodicalLaneCost] = []
    for _, row in rate_card.iterrows():
        key = cell_text(row.get("KEY"))
        if not key:
            continue

        valid_from = _parse_rate_card_date(row.get("RATE_EFFECTIVE_DATE__C"))
        valid_to = _parse_rate_card_date(row.get("RATE_EXPIRATION_DATE__C"))
        if valid_from is None or valid_to is None:
            continue

        for kind in PERIODICAL_COST_KINDS:
            block = blocks.get(kind.key)
            if block is None:
                continue
            values = _extract_block_values(row, block)
            if _values_are_empty(values):
                continue
            entries.append(
                PeriodicalLaneCost(
                    key=key,
                    valid_from=valid_from,
                    valid_to=valid_to,
                    kind=kind,
                    values=values,
                )
            )
    return entries


def _values_are_empty(values: dict[str, object]) -> bool:
    for value in values.values():
        if value in (None, ""):
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        return False
    return True


def apply_periodical_costs(
    matrix: MatrixWorkbookData,
    rate_card: pd.DataFrame,
    *,
    matching_keys_only: bool = False,
) -> tuple[MatrixWorkbookData, set[tuple[int, int]], set[int], set[str]]:
    entries = parse_periodical_lane_costs(rate_card)
    if not entries:
        print("  No periodical cost rows with KEY and validity dates found.")
        return matrix, set(), set(), set()

    matrix_keys = set(group_lanes_by_key(matrix.lanes, matrix.columns))
    if matching_keys_only:
        entries = [entry for entry in entries if entry.key in matrix_keys]
        if not entries:
            print("  No periodical rows match keys in the RA.")
            return matrix, set(), set(), set()

    changed_cells: set[tuple[int, int]] = set()
    updated_rows: set[int] = set()
    updated_keys: set[str] = set()
    lanes_by_key = group_lanes_by_key(matrix.lanes, matrix.columns)

    unique_blocks: dict[str, PeriodicalLaneCost] = {}
    for entry in entries:
        cost_name = _build_display_cost_name(entry.kind, entry.valid_from, entry.valid_to)
        unique_blocks[cost_name] = entry

    for cost_name, sample_entry in sorted(unique_blocks.items()):
        if find_cost_block_columns(matrix.columns, cost_name):
            continue
        insert_cost_block_after(
            matrix,
            sample_entry.kind.anchor_cost_name,
            cost_name,
            list(PERIODICAL_SUB_HEADERS),
            _build_header_cells(
                matrix,
                sample_entry.kind,
                cost_name,
                sample_entry.valid_from,
                sample_entry.valid_to,
            ),
            display_base_name=sample_entry.kind.display_base_name,
            valid_from=sample_entry.valid_from,
        )

    for entry in entries:
        cost_name = _build_display_cost_name(entry.kind, entry.valid_from, entry.valid_to)
        columns = find_cost_block_columns(matrix.columns, cost_name)
        if not columns:
            continue

        for lane in lanes_by_key.get(entry.key, []):
            lane_start = parse_date_dd_mm_yyyy(lane_value(lane, matrix.columns, "Valid from"))
            lane_end = parse_date_dd_mm_yyyy(lane_value(lane, matrix.columns, "Valid to"))
            if lane_start is None or lane_end is None:
                continue
            if not _periods_overlap(lane_start, lane_end, entry.valid_from, entry.valid_to):
                continue

            lane_offset = matrix.lanes.index(lane)
            row_index = matrix.data_start_row + lane_offset
            lane.is_updated_row = True
            updated_rows.add(row_index)
            updated_keys.add(entry.key)

            for column in columns:
                value = entry.values.get(column.header)
                if value in (None, "") or (isinstance(value, float) and pd.isna(value)):
                    continue
                lane.values[column.index] = value
                changed_cells.add((row_index, column.index + 1))

    return matrix, changed_cells, updated_rows, updated_keys


def run_periodical_step(
    matrix: MatrixWorkbookData,
    *,
    carrier: str,
    auto: bool,
    periodical_file: Path | None,
    changed_cells: set[tuple[int, int]],
    updated_rows: set[int],
) -> tuple[MatrixWorkbookData, set[tuple[int, int]], set[int], Path | None]:
    periodical_files = list_excel_files(PERIODICAL_DIR)
    selected_periodical = periodical_file or prompt_periodical_file_selection(
        periodical_files,
        directory=PERIODICAL_DIR,
        auto=auto,
    )
    if selected_periodical is None:
        return matrix, changed_cells, updated_rows, None

    processing_path, sheet_name = convert_periodical_file(selected_periodical, carrier)
    rate_card = pd.read_excel(processing_path, sheet_name=sheet_name)
    print(f"  Periodical rate card: {len(rate_card)} rows")

    periodical_matrix, periodical_changed, periodical_rows, _ = apply_periodical_costs(
        matrix,
        rate_card,
    )
    changed_cells |= periodical_changed
    updated_rows |= periodical_rows
    print(
        f"  Periodical costs applied: {len(periodical_changed)} cells, "
        f"{len(periodical_rows)} rows highlighted"
    )
    return periodical_matrix, changed_cells, updated_rows, selected_periodical


def run_periodical_ra(
    *,
    auto: bool = False,
    previous_ra_file: Path | None = None,
    periodical_file: Path | None = None,
    output_path: Path | None = None,
) -> PeriodicalRaResult:
    ensure_workspace_dirs()

    previous_files = list_excel_files(PREVIOUS_RA_DIR)
    periodical_files = list_excel_files(PERIODICAL_DIR)

    selected_previous = previous_ra_file or (
        previous_files[0]
        if auto and previous_files
        else prompt_file_selection(
            previous_files,
            "Select previous RA file:",
            directory=PREVIOUS_RA_DIR,
            always_prompt=True,
        )
    )
    selected_periodical = periodical_file or prompt_periodical_file_selection(
        periodical_files,
        directory=PERIODICAL_DIR,
        auto=auto,
        required=True,
    )
    if selected_periodical is None:
        raise ValueError("A periodical costs file is required for periodical-only update.")

    carrier = detect_carrier_from_filename(selected_previous.name)
    print(f"\nDetected carrier from previous RA: {carrier}")

    print(f"\nLoading previous RA '{RATE_CARD_SHEET_NAME}' tab from {selected_previous.name}...")
    previous_matrix = read_matrix_workbook(selected_previous, rate_card_only=True)
    print(f"  {len(previous_matrix.lanes)} lanes")

    processing_path, sheet_name = convert_periodical_file(selected_periodical, carrier)
    rate_card = pd.read_excel(processing_path, sheet_name=sheet_name)
    print(f"  Periodical rate card: {len(rate_card)} rows")

    updated_matrix, changed_cells, updated_rows, updated_keys = apply_periodical_costs(
        previous_matrix,
        rate_card,
        matching_keys_only=True,
    )

    final_output = output_path or (
        OUTPUT_DIR / f"{selected_previous.stem}_periodical.xlsx"
    )
    print(f"\nWriting updated '{RATE_CARD_SHEET_NAME}' tab to {final_output.name}...")
    write_matrix_workbook(
        updated_matrix,
        final_output,
        changed_cells=changed_cells,
        updated_rows=updated_rows,
        source_workbook_path=selected_previous,
    )

    print("\n=== Periodical update complete ===")
    print(f"  Previous RA:     {selected_previous}")
    print(f"  Periodical file: {selected_periodical}")
    print(f"  Output workbook: {final_output}")
    print(f"  Matching keys:   {len(updated_keys)}")
    print(f"  Updated rows:    {len(updated_rows)}")
    print(f"  Updated cells:   {len(changed_cells)}")

    return PeriodicalRaResult(
        previous_ra_file=selected_previous,
        periodical_file=selected_periodical,
        output_file=final_output,
        matching_key_count=len(updated_keys),
        updated_row_count=len(updated_rows),
        updated_cell_count=len(changed_cells),
    )
