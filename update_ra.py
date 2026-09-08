"""
Apply Caterpillar Air rate updates to a previous RA matrix workbook.

Inputs:
  - Previous RA matrix from input/previous RA/
  - Update rate card from input/update/ (carrier tab selected from previous RA name)

Output:
  - Updated matrix workbook in output/ with _update suffix and green highlights
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from build_matrix import (
    build_matrix_data_from_rate_card,
    load_rate_card_from_processing,
    resolve_carrier_workbook_sheet,
)
from convert_to_processing import convert_workbook_to_dataframes, save_dataframes
from matrix_io import (
    MatrixColumn,
    MatrixLane,
    MatrixWorkbookData,
    RATE_CARD_SHEET_NAME,
    build_cost_column_map,
    cell_text,
    find_cost_column_index_fuzzy,
    find_shipment_column_index,
    format_date_dd_mm_yyyy,
    group_lanes_by_key,
    lane_value,
    list_excel_files,
    parse_date_dd_mm_yyyy,
    prompt_file_selection,
    read_matrix_workbook,
    set_lane_value,
    write_matrix_workbook,
)
from periodical_costs import run_periodical_step
from project_paths import (
    OUTPUT_DIR,
    PREVIOUS_RA_DIR,
    PROCESSING_DIR,
    UPDATE_DIR,
    ensure_workspace_dirs,
)

CARRIER_PATTERN = re.compile(r"\b(?:Air\s+)?(DHL|KN)\b", re.IGNORECASE)


@dataclass(frozen=True)
class UpdateRaResult:
    previous_ra_file: Path
    update_file: Path
    output_file: Path
    updated_lane_count: int
    split_lane_count: int
    new_lane_count: int
    periodical_file: Path | None = None


def detect_carrier_from_filename(filename: str) -> str:
    match = CARRIER_PATTERN.search(filename)
    if not match:
        raise ValueError(
            f"Could not detect carrier (DHL/KN) from filename: {filename}"
        )
    return match.group(1).upper()


def _values_equal(left: object, right: object) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return str(left).strip() == str(right).strip()


def _find_lane_covering(target_date: date, lanes: list[MatrixLane], columns: list[MatrixColumn]) -> MatrixLane | None:
    matches: list[MatrixLane] = []
    for lane in lanes:
        valid_from = parse_date_dd_mm_yyyy(lane_value(lane, columns, "Valid from"))
        valid_to = parse_date_dd_mm_yyyy(lane_value(lane, columns, "Valid to"))
        if valid_from is None or valid_to is None:
            continue
        if valid_from <= target_date <= valid_to:
            matches.append(lane)
    if not matches:
        return None
    return max(
        matches,
        key=lambda lane: parse_date_dd_mm_yyyy(lane_value(lane, columns, "Valid from")) or date.min,
    )


def _copy_lane(lane: MatrixLane) -> MatrixLane:
    return MatrixLane(values=list(lane.values))


def merge_lane_row(
    previous_lane: MatrixLane,
    update_lane: MatrixLane,
    target_columns: list[MatrixColumn],
    update_columns: list[MatrixColumn],
    *,
    cost_column_map: dict[int, int] | None = None,
) -> tuple[MatrixLane, set[int]]:
    merged = _copy_lane(previous_lane)
    changed: set[int] = set()
    cost_map = cost_column_map or build_cost_column_map(target_columns, update_columns)

    for update_column in update_columns:
        update_value = update_lane.values[update_column.index]
        if _is_empty(update_value):
            continue

        if update_column.cost_name is None:
            if update_column.header in {"Lane #", "Valid from", "Valid to", "OPERATIONAL_FLOW"}:
                continue
            target_index = find_shipment_column_index(target_columns, update_column.header)
        else:
            target_index = cost_map.get(update_column.index)

        if target_index is None:
            continue

        if not _values_equal(merged.values[target_index], update_value):
            changed.add(target_index)
        merged.values[target_index] = update_value

    return merged, changed


def _is_empty(value: object) -> bool:
    if value is None:
        return True
    return str(value).strip() == ""


def split_lane_by_update(
    previous_lane: MatrixLane,
    update_lanes: list[MatrixLane],
    target_columns: list[MatrixColumn],
    update_columns: list[MatrixColumn],
    *,
    cost_column_map: dict[int, int] | None = None,
) -> list[MatrixLane]:
    prev_start = parse_date_dd_mm_yyyy(lane_value(previous_lane, target_columns, "Valid from"))
    prev_end = parse_date_dd_mm_yyyy(lane_value(previous_lane, target_columns, "Valid to"))
    if prev_start is None or prev_end is None:
        raise ValueError("Previous lane is missing Valid from / Valid to dates.")

    boundaries: set[date] = {prev_start, prev_end + timedelta(days=1)}
    for update_lane in update_lanes:
        update_start = parse_date_dd_mm_yyyy(lane_value(update_lane, update_columns, "Valid from"))
        update_end = parse_date_dd_mm_yyyy(lane_value(update_lane, update_columns, "Valid to"))
        if update_start is not None:
            boundaries.add(update_start)
        if update_end is not None:
            boundaries.add(update_end + timedelta(days=1))

    sorted_bounds = sorted(boundaries)
    results: list[MatrixLane] = []

    for index in range(len(sorted_bounds) - 1):
        segment_start = sorted_bounds[index]
        segment_end = sorted_bounds[index + 1] - timedelta(days=1)
        if segment_start > segment_end:
            continue

        covering_update = _find_lane_covering(segment_start, update_lanes, update_columns)
        segment_within_previous = segment_start <= prev_end and segment_end >= prev_start

        if covering_update is not None:
            merged_lane, changed_columns = merge_lane_row(
                previous_lane,
                covering_update,
                target_columns,
                update_columns,
                cost_column_map=cost_column_map,
            )
            set_lane_value(merged_lane, target_columns, "Valid from", format_date_dd_mm_yyyy(segment_start))
            set_lane_value(merged_lane, target_columns, "Valid to", format_date_dd_mm_yyyy(segment_end))
            valid_from_index = find_shipment_column_index(target_columns, "Valid from")
            valid_to_index = find_shipment_column_index(target_columns, "Valid to")
            if valid_from_index is not None:
                changed_columns.add(valid_from_index)
            if valid_to_index is not None:
                changed_columns.add(valid_to_index)
            merged_lane.changed_columns = changed_columns
            results.append(merged_lane)
        elif segment_within_previous:
            preserved = _copy_lane(previous_lane)
            set_lane_value(preserved, target_columns, "Valid from", format_date_dd_mm_yyyy(segment_start))
            set_lane_value(preserved, target_columns, "Valid to", format_date_dd_mm_yyyy(segment_end))
            preserved.changed_columns = set()
            results.append(preserved)

    return results


def _group_lanes_by_key(
    lanes: list[MatrixLane],
    columns: list[MatrixColumn],
) -> dict[str, list[MatrixLane]]:
    return group_lanes_by_key(lanes, columns)


def _cell_text(value: object) -> str:
    return cell_text(value)


def _extend_lane_values(lane: MatrixLane, target_length: int) -> None:
    while len(lane.values) < target_length:
        lane.values.append(None)


def apply_updates(
    previous: MatrixWorkbookData,
    update_matrix: MatrixWorkbookData,
) -> tuple[MatrixWorkbookData, set[tuple[int, int]], set[int]]:
    result = MatrixWorkbookData(
        sheet_name=previous.sheet_name,
        columns=[
            MatrixColumn(index=column.index, header=column.header, cost_name=column.cost_name)
            for column in previous.columns
        ],
        header_rows=[list(row) for row in previous.header_rows],
        lanes=[],
        extra_sheets=dict(previous.extra_sheets),
        header_row=previous.header_row,
        data_start_row=previous.data_start_row,
    )
    cost_column_map = build_cost_column_map(result.columns, update_matrix.columns)

    for lane in previous.lanes:
        _extend_lane_values(lane, len(result.columns))

    previous_by_key = _group_lanes_by_key(previous.lanes, previous.columns)
    update_by_key = _group_lanes_by_key(update_matrix.lanes, update_matrix.columns)
    changed_cells: set[tuple[int, int]] = set()
    updated_rows: set[int] = set()
    processed_keys: set[str] = set()

    for lane in previous.lanes:
        key = _cell_text(lane_value(lane, previous.columns, "KEY"))
        if not key or key not in update_by_key:
            copied = _copy_lane(lane)
            _extend_lane_values(copied, len(result.columns))
            copied.changed_columns = set()
            result.lanes.append(copied)
            continue

        if key in processed_keys:
            continue

        for previous_lane in previous_by_key[key]:
            split_lanes = split_lane_by_update(
                previous_lane,
                update_by_key[key],
                result.columns,
                update_matrix.columns,
                cost_column_map=cost_column_map,
            )
            for split_lane in split_lanes:
                _extend_lane_values(split_lane, len(result.columns))
                split_lane.is_updated_row = True
                result.lanes.append(split_lane)
        processed_keys.add(key)

    previous_keys = set(previous_by_key)
    for key, update_lanes in update_by_key.items():
        if key in previous_keys:
            continue
        for update_lane in update_lanes:
            merged_lane, changed_columns = merge_lane_row(
                MatrixLane(values=[None] * len(result.columns)),
                update_lane,
                result.columns,
                update_matrix.columns,
                cost_column_map=cost_column_map,
            )
            key_index = find_shipment_column_index(result.columns, "KEY")
            if key_index is not None:
                merged_lane.values[key_index] = key
                changed_columns.add(key_index)
            for header in ("Valid from", "Valid to"):
                header_index = find_shipment_column_index(result.columns, header)
                update_index = find_shipment_column_index(update_matrix.columns, header)
                if header_index is not None and update_index is not None:
                    merged_lane.values[header_index] = update_lane.values[update_index]
                    changed_columns.add(header_index)
            merged_lane.changed_columns = changed_columns
            merged_lane.is_updated_row = True
            result.lanes.append(merged_lane)

    for lane_offset, lane in enumerate(result.lanes):
        row_index = result.data_start_row + lane_offset
        if lane.is_updated_row:
            updated_rows.add(row_index)
        for column_index in lane.changed_columns:
            changed_cells.add((row_index, column_index + 1))

    return result, changed_cells, updated_rows


def convert_update_file(
    update_file: Path,
    carrier: str,
) -> tuple[Path, str]:
    workbook = pd.ExcelFile(update_file)
    sheet_name = resolve_carrier_workbook_sheet(workbook.sheet_names, carrier)

    print(f"\nConverting update file '{update_file.name}' (tab: {sheet_name})...")
    dataframes = convert_workbook_to_dataframes(update_file, [sheet_name])
    output_path = PROCESSING_DIR / f"{update_file.stem}_update.xlsx"
    save_dataframes(dataframes, output_path)
    print(f"Saved update processing workbook to: {output_path}")
    return output_path, sheet_name


def run_update_ra(
    *,
    auto: bool = False,
    previous_ra_file: Path | None = None,
    update_file: Path | None = None,
    periodical_file: Path | None = None,
    output_path: Path | None = None,
) -> UpdateRaResult:
    ensure_workspace_dirs()

    previous_files = list_excel_files(PREVIOUS_RA_DIR)
    update_files = list_excel_files(UPDATE_DIR)

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
    selected_update = update_file or (
        update_files[0]
        if auto and update_files
        else prompt_file_selection(
            update_files,
            "Select update file:",
            directory=UPDATE_DIR,
            always_prompt=True,
        )
    )

    carrier = detect_carrier_from_filename(selected_previous.name)
    print(f"\nDetected carrier from previous RA: {carrier}")

    print(f"\nLoading previous RA '{RATE_CARD_SHEET_NAME}' tab from {selected_previous.name}...")
    previous_matrix = read_matrix_workbook(selected_previous, rate_card_only=True)
    print(f"  {len(previous_matrix.lanes)} lanes")

    processing_path, sheet_name = convert_update_file(selected_update, carrier)
    rate_card = load_rate_card_from_processing(processing_path, sheet_name=sheet_name)
    print(f"  Update rate card: {len(rate_card)} rows")

    update_matrix = build_matrix_data_from_rate_card(
        rate_card,
        sheet_name=RATE_CARD_SHEET_NAME,
    )
    print(f"  Update matrix lanes: {len(update_matrix.lanes)}")

    previous_keys = set(_group_lanes_by_key(previous_matrix.lanes, previous_matrix.columns))
    update_keys = set(_group_lanes_by_key(update_matrix.lanes, update_matrix.columns))
    matching_keys = previous_keys & update_keys

    updated_matrix, changed_cells, updated_rows = apply_updates(previous_matrix, update_matrix)

    updated_matrix, changed_cells, updated_rows, selected_periodical = run_periodical_step(
        updated_matrix,
        carrier=carrier,
        auto=auto,
        periodical_file=periodical_file,
        changed_cells=changed_cells,
        updated_rows=updated_rows,
    )

    final_output = output_path or (
        OUTPUT_DIR / f"{selected_previous.stem}_update.xlsx"
    )
    print(f"\nWriting updated '{RATE_CARD_SHEET_NAME}' tab to {final_output.name}...")
    write_matrix_workbook(
        updated_matrix,
        final_output,
        changed_cells=changed_cells,
        updated_rows=updated_rows,
        source_workbook_path=selected_previous,
    )

    split_lane_count = sum(
        1
        for lane in updated_matrix.lanes
        if lane.changed_columns
    )
    new_lane_count = sum(
        1
        for lane in updated_matrix.lanes
        if _cell_text(lane_value(lane, updated_matrix.columns, "KEY")) in update_keys - previous_keys
    )

    print("\n=== RA update complete ===")
    print(f"  Previous RA:     {selected_previous}")
    print(f"  Update source:   {selected_update}")
    if selected_periodical is not None:
        print(f"  Periodical file: {selected_periodical}")
    print(f"  Output workbook: {final_output}")
    print(f"  Matching keys:   {len(matching_keys)}")
    print(f"  Output lanes:    {len(updated_matrix.lanes)}")
    print(f"  Updated rows:    {len(updated_rows)}")
    print(f"  Updated cells:   {len(changed_cells)}")

    return UpdateRaResult(
        previous_ra_file=selected_previous,
        update_file=selected_update,
        output_file=final_output,
        updated_lane_count=len(matching_keys),
        split_lane_count=split_lane_count,
        new_lane_count=new_lane_count,
        periodical_file=selected_periodical,
    )


def main() -> int:
    try:
        run_update_ra()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
