"""
Load a user-selected input workbook, convert sheet tabs to dataframes,
and save the result to the processing/ folder as .xlsx.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from number_utils import normalize_dataframe_numbers
from project_paths import INPUT_DIR, PROCESSING_DIR, ensure_workspace_dirs

EXCEL_SUFFIXES = {".xlsx", ".xls", ".xlsm"}


@dataclass(frozen=True)
class ConvertResult:
    source_file: Path
    output_file: Path
    sheets: dict[str, pd.DataFrame]


def list_input_files() -> list[Path]:
    return [
        path
        for path in sorted(INPUT_DIR.iterdir())
        if path.is_file()
        and path.suffix.lower() in EXCEL_SUFFIXES
        and not path.name.startswith("~$")
    ]


def list_workbook_sheets(file_path: Path) -> list[str]:
    workbook = pd.ExcelFile(file_path)
    return list(workbook.sheet_names)


def _cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _drop_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    non_empty_mask = df.apply(
        lambda row: any(_cell_text(value) for value in row),
        axis=1,
    )
    return df.loc[non_empty_mask].copy()


def _drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    keep_columns = [
        column
        for column in df.columns
        if any(_cell_text(value) for value in df[column])
    ]
    return df.loc[:, keep_columns].copy()


def sheet_to_df(file_path: Path, sheet_name: str) -> pd.DataFrame:
    """Read one sheet with the first row as column headers."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    df = _drop_empty_rows(df)
    df = _drop_empty_columns(df)
    df = normalize_dataframe_numbers(df)
    return df.reset_index(drop=True)


def sanitize_output_stem(text: str) -> str:
    return re.sub(r"[^\w\-]+", "_", text).strip("_") or "output"


def save_dataframes(
    sheets: dict[str, pd.DataFrame],
    output_path: Path,
) -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output_path


def prompt_file_selection(files: list[Path]) -> Path:
    if not files:
        raise FileNotFoundError(f"No Excel files found in: {INPUT_DIR}")

    if len(files) == 1:
        print(f"\nUsing only file in input/: {files[0].name}")
        return files[0]

    print("\nSelect input file:")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name}")

    while True:
        raw = input("Enter file number: ").strip()
        try:
            choice = int(raw) - 1
            if 0 <= choice < len(files):
                return files[choice]
        except ValueError:
            pass
        print(f"Please enter a number between 1 and {len(files)}.")


def prompt_sheet_selection(file_path: Path, sheet_names: list[str]) -> list[str]:
    if len(sheet_names) == 1:
        print(f"\nUsing only tab in {file_path.name}: {sheet_names[0]}")
        return sheet_names

    print(f"\nAvailable tabs in {file_path.name}:")
    for index, name in enumerate(sheet_names, start=1):
        print(f"  {index}. {name}")
    print("  all. Convert all tabs")

    while True:
        raw = input("Enter tab number(s), comma-separated, or 'all': ").strip().lower()
        if raw in {"all", "*"}:
            return sheet_names

        selected: list[str] = []
        try:
            for part in re.split(r"\s*,\s*", raw):
                if not part:
                    continue
                choice = int(part) - 1
                if 0 <= choice < len(sheet_names):
                    selected.append(sheet_names[choice])
                else:
                    raise ValueError
            if selected:
                return selected
        except ValueError:
            pass
        print(f"Please enter valid tab number(s) between 1 and {len(sheet_names)}, or 'all'.")


def select_input_file(files: list[Path], *, auto: bool) -> Path:
    if auto:
        selected = files[0]
        if len(files) == 1:
            print(f"\nAuto mode: using file {selected.name}")
        else:
            print(f"\nAuto mode: using first file {selected.name}")
        return selected
    return prompt_file_selection(files)


def select_sheets(file_path: Path, sheet_names: list[str], *, auto: bool) -> list[str]:
    if auto:
        print(f"Auto mode: converting all tabs ({', '.join(sheet_names)})")
        return sheet_names
    return prompt_sheet_selection(file_path, sheet_names)


def convert_workbook_to_dataframes(
    file_path: Path,
    sheet_names: list[str],
) -> dict[str, pd.DataFrame]:
    sheets: dict[str, pd.DataFrame] = {}
    for sheet_name in sheet_names:
        print(f"  Loading '{sheet_name}'...")
        df = sheet_to_df(file_path, sheet_name)
        print(f"    {len(df)} rows, {len(df.columns)} columns")
        sheets[sheet_name] = df
    return sheets


def run_convert(
    *,
    auto: bool = False,
    input_file: Path | None = None,
    sheet_names: list[str] | None = None,
    output_path: Path | None = None,
) -> ConvertResult:
    ensure_workspace_dirs()

    input_files = list_input_files()
    selected_file = input_file or select_input_file(input_files, auto=auto)

    if not selected_file.exists():
        raise FileNotFoundError(f"Input file not found: {selected_file}")

    available_sheets = list_workbook_sheets(selected_file)
    selected_sheets = sheet_names or select_sheets(selected_file, available_sheets, auto=auto)

    for sheet_name in selected_sheets:
        if sheet_name not in available_sheets:
            raise ValueError(f"Sheet not found in {selected_file.name}: {sheet_name}")

    print(f"\nConverting {selected_file.name}...")
    dataframes = convert_workbook_to_dataframes(selected_file, selected_sheets)

    final_output = output_path or (
        PROCESSING_DIR / f"{sanitize_output_stem(selected_file.stem)}.xlsx"
    )
    save_dataframes(dataframes, final_output)
    print(f"\nSaved processing workbook to: {final_output}")

    return ConvertResult(
        source_file=selected_file,
        output_file=final_output,
        sheets=dataframes,
    )


def main() -> int:
    try:
        run_convert()
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
