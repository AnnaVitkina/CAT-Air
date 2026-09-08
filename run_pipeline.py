"""
Caterpillar Air — end-to-end rate pipeline.

Steps:
  1. Select input file and tabs, convert to processing/
  2. Build matrix workbook in output/

Update mode (--update):
  1. Select previous RA from input/previous RA/
  2. Select update file from input/update/
  3. Select periodical costs file from input/periodical/ (optional)
  4. Match lanes by KEY, split validity dates, merge costs/locations
  5. Add periodical cost columns (Base Rate Capacity, Fuel Surcharge Adder)
  6. Save updated matrix to output/ with _update suffix

Periodical-only mode (--periodical-only):
  1. Select previous RA from input/previous RA/
  2. Select periodical costs file from input/periodical/
  3. Apply periodical costs only to lanes with matching KEY
  4. Save updated matrix to output/ with _periodical suffix

Usage (local):
  python run_pipeline.py                  # asks: create, update, or periodical-only
  python run_pipeline.py --auto
  python run_pipeline.py --convert-only
  python run_pipeline.py --matrix-only
  python run_pipeline.py --update         # skip prompt, go straight to update
  python run_pipeline.py --update --auto
  python run_pipeline.py --periodical-only
  python run_pipeline.py --periodical-only --auto

Usage (Google Colab):
    from google.colab import drive
    drive.mount("/content/drive")
    exec(open("/content/CAT-Air/run_pipeline.py").read())

  Data is read from Google Drive automatically when this folder exists:
    .../RMT Caterpillar/Air/input
    .../RMT Caterpillar/Air/processing
    .../RMT Caterpillar/Air/output

  # Optional overrides:
  # os.environ["CAT_AIR_DRIVE_BASE"] = "/content/drive/.../RMT Caterpillar/Air"
  # os.environ["CAT_AIR_AUTO"] = "1"
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path

_COLAB_CODE_DIRS = (
    Path("/content/CAT-Air"),
    Path("/content/CAT-air"),
)


def _resolve_code_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    for path in _COLAB_CODE_DIRS:
        if path.is_dir():
            return path
    return Path.cwd()


_CODE_DIR = _resolve_code_dir()
if str(_CODE_DIR) not in sys.path:
    sys.path.insert(0, str(_CODE_DIR))

_PIPELINE_MODULES = (
    "project_paths",
    "number_utils",
    "convert_to_processing",
    "build_accessorial_costs",
    "export_origin_postal_code_zones",
    "build_matrix",
    "matrix_io",
    "update_ra",
    "periodical_costs",
)


def _bootstrap_paths() -> object:
    for module_name in _PIPELINE_MODULES:
        sys.modules.pop(module_name, None)

    import project_paths

    project_paths.configure_paths_from_env()
    return project_paths


_project_paths = _bootstrap_paths()
configure_paths_from_env = _project_paths.configure_paths_from_env
print_path_config = _project_paths.print_path_config

from build_matrix import MatrixBuildResult, run_build_matrix  # noqa: E402
from convert_to_processing import ConvertResult, run_convert  # noqa: E402
from periodical_costs import PeriodicalRaResult, run_periodical_ra  # noqa: E402
from update_ra import UpdateRaResult, run_update_ra  # noqa: E402


PipelineMode = str  # "create" | "update" | "periodical"


@dataclass(frozen=True)
class PipelineResult:
    conversion: ConvertResult | None
    matrix: MatrixBuildResult | None
    update: UpdateRaResult | None = None
    periodical: PeriodicalRaResult | None = None

    @property
    def processing_path(self) -> Path | None:
        if self.conversion is None:
            return None
        return self.conversion.output_file

    @property
    def output_path(self) -> Path | None:
        if self.periodical is not None:
            return self.periodical.output_file
        if self.update is not None:
            return self.update.output_file
        if self.matrix is None:
            return None
        return self.matrix.matrix_path


def run_periodical_pipeline(
    *,
    auto: bool = False,
    previous_ra_file: Path | None = None,
    periodical_file: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    configure_paths_from_env()
    print_path_config()

    print("\n=== RA periodical update ===")
    periodical = run_periodical_ra(
        auto=auto,
        previous_ra_file=previous_ra_file,
        periodical_file=periodical_file,
        output_path=output_path,
    )
    return PipelineResult(conversion=None, matrix=None, periodical=periodical)


def run_update_pipeline(
    *,
    auto: bool = False,
    previous_ra_file: Path | None = None,
    update_file: Path | None = None,
    periodical_file: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    configure_paths_from_env()
    print_path_config()

    print("\n=== RA update ===")
    update = run_update_ra(
        auto=auto,
        previous_ra_file=previous_ra_file,
        update_file=update_file,
        periodical_file=periodical_file,
        output_path=output_path,
    )
    return PipelineResult(conversion=None, matrix=None, update=update)


def run_pipeline(
    *,
    auto: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    mode: PipelineMode = "create",
    input_file: Path | None = None,
    processing_file: Path | None = None,
    previous_ra_file: Path | None = None,
    update_file: Path | None = None,
    periodical_file: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
    if mode == "update":
        if convert_only or matrix_only:
            raise ValueError("Use --update without --convert-only or --matrix-only.")
        return run_update_pipeline(
            auto=auto,
            previous_ra_file=previous_ra_file,
            update_file=update_file,
            periodical_file=periodical_file,
            output_path=output_path,
        )

    if mode == "periodical":
        if convert_only or matrix_only:
            raise ValueError("Use --periodical-only without --convert-only or --matrix-only.")
        return run_periodical_pipeline(
            auto=auto,
            previous_ra_file=previous_ra_file,
            periodical_file=periodical_file,
            output_path=output_path,
        )

    if convert_only and matrix_only:
        raise ValueError("Use only one of --convert-only or --matrix-only.")

    configure_paths_from_env()
    print_path_config()

    conversion: ConvertResult | None = None
    matrix: MatrixBuildResult | None = None
    selected_processing_file = processing_file

    if not matrix_only:
        print("\n=== Step 1/2: Convert input to processing ===")
        conversion = run_convert(auto=auto, input_file=input_file)
        selected_processing_file = conversion.output_file

    if not convert_only:
        step_label = "Step 2/2" if not matrix_only else "Step 1/1"
        print(f"\n=== {step_label}: Build matrix ===")
        rate_card_sheet: str | None = None
        if conversion is not None and len(conversion.sheets) == 1:
            rate_card_sheet = next(iter(conversion.sheets))
        matrix = run_build_matrix(
            auto=auto or matrix_only,
            processing_file=selected_processing_file,
            sheet_name=rate_card_sheet,
            output_path=output_path,
        )

    print("\n=== Pipeline complete ===")
    if conversion is not None:
        print(f"  Source file:         {conversion.source_file}")
        print(f"  Processing workbook: {conversion.output_file}")
    elif selected_processing_file is not None:
        print(f"  Processing workbook: {selected_processing_file}")
    if matrix is not None:
        print(f"  Matrix workbook:     {matrix.matrix_path}")
        print(
            f"    Rows: {matrix.row_count} | "
            f"Shipment columns: {matrix.shipment_column_count} | "
            f"Cost blocks: {matrix.cost_block_count}"
        )
        if matrix.accessorial is not None:
            print(
                f"    Accessorial tab:   {matrix.accessorial.sheet_name} "
                f"({matrix.accessorial.row_count} rows)"
            )
        if matrix.postal_zones is not None:
            print(
                f"    Postal zones txt:  {matrix.postal_zones.output_path} "
                f"({matrix.postal_zones.zone_count} zones)"
            )

    return PipelineResult(conversion=conversion, matrix=matrix)


def prompt_pipeline_mode() -> PipelineMode:
    """Ask the user to choose between creation, update, and periodical-only."""
    print("\nWhat would you like to do?")
    print("  1. Create new matrix from rate card")
    print("  2. Update existing RA with rate update file")
    print("  3. Apply periodical costs to existing RA only")

    while True:
        raw = input("Enter choice (1, 2, or 3): ").strip()
        if raw == "1":
            return "create"
        if raw == "2":
            return "update"
        if raw == "3":
            return "periodical"
        print("Please enter 1, 2, or 3.")


def resolve_pipeline_mode(
    *,
    update: bool = False,
    periodical_only: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    auto: bool = False,
) -> PipelineMode:
    if periodical_only:
        return "periodical"
    if update:
        return "update"
    if convert_only or matrix_only:
        return "create"
    if auto:
        return "create"
    return prompt_pipeline_mode()


def resolve_update_mode(
    *,
    update: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    auto: bool = False,
) -> bool:
    return resolve_pipeline_mode(
        update=update,
        convert_only=convert_only,
        matrix_only=matrix_only,
        auto=auto,
    ) == "update"


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the Caterpillar Air end-to-end rate pipeline.",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Use the first input file, all tabs, and skip interactive prompts.",
    )
    parser.add_argument(
        "--convert-only",
        action="store_true",
        help="Only run input -> processing conversion.",
    )
    parser.add_argument(
        "--matrix-only",
        action="store_true",
        help="Only build matrix from a processing workbook.",
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help="Update a previous RA matrix using a rate update file.",
    )
    parser.add_argument(
        "--periodical-only",
        action="store_true",
        help="Apply periodical costs to a previous RA without a rate update file.",
    )
    parser.add_argument(
        "--previous-ra",
        type=Path,
        default=None,
        help="Optional previous RA workbook for --update.",
    )
    parser.add_argument(
        "--update-file",
        type=Path,
        default=None,
        help="Optional update workbook for --update.",
    )
    parser.add_argument(
        "--periodical-file",
        type=Path,
        default=None,
        help="Optional periodical costs workbook for --update.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Optional input workbook for conversion.",
    )
    parser.add_argument(
        "--processing",
        type=Path,
        default=None,
        help="Optional processing workbook for --matrix-only.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional output matrix workbook path.",
    )
    return parser.parse_args()


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "on"}


def _running_in_notebook() -> bool:
    if "colab_kernel_launcher" in Path(sys.argv[0]).name:
        return True
    if any(arg == "-f" for arg in sys.argv):
        return True
    return "ipykernel" in sys.modules or "IPython" in sys.modules


def main() -> int:
    try:
        args = _parse_args()
        if args.update and args.periodical_only:
            raise ValueError("Use only one of --update or --periodical-only.")
        pipeline_mode = resolve_pipeline_mode(
            update=args.update,
            periodical_only=args.periodical_only,
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
            auto=args.auto,
        )
        run_pipeline(
            auto=args.auto,
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
            mode=pipeline_mode,
            input_file=args.input,
            processing_file=args.processing,
            previous_ra_file=args.previous_ra,
            update_file=args.update_file,
            periodical_file=args.periodical_file,
            output_path=args.output,
        )
        return 0
    except KeyboardInterrupt:
        print("\nOperation cancelled.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    if _running_in_notebook():
        run_pipeline(
            auto=_env_flag("CAT_AIR_AUTO"),
            convert_only=_env_flag("CAT_AIR_CONVERT_ONLY"),
            matrix_only=_env_flag("CAT_AIR_MATRIX_ONLY"),
            mode=resolve_pipeline_mode(
                update=_env_flag("CAT_AIR_UPDATE"),
                periodical_only=_env_flag("CAT_AIR_PERIODICAL_ONLY"),
                convert_only=_env_flag("CAT_AIR_CONVERT_ONLY"),
                matrix_only=_env_flag("CAT_AIR_MATRIX_ONLY"),
                auto=_env_flag("CAT_AIR_AUTO"),
            ),
        )
    else:
        raise SystemExit(main())
