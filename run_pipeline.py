"""
Caterpillar Air — end-to-end rate pipeline.

Steps:
  1. Select input file and tabs, convert to processing/
  2. Build matrix workbook in output/

Usage (local):
  python run_pipeline.py
  python run_pipeline.py --auto
  python run_pipeline.py --convert-only
  python run_pipeline.py --matrix-only

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
    "build_matrix",
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


@dataclass(frozen=True)
class PipelineResult:
    conversion: ConvertResult | None
    matrix: MatrixBuildResult | None

    @property
    def processing_path(self) -> Path | None:
        if self.conversion is None:
            return None
        return self.conversion.output_file

    @property
    def output_path(self) -> Path | None:
        if self.matrix is None:
            return None
        return self.matrix.matrix_path


def run_pipeline(
    *,
    auto: bool = False,
    convert_only: bool = False,
    matrix_only: bool = False,
    input_file: Path | None = None,
    processing_file: Path | None = None,
    output_path: Path | None = None,
) -> PipelineResult:
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
        matrix = run_build_matrix(
            auto=auto or matrix_only,
            processing_file=selected_processing_file,
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

    return PipelineResult(conversion=conversion, matrix=matrix)


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
        run_pipeline(
            auto=args.auto,
            convert_only=args.convert_only,
            matrix_only=args.matrix_only,
            input_file=args.input,
            processing_file=args.processing,
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
        )
    else:
        raise SystemExit(main())
