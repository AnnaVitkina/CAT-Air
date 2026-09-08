"""
Project folder paths for Caterpillar Air rate conversion.

Colab layout:
  - Python code:  /content/CAT-Air/
  - Data folders: Google Drive (input / processing / output)

Colab — run the pipeline:
    from google.colab import drive
    drive.mount("/content/drive")
    exec(open("/content/CAT-Air/run_pipeline.py").read())
"""

from __future__ import annotations

import os
from pathlib import Path

# Where .py files live when uploaded to Colab.
_COLAB_CODE_DIRS = (
    Path("/content/CAT-Air"),
    Path("/content/CAT-air"),
)

COLAB_DRIVE_BASE = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_CAT/Air"
)


def _resolve_script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        pass
    for path in _COLAB_CODE_DIRS:
        if path.is_dir():
            return path
    return Path.cwd()


def _is_colab_runtime() -> bool:
    return Path("/content").is_dir() and any(path.is_dir() for path in _COLAB_CODE_DIRS)


_SCRIPT_DIR = _resolve_script_dir()
BASE_DIR = next((path for path in _COLAB_CODE_DIRS if path.is_dir()), _SCRIPT_DIR)
ROOT = BASE_DIR
INPUT_DIR = ROOT / "input"
PREVIOUS_RA_DIR = INPUT_DIR / "previous RA"
UPDATE_DIR = INPUT_DIR / "update"
PERIODICAL_DIR = INPUT_DIR / "periodical"
PROCESSING_DIR = ROOT / "processing"
OUTPUT_DIR = ROOT / "output"


def _path_from_env(name: str) -> Path | None:
    value = os.environ.get(name)
    if not value:
        return None
    return Path(value).expanduser()


def _drive_data_root() -> Path | None:
    candidates: list[Path] = []
    env_base = _path_from_env("CAT_AIR_DRIVE_BASE")
    if env_base is not None:
        candidates.append(env_base)
    candidates.append(COLAB_DRIVE_BASE)

    for candidate in candidates:
        if candidate.is_dir():
            return candidate
    return None


def _apply_default_data_paths() -> None:
    """Point data folders to Google Drive in Colab when the Drive root exists."""
    global INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR, ROOT

    code_root = _path_from_env("CAT_AIR_CODE_DIR")
    ROOT = code_root.expanduser().resolve() if code_root is not None else BASE_DIR

    input_dir = _path_from_env("CAT_AIR_INPUT_DIR")
    processing_dir = _path_from_env("CAT_AIR_PROCESSING_DIR")
    output_dir = _path_from_env("CAT_AIR_OUTPUT_DIR")

    if input_dir is None and processing_dir is None and output_dir is None:
        drive_root = _drive_data_root()
        if drive_root is not None and _is_colab_runtime():
            input_dir = drive_root / "input"
            processing_dir = drive_root / "processing"
            output_dir = drive_root / "output"

    if input_dir is None:
        input_dir = BASE_DIR / "input"
    if processing_dir is None:
        processing_dir = BASE_DIR / "processing"
    if output_dir is None:
        output_dir = BASE_DIR / "output"

    INPUT_DIR = input_dir.expanduser().resolve()
    PROCESSING_DIR = processing_dir.expanduser().resolve()
    OUTPUT_DIR = output_dir.expanduser().resolve()


_apply_default_data_paths()


def configure_paths(
    *,
    root: Path | str | None = None,
    input_dir: Path | str | None = None,
    processing_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
) -> None:
    """Override data folder locations (for Colab / Google Drive)."""
    global ROOT, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR

    if root is not None:
        ROOT = Path(root).expanduser().resolve()
    if input_dir is not None:
        INPUT_DIR = Path(input_dir).expanduser().resolve()
    if processing_dir is not None:
        PROCESSING_DIR = Path(processing_dir).expanduser().resolve()
    if output_dir is not None:
        OUTPUT_DIR = Path(output_dir).expanduser().resolve()


def configure_paths_from_env() -> None:
    """Re-apply CAT_AIR_* environment variables and Colab Drive defaults."""
    _apply_default_data_paths()


def ensure_workspace_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PREVIOUS_RA_DIR.mkdir(parents=True, exist_ok=True)
    UPDATE_DIR.mkdir(parents=True, exist_ok=True)
    PERIODICAL_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def print_path_config() -> None:
    print("Caterpillar Air paths:")
    print(f"  Code:       {BASE_DIR}")
    print(f"  Root:       {ROOT}")
    print(f"  Input:      {INPUT_DIR}")
    print(f"  Previous RA:{PREVIOUS_RA_DIR}")
    print(f"  Update:     {UPDATE_DIR}")
    print(f"  Periodical: {PERIODICAL_DIR}")
    print(f"  Processing: {PROCESSING_DIR}")
    print(f"  Output:     {OUTPUT_DIR}")
    if INPUT_DIR != BASE_DIR / "input":
        print("  (data folders on Google Drive)")
