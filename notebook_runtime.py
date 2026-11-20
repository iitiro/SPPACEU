"""Shared notebook configuration. Importing this module never makes API calls."""
from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("SPPACEU_DATA_DIR", str(BASE_DIR / "outputs")))
RECORDS_DIR = Path(os.environ.get("SPPACEU_RECORDS_DIR", str(DATA_DIR / "records")))
MODEL_NAME = os.environ.get("OPENAI_MODEL", "").strip()


def load_openai_key(path=None):
    value = os.environ.get("OPENAI_API_KEY", "").strip()
    if not value and path is not None and Path(path).is_file():
        value = Path(path).read_text(encoding="utf-8").strip()
    if not value:
        raise ValueError("Set OPENAI_API_KEY before enabling RUN_ANALYSIS.")
    return value


def prepare_run(model_name, *outputs):
    if not model_name:
        raise ValueError("Set OPENAI_MODEL to a model available in your API account.")
    for output in outputs:
        Path(output).parent.mkdir(parents=True, exist_ok=True)
