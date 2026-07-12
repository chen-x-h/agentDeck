import os
import yaml
import uuid
from pathlib import Path
from ppt_render_engine.log_config import app_logger

logger = app_logger("temp_util")
_cfg_path = Path(__file__).resolve().parents[3] / "config.yaml"


def get_temp_dir() -> str:
    try:
        with open(_cfg_path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        rel = cfg.get("engine", {}).get("temp_dir", "./temp")
    except Exception:
        rel = "./temp"
    project_root = Path(__file__).resolve().parents[3]
    d = os.path.abspath(os.path.join(project_root, rel))
    os.makedirs(d, exist_ok=True)
    return d


def get_temp_path(suffix: str = ".tmp") -> str:
    return os.path.join(get_temp_dir(), f"{uuid.uuid4().hex}{suffix}")
