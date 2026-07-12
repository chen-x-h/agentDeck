import logging
from pathlib import Path

from ppt_render_engine.logger import get_logger

_loggers = {}


def resolve_log_path() -> str:
    project_root = Path(__file__).resolve().parent.parent.parent
    log_dir = project_root / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return str(log_dir / "out.log")


def app_logger(name: str = "ppt_engine") -> object:
    if name not in _loggers:
        log_file = resolve_log_path()
        _loggers[name] = get_logger(name=name, level=logging.DEBUG, log_file=log_file)
    return _loggers[name]
