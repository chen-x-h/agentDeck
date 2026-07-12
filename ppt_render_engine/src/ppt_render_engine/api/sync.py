import os
import json
import yaml
from pathlib import Path
from fastapi import APIRouter, HTTPException
from ppt_render_engine.log_config import app_logger

logger = app_logger("api.sync")

router = APIRouter(prefix="/sync", tags=["Sync"])

_cfg_path = Path(__file__).parents[3] / "config.yaml"


def get_sync_path() -> str:
    try:
        with open(_cfg_path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        path = cfg.get("sync", {}).get("json_path", "./sync/presentation.json")
    except Exception:
        path = "./sync/presentation.json"
    return os.path.abspath(path)


def _read_sync():
    path = get_sync_path()
    if not os.path.isfile(path):
        return None, None
    mtime = os.path.getmtime(path)
    with open(path, encoding="utf-8") as f:
        content = json.load(f)
    return mtime, content


@router.get("")
async def sync_load():
    path = get_sync_path()
    if not os.path.isfile(path):
        return {"mtime": None, "content": None}
    mtime = os.path.getmtime(path)
    try:
        with open(path, encoding="utf-8") as f:
            content = json.load(f)
        return {"mtime": mtime, "content": content}
    except (json.JSONDecodeError, Exception) as e:
        raise HTTPException(400, detail=f"同步文件解析失败: {e}")


@router.put("")
async def sync_save(body: dict):
    path = get_sync_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False, indent=2)
    mtime = os.path.getmtime(path)
    logger.info("Sync file saved", path=path)
    return {"mtime": mtime, "status": "saved"}


@router.post("/notify")
async def sync_notify():
    """Agent calls this after writing the sync file."""
    path = get_sync_path()
    if not os.path.isfile(path):
        raise HTTPException(404, detail="同步文件不存在")
    mtime, content = _read_sync()
    if content is None:
        raise HTTPException(404, detail="同步文件解析失败")
    logger.info("Sync notified by agent", mtime=mtime)
    return {"mtime": mtime, "content": content}
