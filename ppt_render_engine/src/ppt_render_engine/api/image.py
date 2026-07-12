import os
import yaml
from pathlib import Path
from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from PIL import Image
from ppt_render_engine.log_config import app_logger
from ppt_render_engine.temp_util import get_temp_path

router = APIRouter(prefix="/image", tags=["Image"])
logger = app_logger("api.image")

_cfg_path = Path(__file__).parents[3] / "config.yaml"


def _get_storage_dir() -> str:
    try:
        with open(_cfg_path, encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        return cfg.get("image", {}).get("storage_dir", "./storage/images")
    except Exception:
        return "./storage/images"


@router.post("/info")
async def image_info(file: UploadFile = File(...)):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(400, detail="文件格式错误：仅支持图片文件")
    raw = await file.read()
    tmp_path = get_temp_path(os.path.splitext(file.filename or ".png")[1])
    try:
        with open(tmp_path, "wb") as f:
            f.write(raw)
        img = Image.open(tmp_path)
        w, h = img.size
        fmt = img.format or "UNKNOWN"
        logger.info("Image info read", file=file.filename, width=w, height=h, format=fmt)
        return {"width": w, "height": h, "format": fmt, "file_size": len(raw)}
    except Exception as e:
        raise HTTPException(400, detail=f"图片解析失败：{e}")
    finally:
        os.unlink(tmp_path)


@router.post("/store")
async def store_image(file: UploadFile = File(...)):
    if not file.content_type or not file.content_type.startswith("image/"):
        raise HTTPException(400, detail="文件格式错误：仅支持图片文件")
    storage_dir = _get_storage_dir()
    os.makedirs(storage_dir, exist_ok=True)
    name = file.filename or f"unnamed_{len(os.listdir(storage_dir))}.png"
    dest = os.path.abspath(os.path.join(storage_dir, name))
    raw = await file.read()
    with open(dest, "wb") as f:
        f.write(raw)
    logger.info("Image stored", file=name, path=dest, size=len(raw))
    return {"path": dest, "filename": name, "size": len(raw)}


@router.get("/list")
async def list_images():
    storage_dir = _get_storage_dir()
    if not os.path.isdir(storage_dir):
        return {"images": []}
    images = []
    for fname in sorted(os.listdir(storage_dir)):
        fpath = os.path.join(storage_dir, fname)
        if os.path.isfile(fpath) and fname.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")):
            stat = os.stat(fpath)
            images.append({
                "filename": fname,
                "path": os.path.abspath(fpath),
                "size": stat.st_size,
                "mtime": stat.st_mtime,
            })
    logger.info("Images listed", count=len(images))
    return {"images": images}


@router.get("/{filename}")
async def serve_image(filename: str):
    storage_dir = _get_storage_dir()
    fpath = os.path.abspath(os.path.join(storage_dir, filename))
    if not fpath.startswith(os.path.abspath(storage_dir)):
        raise HTTPException(400, detail="非法文件名")
    if not os.path.isfile(fpath):
        raise HTTPException(404, detail="文件不存在")
    return FileResponse(fpath)


@router.delete("/{filename}")
async def delete_image(filename: str):
    storage_dir = _get_storage_dir()
    fpath = os.path.abspath(os.path.join(storage_dir, filename))
    if not fpath.startswith(os.path.abspath(storage_dir)):
        raise HTTPException(400, detail="非法文件名")
    if not os.path.isfile(fpath):
        raise HTTPException(404, detail="文件不存在")
    os.unlink(fpath)
    logger.info("Image deleted", file=filename)
    return {"status": "deleted", "filename": filename}
