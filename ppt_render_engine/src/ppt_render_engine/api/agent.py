import os
import re
import json
import yaml
import traceback
from pathlib import Path
from fastapi import APIRouter, HTTPException, Query, Body
from fastapi.responses import PlainTextResponse
from ppt_render_engine.log_config import app_logger
from pydantic import ValidationError
from ppt_render_engine.models.schema import Presentation, Slide, Shape
from ppt_render_engine.core.pptx_builder import build_pptx


def _overlap(a: Shape, b: Shape) -> bool:
    return (a.left < b.left + b.width and a.left + a.width > b.left and
            a.top < b.top + b.height and a.top + a.height > b.top)


def _overlap_ratio(a: Shape, b: Shape) -> float:
    x1 = max(a.left, b.left)
    y1 = max(a.top, b.top)
    x2 = min(a.left + a.width, b.left + b.width)
    y2 = min(a.top + a.height, b.top + b.height)
    overlap = max(0, x2 - x1) * max(0, y2 - y1)
    area_a = a.width * a.height or 1
    return overlap / area_a

logger = app_logger("api.agent")

router = APIRouter(prefix="/agent", tags=["Agent"])

PROJECT_ROOT = Path(__file__).resolve().parents[3]


def _load_config():
    cfg_path = PROJECT_ROOT / "config.yaml"
    try:
        with open(cfg_path, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}


def _get_sync_path():
    cfg = _load_config()
    rel = cfg.get("sync", {}).get("json_path", "./sync/presentation.json")
    return os.path.abspath(os.path.join(PROJECT_ROOT, rel))


def _load_sync():
    path = _get_sync_path()
    if not os.path.isfile(path):
        raise HTTPException(404, detail="同步文件不存在")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise HTTPException(400, detail=f"同步文件解析失败: {e}")


def _save_sync(data: dict):
    path = _get_sync_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    logger.info("Agent sync saved")


def _get_image_dir():
    cfg = _load_config()
    rel = cfg.get("image", {}).get("storage_dir", "./sync/images")
    return os.path.abspath(os.path.join(PROJECT_ROOT, rel))


_ID_PATTERN = re.compile(r"^(\d+)-")


def _normalize_id(shape_id: str, page: int) -> str:
    """Replace the numeric prefix in an ID to match the given page number."""
    m = _ID_PATTERN.match(shape_id)
    if m:
        suffix = shape_id[m.end():]
        return f"{page}-{suffix}"
    return shape_id


def _reindex_ids(data: dict):
    """Rewrite the numeric prefix of every shape ID to match its slide index."""
    for i, slide in enumerate(data.get("slides", [])):
        for shape in slide.get("shapes", []):
            sid = shape.get("id")
            if sid:
                shape["id"] = _normalize_id(sid, i)


def _normalize_shape(shape: dict):
    """Auto-fix common model mistakes in shape data."""
    if shape.get("type") == "text":
        shape["type"] = "textbox"
    if shape.get("type") == "textbox":
        if "text" in shape and "text_content" not in shape:
            shape["text_content"] = [{
                "alignment": shape.pop("align", "left"),
                "runs": [{
                    "text": shape.pop("text"),
                    "font_size": shape.pop("font_size", None) or None,
                    "font_name": shape.pop("font_name", None),
                    "bold": shape.pop("bold", False),
                    "italic": shape.pop("italic", False),
                    "underline": shape.pop("underline", False),
                    "color": shape.pop("color", None),
                    "hyperlink": shape.pop("hyperlink", None),
                }],
            }]
    return shape


def _validate_shape(shape: dict, index: int = 0) -> list[str]:
    """Validate a single shape dict against Pydantic schema. Returns list of errors."""
    try:
        Shape(**shape)
    except ValidationError as e:
        return [_fmt_val_err(f"形状[{index}]", err) for err in e.errors()]
    return []


def _validate_slide(slide: dict, page: int = 0) -> list[str]:
    """Validate a slide dict and all its shapes. Returns list of errors."""
    errors = []
    try:
        Slide(**slide)
    except ValidationError as e:
        errors.extend(_fmt_val_err(f"第{page}页", err) for err in e.errors())
    for i, shape in enumerate(slide.get("shapes", [])):
        errors.extend(_validate_shape(shape, i))
        _warn_text_overflow(shape, page, i, errors)
    return errors


def _warn_text_overflow(shape: dict, page: int, idx: int, errors: list):
    if shape.get("type") not in ("textbox", "shape"):
        return
    if shape.get("placeholder"):
        return
    tc = shape.get("text_content") or []
    w = shape.get("width") or 1
    h = shape.get("height") or 1
    total_chars = sum(len(r.get("text", "")) for p in tc for r in p.get("runs", []))
    max_fs = max(
        (r.get("font_size") or 177800 for p in tc for r in p.get("runs", [])),
        default=177800,
    )
    est_w = total_chars * max_fs * 0.6
    est_h = sum(p.get("runs", []) and max(r.get("font_size", 177800) for r in p.get("runs", [])) * 1.4 or 0 for p in tc)
    if est_w > w * 1.2 or est_h > h * 1.2:
        errors.append(
            f"第{page}页.形状[{idx}]: 文本可能溢出 (约{int(est_w)}x{int(est_h)} EMU, "
            f"框 {w}x{h} EMU)。建议增大框尺寸、减小字号或缩短文字"
        )


def _fmt_val_err(prefix: str, err: dict) -> str:
    loc = ".".join(str(x) for x in err["loc"])
    val = err.get("input")
    return f"{prefix}.{loc}: {err['msg']} (got {val!r})"


def _raise_if_invalid(errors: list[str]):
    if errors:
        msg = "数据校验失败，请修正后重试：\n" + "\n".join(errors)
        raise HTTPException(400, detail=msg)


# ---------------------------------------------------------------------------
# 1. POST /agent/sync  —  Upload JSON, overwrite sync file, fix ID prefixes
# ---------------------------------------------------------------------------

@router.post("/sync", response_class=PlainTextResponse)
async def agent_sync(body: dict = Body(...)):
    """Upload a full JSON presentation to overwrite the sync file.
    Automatically fixes shape ID prefixes to match each slide's page number.
    """
    if "slides" not in body:
        raise HTTPException(400, detail="缺少 'slides' 字段")
    for slide in body.get("slides", []):
        for i, shape in enumerate(slide.get("shapes", [])):
            slide["shapes"][i] = _normalize_shape(shape)
    errors = []
    for i, slide in enumerate(body.get("slides", [])):
        errors.extend(_validate_slide(slide, i))
    _raise_if_invalid(errors)
    _reindex_ids(body)
    _save_sync(body)
    logger.info("Agent sync upload", slides=len(body.get("slides", [])))
    return json.dumps(body, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 2. GET /agent/page/{page}  —  Return a single page's JSON
# ---------------------------------------------------------------------------

@router.get("/page/{page}", response_class=PlainTextResponse)
async def agent_get_page(page: int):
    data = _load_sync()
    slides = data.get("slides", [])
    if page < 0 or page >= len(slides):
        raise HTTPException(400, detail=f"页码越界：共 {len(slides)} 页")
    return json.dumps(slides[page], ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 3. GET /agent/element  —  Return element by page + id
# ---------------------------------------------------------------------------

@router.get("/element", response_class=PlainTextResponse)
async def agent_get_element(
    page: int = Query(..., ge=0),
    id: str = Query(..., min_length=1),
):
    data = _load_sync()
    slides = data.get("slides", [])
    if page < 0 or page >= len(slides):
        raise HTTPException(400, detail=f"页码越界：共 {len(slides)} 页")
    for s in slides[page].get("shapes", []):
        if s.get("id") == id:
            return json.dumps(s, ensure_ascii=False, indent=2)
    raise HTTPException(404, detail=f"未找到 id='{id}' 的形状")


# ---------------------------------------------------------------------------
# 4. DELETE /agent/page/{page}  —  Delete page, re-index all IDs
# ---------------------------------------------------------------------------

@router.delete("/page/{page}")
async def agent_delete_page(page: int):
    data = _load_sync()
    slides = data.get("slides", [])
    if page < 0 or page >= len(slides):
        raise HTTPException(400, detail=f"页码越界：共 {len(slides)} 页")
    deleted = slides.pop(page)
    _reindex_ids(data)
    _save_sync(data)
    logger.info("Agent deleted page", page=page)
    return {
        "status": "deleted",
        "page": page,
        "shapes_removed": len(deleted.get("shapes", [])),
        "slides_remaining": len(slides),
    }


# ---------------------------------------------------------------------------
# 5. GET /agent/ids  —  Return flat list of all element IDs
# ---------------------------------------------------------------------------

@router.get("/ids")
async def agent_list_ids():
    data = _load_sync()
    result = []
    for i, slide in enumerate(data.get("slides", [])):
        for s in slide.get("shapes", []):
            result.append({
                "page": i,
                "id": s.get("id"),
                "type": s.get("type"),
                "role": s.get("role"),
            })
    return {"shapes": result}


# ---------------------------------------------------------------------------
# 6. POST /agent/page  —  Insert an empty page at position
# ---------------------------------------------------------------------------

@router.post("/page")
async def agent_insert_page(
    position: int = Query(..., ge=0, description="插入位置索引，0=最前面"),
):
    data = _load_sync()
    slides = data.get("slides", [])
    if position < 0 or position > len(slides):
        raise HTTPException(400, detail=f"插入位置越界：共 {len(slides)} 页")
    slides.insert(position, {"shapes": []})
    _reindex_ids(data)
    _save_sync(data)
    logger.info("Agent inserted page", position=position, total=len(slides))
    return {"status": "ok", "position": position, "slides_total": len(slides)}


# ---------------------------------------------------------------------------
# 7. PUT /agent/page/{page}  —  Replace a whole page
# ---------------------------------------------------------------------------

@router.put("/page/{page}")
async def agent_put_page(
    page: int,
    body: dict = Body(...),
):
    data = _load_sync()
    slides = data.get("slides", [])
    if page < 0 or page > len(slides):
        raise HTTPException(400, detail=f"页码越界：共 {len(slides)} 页，不能跳跃创建")
    if not isinstance(body, dict):
        raise HTTPException(400, detail=f"请求体必须是 JSON 对象，收到 {type(body).__name__}")
    if page == len(slides):
        slides.append({"shapes": []})
    existing = slides[page]
    for k, v in body.items():
        if k == "shapes":
            existing["shapes"] = [_normalize_shape(s) for s in (v or [])]
        else:
            existing[k] = v
    errors = _validate_slide(existing, page)
    _raise_if_invalid(errors)
    _reindex_ids(data)
    _save_sync(data)
    logger.info("Agent replaced page", page=page)
    return {"status": "ok", "page": page, "shapes": len(existing.get("shapes", []))}


# ---------------------------------------------------------------------------
# 8. PUT /agent/element  —  Replace element by page + id (not-found → append)
# ---------------------------------------------------------------------------

@router.put("/element")
async def agent_put_element(
    page: int = Query(..., ge=0),
    id: str = Query(..., min_length=1),
    body: dict = Body(...),
):
    data = _load_sync()
    slides = data.get("slides", [])
    if page < 0 or page >= len(slides):
        raise HTTPException(400, detail=f"页码越界：共 {len(slides)} 页")
    shapes = slides[page].get("shapes", [])
    normalized = _normalize_shape(dict(body))
    _raise_if_invalid(_validate_shape(normalized))
    for i, s in enumerate(shapes):
        if s.get("id") == id:
            normalized["id"] = id
            shapes[i] = normalized
            _save_sync(data)
            logger.info("Agent replaced element by id", page=page, id=id)
            return {"status": "ok", "page": page, "id": id, "action": "updated"}
    normalized["id"] = id
    shapes.append(normalized)
    _save_sync(data)
    logger.info("Agent appended new element", page=page, id=id)
    return {"status": "ok", "page": page, "id": id, "action": "appended"}


# ---------------------------------------------------------------------------
# 9. POST /agent/export  —  Build PPTX from sync, save to specified path
# ---------------------------------------------------------------------------

@router.post("/reset")
async def agent_reset():
    """清空同步区，写入包含一页空白幻灯片的演示文稿，用于新建 PPT 重新开始。"""
    empty = {"title": "", "slides": [{"shapes": []}]}
    _save_sync(empty)
    logger.info("Agent reset sync area")
    return {"status": "ok", "message": "同步区已重置，默认已创建第 0 页空白幻灯片，可用 write_page 写入内容"}


@router.post("/export")
async def agent_export(
    output_path: str = Query(..., description="完整输出路径，如 D:/output/my.pptx"),
    template_name: str = Query(None, description="覆盖模板名"),
    color_scheme: str = Query(None, description="覆盖配色名"),
):
    data = _load_sync()
    try:
        pres = Presentation(**data)
    except Exception as e:
        logger.error("Agent export schema validation failed", error=f"{e}\n{traceback.format_exc()}")
        raise HTTPException(400, detail=f"同步文件数据结构校验失败，可能需要先重建：{e}")
    if template_name is not None:
        pres.template_name = template_name
    if color_scheme is not None:
        pres.color_scheme = color_scheme
    # 模板模式检查：对照版式的实际占位符定义做校验
    if pres.template_name:
        from ppt_render_engine.core.template import get_template_manager
        tm = get_template_manager()
        try:
            layouts = tm.list_layouts(pres.template_name)
        except Exception:
            layouts = []
        layout_index: dict[str, dict] = {}
        for lay in layouts:
            layout_index[str(lay["index"])] = lay
            layout_index[lay["name"]] = lay
        names = sorted(layout_index.keys())
        for si, slide in enumerate(pres.slides):
            lid = slide.layout_id or ""
            lay = layout_index.get(lid) or layout_index.get(str(lid))
            if not lay and lid:
                raise HTTPException(400, detail=(
                    f"套用模板「{pres.template_name}」第{si}页的 layout_id「{lid}」"
                    f"在模板中不存在。可用版式: {names}"
                ))
            for sh in slide.shapes:
                if not sh.placeholder:
                    raise HTTPException(400, detail=(
                        f"套用模板「{pres.template_name}」第{si}页形状「{sh.id}」"
                        f"缺少 placeholder。所有元素都必须关联一个版式占位符，"
                        f"标明它属于哪个槽位。可用占位符: "
                        + ", ".join(f"{p['idx']}={p['name']}" for p in lay.get("placeholders", []))
                    ))
            # 检查形状间重叠
            for i, a in enumerate(slide.shapes):
                for b in slide.shapes[i + 1:]:
                    if not _overlap(a, b):
                        continue
                    ratio = _overlap_ratio(a, b)
                    if ratio > 0.3:
                        logger.warning(
                            f"套用模板第{si}页形状重叠",
                            a=a.id, b=b.id, overlap=f"{ratio:.0%}"
                        )
    abs_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    try:
        build_pptx(pres, abs_path)
    except Exception as e:
        logger.error("Agent export failed", error=f"{e}\n{traceback.format_exc()}")
        raise HTTPException(500, detail=f"PPTX 生成失败: {e}")
    logger.info("Agent exported PPTX", path=abs_path, slides=len(pres.slides))
    return {
        "status": "ok",
        "path": abs_path,
        "slides": len(pres.slides),
        "file_size": os.path.getsize(abs_path),
    }


# ---------------------------------------------------------------------------
# 10. GET /agent/example  —  Return a complete example for the model to learn
# ---------------------------------------------------------------------------

_EXAMPLE_FILE = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'storage', 'examples', 'default.json')


@router.get("/example")
async def agent_example():
    try:
        with open(_EXAMPLE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        raise HTTPException(500, detail=f"读取示例文件失败: {e}")


# ---------------------------------------------------------------------------
# Supporting endpoints
# ---------------------------------------------------------------------------

@router.get("/templates")
async def agent_list_templates():
    from ppt_render_engine.core.template import get_template_manager
    tm = get_template_manager()
    names = tm.list_templates()
    result = {}
    for n in names:
        layouts = tm.list_layouts(n)
        result[n] = {"layouts": layouts} if layouts else {}
    return {"templates": result}


@router.get("/schemes")
async def agent_list_schemes():
    from ppt_render_engine.core.color_scheme import get_color_scheme_manager
    csm = get_color_scheme_manager()
    return {"schemes": csm.list_schemes(), "default": csm.get_default_name()}


@router.get("/images")
async def agent_list_images():
    storage = _get_image_dir()
    if not os.path.isdir(storage):
        return {"images": []}
    images = []
    for fname in sorted(os.listdir(storage)):
        fpath = os.path.join(storage, fname)
        if os.path.isfile(fpath) and fname.lower().endswith(
            (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp")
        ):
            stat = os.stat(fpath)
            images.append({
                "filename": fname,
                "size": stat.st_size,
                "mtime": stat.st_mtime,
                "url": f"http://localhost:8000/image/{fname}",
            })
    return {"images": images}


@router.post("/images/upload")
async def agent_upload_image(
    filename: str = Query(..., description="文件名，如 sample.png"),
    data: str = Body(..., description="Base64 编码的图片数据"),
):
    import base64
    storage = _get_image_dir()
    os.makedirs(storage, exist_ok=True)
    dest = os.path.abspath(os.path.join(storage, filename))
    try:
        raw = base64.b64decode(data)
    except Exception as e:
        raise HTTPException(400, detail=f"Base64 解码失败: {e}")
    with open(dest, "wb") as f:
        f.write(raw)
    logger.info("Agent uploaded image", filename=filename, size=len(raw))
    return {"status": "ok", "filename": filename, "size": len(raw), "path": dest}
