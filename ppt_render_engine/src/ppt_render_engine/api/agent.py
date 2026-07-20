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


def _validate_slide(slide: dict, page: int = 0) -> tuple[list[str], list[str]]:
    """Validate a slide dict and all its shapes. Returns (errors, warnings)."""
    errors, warnings = [], []
    try:
        Slide(**slide)
    except ValidationError as e:
        errors.extend(_fmt_val_err(f"第{page}页", err) for err in e.errors())
    for i, shape in enumerate(slide.get("shapes", [])):
        errors.extend(_validate_shape(shape, i))
        warn = _warn_text_overflow(shape, page, str(i))
        if warn:
            warnings.append(warn)
        for ci, child in enumerate(shape.get("children", [])):
            child_warn = _warn_text_overflow(child, page, f"{i}.children[{ci}]")
            if child_warn:
                warnings.append(child_warn)
    errors.extend(_check_container_errors(slide, page))
    warnings.extend(_warn_overlaps(slide, page))
    return errors, warnings


def _warn_text_overflow(shape: dict, page: int, idx: str) -> str | None:
    if shape.get("type") not in ("textbox", "shape"):
        return None
    tc = shape.get("text_content") or []
    w = shape.get("width") or 1
    h = shape.get("height") or 1
    # 百分比转 EMU 以便估算
    slide_w = 12192000
    slide_h = 6858000
    if 0 < w <= 100:
        w = w / 100 * slide_w
    if 0 < h <= 100:
        h = h / 100 * slide_h
    total_chars = 0
    cjk_count = 0
    for p in tc:
        for r in p.get("runs", []):
            txt = r.get("text", "")
            total_chars += len(txt)
            cjk_count += sum(1 for c in txt if '\u4e00' <= c <= '\u9fff' or '\u3000' <= c <= '\u303f')
    max_fs = max(
        (r.get("font_size") or 177800 for p in tc for r in p.get("runs", [])),
        default=177800,
    )
    cjk_ratio = cjk_count / max(total_chars, 1)
    char_w = max_fs * (0.9 if cjk_ratio > 0.3 else 0.55)
    est_w = total_chars * char_w
    est_h = sum(
        max(r.get("font_size", 177800) for r in p.get("runs", [])) * 1.4
        for p in tc if p.get("runs")
    ) or max_fs * 1.4
    ratio_w = est_w / w
    ratio_h = est_h / h
    if ratio_w > 1.2 or ratio_h > 1.2:
        return (
            f"第{page}页.形状[{idx}]: 文本溢出 "
            f"(估算需 {int(est_w)}x{int(est_h)} EMU，框 {w}x{h} EMU，"
            f"{'宽' if ratio_w > 1.2 else ''}{'高' if ratio_h > 1.2 else ''})。"
            f"建议增大框尺寸、减小字号或缩短文字"
        )
    return None


def _check_container_errors(slide: dict, page: int) -> list[str]:
    errors = []
    used_ph: set[str] = set()
    for i, shape in enumerate(slide.get("shapes", [])):
        if shape.get("children"):
            ref = shape.get("placeholder")
            if ref:
                if ref in used_ph:
                    errors.append(f"第{page}页: 多个元素引用同一 placeholder '{ref}'，请合并到一个容器中")
                used_ph.add(ref)
            for ci, child in enumerate(shape.get("children", [])):
                if child.get("placeholder"):
                    errors.append(f"第{page}页.形状[{i}]容器.子[{ci}]: 子元素不应有 placeholder")
                if child.get("children"):
                    errors.append(f"第{page}页.形状[{i}]容器.子[{ci}]: 不支持嵌套 children")
        elif shape.get("placeholder"):
            ref = shape["placeholder"]
            if ref in used_ph:
                errors.append(f"第{page}页: 多个元素引用同一 placeholder '{ref}'，请合并到一个容器中")
            used_ph.add(ref)
    return errors


def _warn_overlaps(slide: dict, page: int) -> list[str]:
    """Check for overlapping shapes and return warnings."""
    warnings = []
    shapes = slide.get("shapes", [])
    for i, a in enumerate(shapes):
        for b in shapes[i + 1:]:
            ax = a.get("left", 0)
            ay = a.get("top", 0)
            aw = a.get("width", 0)
            ah = a.get("height", 0)
            bx = b.get("left", 0)
            by = b.get("top", 0)
            bw = b.get("width", 0)
            bh = b.get("height", 0)
            l = max(ax, bx)
            t = max(ay, by)
            r = min(ax + aw, bx + bw)
            btm = min(ay + ah, by + bh)
            if l >= r or t >= btm:
                continue
            area = (r - l) * (btm - t)
            ratio = area / min(aw * ah, bw * bh) if min(aw * ah, bw * bh) > 0 else 0
            if ratio > 0.2:
                a_id = a.get("id", i)
                b_id = b.get("id", i + 1)
                warnings.append(
                    f"第{page}页: 形状「{a_id}」与「{b_id}」重叠 "
                    f"(重叠比 {ratio:.0%}，建议调整位置)"
                )
    return warnings


def _fmt_val_err(prefix: str, err: dict) -> str:
    loc = ".".join(str(x) for x in err["loc"])
    val = err.get("input")
    return f"{prefix}.{loc}: {err['msg']} (got {val!r})"


def _reject_if_errors(errors: list[str]):
    if errors:
        msg = "数据校验失败，请修正后重试：\n" + "\n".join(errors)
        raise HTTPException(400, detail=msg)


# ---------------------------------------------------------------------------
# 1. POST /agent/sync  —  Upload JSON, overwrite sync file, fix ID prefixes
# ---------------------------------------------------------------------------

@router.post("/sync")
async def agent_sync(body: dict = Body(...)):
    """Upload a full JSON presentation to overwrite the sync file.
    Automatically fixes shape ID prefixes to match each slide's page number.
    Returns both the saved data and validation warnings for model feedback.
    """
    if "slides" not in body:
        raise HTTPException(400, detail="缺少 'slides' 字段")
    for slide in body.get("slides", []):
        for i, shape in enumerate(slide.get("shapes", [])):
            slide["shapes"][i] = _normalize_shape(shape)
    errors, warnings = [], []
    for i, slide in enumerate(body.get("slides", [])):
        e, w = _validate_slide(slide, i)
        errors.extend(e)
        warnings.extend(w)
    _reject_if_errors(errors)
    _reindex_ids(body)
    _save_sync(body)
    logger.info("Agent sync upload", slides=len(body.get("slides", [])))
    return {
        "status": "ok",
        "data": body,
        "warnings": warnings,
        "slides": len(body.get("slides", [])),
    }


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
    errors, warnings = _validate_slide(existing, page)
    _reject_if_errors(errors)
    _reindex_ids(data)
    _save_sync(data)
    logger.info("Agent replaced page", page=page)
    return {"status": "ok", "page": page, "shapes": len(existing.get("shapes", [])), "warnings": warnings}


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
    _reject_if_errors(_validate_shape(normalized))
    for i, s in enumerate(shapes):
        if s.get("id") == id:
            normalized["id"] = id
            shapes[i] = normalized
            _save_sync(data)
            logger.info("Agent replaced element by id", page=page, id=id)
            errors = _check_container_errors(slides[page], page)
            _reject_if_errors(errors)
            return {
                "status": "ok",
                "page": page,
                "id": id,
                "action": "updated",
            }
    normalized["id"] = id
    shapes.append(normalized)
    _save_sync(data)
    logger.info("Agent appended new element", page=page, id=id)
    errors = _check_container_errors(slides[page], page)
    _reject_if_errors(errors)
    return {
        "status": "ok",
        "page": page,
        "id": id,
        "action": "appended",
    }


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
    template_name: str = Query(..., description="模板名，必填，如 标准1"),
    color_scheme: str = Query(..., description="配色名，必填，如 深邃蓝"),
):
    data = _load_sync()
    try:
        pres = Presentation(**data)
    except Exception as e:
        logger.error("Agent export schema validation failed", error=f"{e}\n{traceback.format_exc()}")
        raise HTTPException(400, detail=f"同步文件数据结构校验失败，可能需要先重建：{e}")
    pres.template_name = template_name
    pres.color_scheme = color_scheme
    # 模板模式检查：对照版式的实际占位符定义做校验
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
                        f"缺少 placeholder。所有顶层形状都必须关联一个版式占位符，"
                        f"标明它属于哪个槽位。可用占位符: "
                        + ", ".join(f"{p['idx']}={p['name']}" for p in lay.get("placeholders", []))
                    ))
                # 检查占位符类型是否为系统槽位
                ph_idx = sh.placeholder
                matched_ph = next((p for p in lay.get("placeholders", []) if str(p["idx"]) == ph_idx), None)
                if matched_ph and matched_ph.get("type") in ("DT", "FTR", "SLD_NUM"):
                    raise HTTPException(400, detail=(
                        f"套用模板「{pres.template_name}」第{si}页形状「{sh.id}」"
                        f"使用了系统占位符 idx={ph_idx}({matched_ph['type']})，"
                        f"该槽位不是内容区（日期/页脚/编号），请使用内容占位符"
                    ))
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
# 10. GET /agent/example  —  Return a minimal example presentation
# ---------------------------------------------------------------------------

_EXAMPLE = {
    "title": "示例演示文稿",
    "slides": [
        {
            "layout_id": "标题幻灯片",
            "shapes": [
                {"id": "0-标题", "type": "textbox", "placeholder": "0", "text_content": [{"alignment": "center", "runs": [{"text": "示例标题", "font_size": 540000, "bold": True}]}]},
                {"id": "0-副标题", "type": "textbox", "placeholder": "1", "text_content": [{"alignment": "center", "runs": [{"text": "副标题文字", "font_size": 228600}]}]}
            ]
        },
        {
            "layout_id": "标题和内容",
            "shapes": [
                {"id": "1-标题", "type": "textbox", "placeholder": "0", "text_content": [{"runs": [{"text": "内容页标题", "font_size": 360000, "bold": True}]}]},
                {"id": "1-正文", "type": "textbox", "placeholder": "1", "text_content": [{"runs": [{"text": "正文内容", "font_size": 177800}]}]}
            ]
        }
    ]
}


@router.get("/example")
async def agent_example():
    return _EXAMPLE


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


@router.get("/schemes/{name}")
async def agent_get_scheme(name: str):
    from ppt_render_engine.core.color_scheme import get_color_scheme_manager
    csm = get_color_scheme_manager()
    colors = csm.get_colors(name)
    if not colors:
        raise HTTPException(404, detail=f"配色方案不存在: {name}")
    return {"name": name, "colors": colors}


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
