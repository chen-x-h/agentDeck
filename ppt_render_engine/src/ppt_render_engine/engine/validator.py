from ppt_render_engine.models.schema import Slide, Shape
from ppt_render_engine.log_config import app_logger

logger = app_logger("engine.validator")

MIN_MARGIN = 91440
MAX_SHAPES_PER_SLIDE = 50


def validate_slide(slide: Slide, slide_index: int | None = None) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    if slide.preset and slide.preset not in _known_presets():
        warnings.append(f"Unknown preset '{slide.preset}'")
    e, w = _check_bounds(slide)
    errors.extend(e); warnings.extend(w)
    e, w = _check_overlaps(slide)
    errors.extend(e); warnings.extend(w)
    e, w = _check_count(slide)
    errors.extend(e); warnings.extend(w)
    e, w = _validate_containers(slide)
    errors.extend(e); warnings.extend(w)
    for msg in errors:
        logger.warning("Layout error", slide=slide_index, detail=msg)
    for msg in warnings:
        logger.warning("Layout warning", slide=slide_index, detail=msg)
    return errors, warnings


def _known_presets() -> set[str]:
    return {"cover", "section", "content", "two_column", "three_column", "comparison", "blank"}


def _check_bounds(slide: Slide) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    for s in slide.shapes:
        if s.left < -MIN_MARGIN:
            warnings.append(f"Shape '{s.id or s.role or '?'}' left={s.left} out of bounds")
        if s.top < -MIN_MARGIN:
            warnings.append(f"Shape '{s.id or s.role or '?'}' top={s.top} out of bounds")
        if s.left + s.width > slide.width + MIN_MARGIN:
            warnings.append(f"Shape '{s.id or s.role or '?'}' overflows right edge")
        if s.top + s.height > slide.height + MIN_MARGIN:
            warnings.append(f"Shape '{s.id or s.role or '?'}' overflows bottom edge")
        if s.width < 91440:
            errors.append(f"Shape '{s.id or s.role or '?'}' width={s.width} too small")
        if s.height < 91440:
            errors.append(f"Shape '{s.id or s.role or '?'}' height={s.height} too small")
    return errors, warnings


def _check_overlaps(slide: Slide) -> tuple[list[str], list[str]]:
    warnings = []
    shapes = list(slide.shapes)
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            if _overlap_area(a, b) > 0:
                area = _overlap_area(a, b)
                if area > min(a.width * a.height, b.width * b.height) * 0.3:
                    warnings.append(
                        f"Shapes '{a.id or a.role or i}' and '{b.id or b.role or j}' "
                        f"overlap significantly ({area:.0f} EMU²)"
                    )
    return [], warnings


def _check_count(slide: Slide) -> tuple[list[str], list[str]]:
    if len(slide.shapes) > MAX_SHAPES_PER_SLIDE:
        return [f"Slide has {len(slide.shapes)} shapes, max is {MAX_SHAPES_PER_SLIDE}"], []
    return [], []


def _validate_containers(slide: Slide) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    used_placeholders: set[str] = set()
    for s in slide.shapes:
        if s.children:
            if not s.placeholder and not s.role:
                errors.append(f"容器 '{s.id or '?'}' 没有 placeholder 或 role")
            if s.placeholder:
                if s.placeholder in used_placeholders:
                    errors.append(f"重复 placeholder '{s.placeholder}'（容器 '{s.id}'）")
                used_placeholders.add(s.placeholder)
            for ci, child in enumerate(s.children):
                if child.placeholder:
                    errors.append(f"容器 '{s.id}'.子[{ci}] 不应有 placeholder")
                if child.children:
                    errors.append(f"容器 '{s.id}'.子[{ci}] 不支持嵌套 children")
        elif s.placeholder:
            if s.placeholder in used_placeholders:
                errors.append(f"重复 placeholder '{s.placeholder}'（形状 '{s.id}'）")
            used_placeholders.add(s.placeholder)
    return errors, warnings


def _overlap_area(a: Shape, b: Shape) -> float:
    l = max(a.left, b.left)
    t = max(a.top, b.top)
    r = min(a.left + a.width, b.left + b.width)
    btm = min(a.top + a.height, b.top + b.height)
    if l < r and t < btm:
        return (r - l) * (btm - t)
    return 0
