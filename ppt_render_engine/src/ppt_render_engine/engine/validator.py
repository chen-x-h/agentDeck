from ppt_render_engine.models.schema import Slide, Shape
from ppt_render_engine.log_config import app_logger

logger = app_logger("engine.validator")

MIN_MARGIN = 91440
MAX_SHAPES_PER_SLIDE = 50


def validate_slide(slide: Slide, slide_index: int | None = None) -> list[str]:
    issues = []
    if slide.preset and slide.preset not in _known_presets():
        issues.append(f"Unknown preset '{slide.preset}'")
    issues.extend(_check_bounds(slide))
    issues.extend(_check_overlaps(slide))
    issues.extend(_check_count(slide))
    for msg in issues:
        logger.warning("Layout issue", slide=slide_index, detail=msg)
    return issues


def _known_presets() -> set[str]:
    return {"cover", "section", "content", "two_column", "three_column", "comparison", "blank"}


def _check_bounds(slide: Slide) -> list[str]:
    issues = []
    for s in slide.shapes:
        if s.left < -MIN_MARGIN:
            issues.append(f"Shape '{s.id or s.role or '?'}' left={s.left} out of bounds")
        if s.top < -MIN_MARGIN:
            issues.append(f"Shape '{s.id or s.role or '?'}' top={s.top} out of bounds")
        if s.left + s.width > slide.width + MIN_MARGIN:
            issues.append(f"Shape '{s.id or s.role or '?'}' overflows right edge")
        if s.top + s.height > slide.height + MIN_MARGIN:
            issues.append(f"Shape '{s.id or s.role or '?'}' overflows bottom edge")
        if s.width < 91440:
            issues.append(f"Shape '{s.id or s.role or '?'}' width={s.width} too small")
        if s.height < 91440:
            issues.append(f"Shape '{s.id or s.role or '?'}' height={s.height} too small")
    return issues


def _check_overlaps(slide: Slide) -> list[str]:
    issues = []
    shapes = list(slide.shapes)
    for i in range(len(shapes)):
        for j in range(i + 1, len(shapes)):
            a, b = shapes[i], shapes[j]
            if _overlap_area(a, b) > 0:
                area = _overlap_area(a, b)
                if area > min(a.width * a.height, b.width * b.height) * 0.3:
                    issues.append(
                        f"Shapes '{a.id or a.role or i}' and '{b.id or b.role or j}' "
                        f"overlap significantly ({area:.0f} EMU²)"
                    )
    return issues


def _check_count(slide: Slide) -> list[str]:
    if len(slide.shapes) > MAX_SHAPES_PER_SLIDE:
        return [f"Slide has {len(slide.shapes)} shapes, max is {MAX_SHAPES_PER_SLIDE}"]
    return []


def _overlap_area(a: Shape, b: Shape) -> float:
    l = max(a.left, b.left)
    t = max(a.top, b.top)
    r = min(a.left + a.width, b.left + b.width)
    btm = min(a.top + a.height, b.top + b.height)
    if l < r and t < btm:
        return (r - l) * (btm - t)
    return 0
