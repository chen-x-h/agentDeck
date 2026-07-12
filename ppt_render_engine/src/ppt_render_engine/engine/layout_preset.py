from ppt_render_engine.models.schema import Slide


class LayoutZone:
    def __init__(self, role: str, left_pct: float, top_pct: float, width_pct: float, height_pct: float):
        self.role = role
        self.left_pct = left_pct
        self.top_pct = top_pct
        self.width_pct = width_pct
        self.height_pct = height_pct

    def to_emu(self, slide_w: float, slide_h: float) -> tuple[float, float, float, float]:
        l = slide_w * self.left_pct / 100
        t = slide_h * self.top_pct / 100
        w = slide_w * self.width_pct / 100
        h = slide_h * self.height_pct / 100
        return (l, t, w, h)


PRESETS: dict[str, list[LayoutZone]] = {
    "cover": [
        LayoutZone("bg_bar", 0, 0, 100, 100),
        LayoutZone("title", 12, 32, 76, 22),
        LayoutZone("subtitle", 18, 56, 64, 12),
        LayoutZone("footer_line", 12, 82, 76, 2),
    ],
    "section": [
        LayoutZone("bg_bar", 0, 0, 8, 100),
        LayoutZone("title", 18, 34, 64, 32),
    ],
    "content": [
        LayoutZone("title", 8, 6, 84, 12),
        LayoutZone("body", 8, 22, 84, 72),
    ],
    "two_column": [
        LayoutZone("title", 8, 6, 84, 10),
        LayoutZone("left_col", 8, 20, 40, 74),
        LayoutZone("right_col", 52, 20, 40, 74),
    ],
    "three_column": [
        LayoutZone("title", 8, 6, 84, 10),
        LayoutZone("col_1", 4, 20, 29, 74),
        LayoutZone("col_2", 35.5, 20, 29, 74),
        LayoutZone("col_3", 67, 20, 29, 74),
    ],
    "comparison": [
        LayoutZone("title", 8, 6, 84, 10),
        LayoutZone("left_heading", 8, 20, 40, 10),
        LayoutZone("right_heading", 52, 20, 40, 10),
        LayoutZone("left_body", 8, 32, 40, 62),
        LayoutZone("right_body", 52, 32, 40, 62),
    ],
}


def resolve_preset(slide: Slide) -> Slide:
    if not slide.preset or slide.preset not in PRESETS:
        return slide
    zones = PRESETS[slide.preset]
    resolved = slide.model_copy(deep=True)
    used_indices: set[int] = set()
    for i, shape in enumerate(resolved.shapes):
        if not shape.role:
            continue
        zone = _find_zone(zones, shape.role)
        if zone is None:
            continue
        l, t, w, h = zone.to_emu(slide.width, slide.height)
        shape.left = l
        shape.top = t
        shape.width = w
        shape.height = h
        used_indices.add(i)
    return resolved


def _find_zone(zones: list[LayoutZone], role: str) -> LayoutZone | None:
    for z in zones:
        if z.role == role:
            return z
    return None
