from pydantic import BaseModel, Field
from typing import List, Optional
from enum import Enum


class ShapeType(str, Enum):
    TEXTBOX = "textbox"
    IMAGE = "image"
    TABLE = "table"
    SHAPE = "shape"


class TextRun(BaseModel):
    text: str
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bold: Optional[bool] = False
    italic: Optional[bool] = False
    underline: Optional[bool] = False
    color: Optional[str] = None
    hyperlink: Optional[str] = None


class Paragraph(BaseModel):
    runs: List[TextRun]
    alignment: Optional[str] = "left"
    line_spacing: Optional[float] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    indent_level: Optional[int] = 0


class ImageContent(BaseModel):
    data: Optional[str] = None
    url: Optional[str] = None
    path: Optional[str] = None
    width: Optional[float] = None
    height: Optional[float] = None
    crop_left: Optional[float] = None
    crop_top: Optional[float] = None
    crop_right: Optional[float] = None
    crop_bottom: Optional[float] = None


class CellContent(BaseModel):
    text: str
    colspan: Optional[int] = 1
    rowspan: Optional[int] = 1
    font_size: Optional[float] = None
    bold: Optional[bool] = False
    alignment: Optional[str] = "left"
    background_color: Optional[str] = None


class TableContent(BaseModel):
    rows: int
    cols: int
    cells: List[List[CellContent]]


class Shape(BaseModel):
    id: Optional[str] = None
    type: ShapeType
    left: float = 0
    top: float = 0
    width: float = 0
    height: float = 0
    z_order: Optional[int] = 0
    rotation: Optional[float] = 0
    role: Optional[str] = None
    auto_shape_type: Optional[str] = None
    text_content: Optional[List[Paragraph]] = None
    image_content: Optional[ImageContent] = None
    table_content: Optional[TableContent] = None
    background_color: Optional[str] = None
    border_color: Optional[str] = None
    border_width: Optional[float] = None
    border_style: Optional[str] = "solid"
    shadow: Optional[bool] = False
    placeholder: Optional[str] = None


class Slide(BaseModel):
    id: Optional[int] = None
    width: float = 12192000
    height: float = 6858000
    background_color: Optional[str] = None
    background_image: Optional[str] = None
    master_id: Optional[str] = None
    layout_id: Optional[str] = None
    preset: Optional[str] = None
    shapes: List[Shape] = Field(default_factory=list)


class SlideMaster(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    slide: Optional[Slide] = None


class Presentation(BaseModel):
    title: Optional[str] = None
    slide_width: float = 12192000
    slide_height: float = 6858000
    template_name: Optional[str] = None
    color_scheme: Optional[str] = None
    masters: Optional[List[SlideMaster]] = None
    slides: List[Slide]


EMU_PER_INCH = 914400
EMU_PER_CM = 360000
DEFAULT_SLIDE_WIDTH = 12192000
DEFAULT_SLIDE_HEIGHT = 6858000
