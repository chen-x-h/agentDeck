"""MCP server for PPT Render Engine (SSE transport).

Two modes:
  - remote: step-by-step (read/write/delete/insert/reset). Default.
  - local:  local JSON file workflow (sync + export + utilities).

Run with: uv run python -m mcp_server --mode remote
"""

import json
import os
import difflib
import yaml
import logging
from pathlib import Path
from typing import Optional
import httpx
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("mcp")

_ROOT = Path(__file__).resolve().parents[3]
_CFG_PATH = _ROOT / "config.yaml"


def _load_cfg() -> dict:
    try:
        with open(_CFG_PATH, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}


_cfg = _load_cfg()
BASE_URL = _cfg.get("render_engine", {}).get("url", "http://127.0.0.1:8000")
MCP_HOST = _cfg.get("mcp_server", {}).get("host", "0.0.0.0")
MCP_PORT = _cfg.get("mcp_server", {}).get("port", 8010)
_UI_URL = BASE_URL


# ============================================================
#  shape reference catalog
# ============================================================

_SHAPE_CATALOG = """## 形状类型参考 (Shape Reference)

所有形状必须用标准嵌套 JSON 格式，不可用扁平字段。

> **关键规则（模板模式）**：每个版式占位符在一页内只能被**一个顶层形状**引用。
> 如果需要一个区域放多个元素（如色块+文字），用**容器（children）**包起来，
> 容器拿 placeholder，子元素**不能**拿。

### 1. textbox — 文本框
无模板（手写坐标）:
{
  "type": "textbox",
  "id": "0-标题",
  "left": 5, "top": 6, "width": 65, "height": 8,
  "text_content": [
    {
      "alignment": "left",
      "runs": [{"text": "Hello", "font_size": 360000, "bold": true, "color": "#333333"}]
    }
  ]
}
使用模板 placeholder:
{
  "type": "textbox",
  "id": "0-标题",
  "placeholder": "0",
  "text_content": [
    {
      "alignment": "center",
      "runs": [{"text": "标题文字", "font_size": 540000}]
    }
  ]
}
字段:
- placeholder: "0"通常为标题栏, "1"通常为正文框，具体以 list_layouts() 返回的 idx 为准。设了 placeholder 后 left/top/width/height 可省略（从版式继承）也可手写（覆盖版式位置）
- text_content: array of paragraph. 每段有 alignment + runs[]
- runs[].font_size: EMU 单位 (1pt=12700, 18pt=228600, 36pt=457200)
- runs[].bold/italic/underline: bool
- runs[].font_name: 可选，字体名称，如 "Microsoft YaHei"
- runs[].color: **除非必要不要设**，背景未知设了可能看不清。null 时自动使用主题色 dk1（黑色）
- （当前不支持 hyperlink，不要使用）

### 2. image — 图片
{
  "type": "image",
  "id": "0-图片",
  "left": 4, "top": 20, "width": 25, "height": 30,
  "image_content": {"path": "photo.png"}
}
字段:
- image_content.path: 图片文件名（必须已上传到图片素材库）
- image_content.data: 图片 base64 编码数据（字符串），二选一
- image_content.url: 图片可下载 URL，二选一
- 图片**不支持** placeholder。使用模板时请放在两栏内容或空白版式上，坐标必须手写

### 3. table — 表格
{
  "type": "table",
  "id": "0-表格",
  "left": 4, "top": 40, "width": 65, "height": 40,
  "table_content": {
    "rows": 2, "cols": 3,
    "cells": [[{"text":"A1"},{"text":"B1"},{"text":"C1"}],[{"text":"A2"},{"text":"B2"},{"text":"C2"}]]
  }
}
字段:
- cells: 二维数组, cells[行][列]，每项为 CellContent 对象：
  {"text":str,"bold":bool,"alignment":"left"|"center"|"right","background_color":"#RRGGBB","colspan":1,"rowspan":1,"font_size":float}
- 使用模板时表格支持 placeholder（继承版式位置），也支持手写坐标覆盖位置

### 4. shape — 图形（形状）
{
  "type": "shape",
  "id": "0-背景",
  "left": 0, "top": 0, "width": 100, "height": 100,
  "auto_shape_type": "rectangle",
  "background_color": "#E8E8E8",
  "text_content": [{"runs": [{"text": "带形状的文字", "font_size": 240000}]}]
}
字段:
- auto_shape_type: 形状类型（见下方支持列表）
- background_color: hex "#RRGGBB"，形状背景色
- 可包含 text_content 在里面写字

支持的 auto_shape_type 值:
- 基础: rectangle, rounded_rectangle, oval, diamond
- 多边形: parallelogram, trapezoid, pentagon, regular_pentagon, hexagon, heptagon, octagon, decagon, dodecagon
- 箭头: right_arrow, left_arrow, up_arrow, down_arrow, chevron, bent_arrow, bent_up_arrow, circular_arrow, curved_right_arrow, curved_left_arrow, curved_up_arrow, curved_down_arrow, striped_right_arrow, notched_right_arrow, left_right_arrow, up_down_arrow, quad_arrow
- 括号/花括号: right_bracket, left_bracket, right_brace, left_brace, double_bracket, double_brace
- 星形: star_5_point, star_6_point, star_7_point, star_10_point, star_12_point, star_16_point, star_24_point, star_32_point
- 特殊: cross, heart, cloud, sun, moon, lightning_bolt, no_symbol
- 饼/环: pie, pie_wedge, block_arc, donut
- 3D/立体: bevel, cube, can, folded_corner
- 卷角/标签: corner_tabs, plaque_tabs, horizontal_scroll, vertical_scroll
- 波形/气泡: wave, double_wave, tear, balloon, funnel
- 齿轮: gear_6, gear_9
- 流程图: flowchart_process, flowchart_decision, flowchart_document, flowchart_terminator, flowchart_data
- 数学符号: math_plus, math_minus, math_multiply, math_divide, math_equal, math_not_equal
- 动作按钮: action_button_custom

### 5. container — 容器（组合元素）

**用途**：当同一个版式占位符区域需要放置多个元素时（如背景色块 + 叠在上层的文字），
用容器将它们包在一起。容器自己绑定 placeholder，子元素**不能**有 placeholder。

{
  "type": "textbox",
  "placeholder": "0",
  "id": "0-title-area",
  "children": [
    {
      "type": "shape",
      "auto_shape_type": "rounded_rectangle",
      "background_color": "#4472C4",
      "left": 0, "top": 0, "width": 100, "height": 100
    },
    {
      "type": "textbox",
      "text_content": [{"runs": [{"text": "Title", "font_size": 360000}]}],
      "left": 3, "top": 10, "width": 95, "height": 40
    },
    {
      "type": "textbox",
      "text_content": [{"runs": [{"text": "Subtitle", "font_size": 180000}]}],
      "left": 3, "top": 55, "width": 95, "height": 30
    }
  ]
}
字段:
- children: 子元素数组，坐标相对于容器的 left/top
- placeholder: 容器的占位符。同一页内该值只能出现一次
- 子元素**不能**有 placeholder，不支持嵌套 children
- 容器的 type 建议用 "textbox"（有 children 时不渲染自身），但仍须是有效值

### 通用字段（所有形状共有）
- type: 必填, "textbox"|"image"|"table"|"shape". 永远不要用 "text"
- id: 推荐格式 "{页码}-{语义}", 如 "0-标题", "1-正文". 服务器自动修正前缀
- left/top/width/height: 百分比 (0-100)，相对于幻灯片或父容器尺寸。引擎自动换算为 EMU
- rotation: 可选, 旋转角度 0-360. 形状和文字一起旋转
- z_order: 可选 int, 默认 0. 绘制层级（值越大越靠前，用于控制形状前后顺序）
- border_color: 可选 "#RRGGBB"，边框颜色
- border_width: 可选 float (EMU)，边框宽度
- border_style: 可选, "solid"(默认)|"dashed"|"dotted"
- shadow: 可选 bool, 默认 false. 是否添加阴影
"""


# ============================================================
#  HTTP helpers
# ============================================================


def _check(r: httpx.Response):
    if r.is_error:
        raise RuntimeError(
            f"HTTP {r.status_code} from {r.request.method} {r.url}\n"
            f"Body: {r.text[:2000]}"
        )


def _get(path: str) -> str:
    logger.info("→ GET %s", path)
    r = httpx.get(f"{BASE_URL}{path}", timeout=30)
    _check(r)
    result = json.dumps(r.json(), ensure_ascii=False, indent=2)
    logger.info("← GET %s  200  (%d chars)", path, len(result))
    return result


def _post(path: str, body: dict | None = None) -> dict:
    logger.info("→ POST %s", path)
    r = httpx.post(f"{BASE_URL}{path}", json=body, timeout=60)
    _check(r)
    result = r.json()
    logger.info("← POST %s  %s", path, result.get("status", "ok"))
    return result


def _put(path: str, body: dict) -> dict:
    logger.info("→ PUT %s", path)
    r = httpx.put(f"{BASE_URL}{path}", json=body, timeout=30)
    _check(r)
    result = r.json()
    logger.info("← PUT %s  %s", path, result.get("status", "ok"))
    return result


def _delete(path: str) -> dict:
    logger.info("→ DELETE %s", path)
    r = httpx.delete(f"{BASE_URL}{path}", timeout=30)
    _check(r)
    result = r.json()
    logger.info("← DELETE %s  %s", path, result.get("status", "ok"))
    return result


# ============================================================
#  Shared tool functions (used by both modes)
# ============================================================


def tool_list_templates() -> str:
    """List all loaded template names. Lightweight — call this first."""
    data = _get("/agent/templates")
    try:
        names = list(json.loads(data).get("templates", {}).keys())
        return json.dumps({"templates": names}, ensure_ascii=False, indent=2)
    except Exception:
        return data


def tool_list_layouts(template_name: str) -> str:
    """List layouts and their placeholders for a specific template.

    Call this AFTER the user picks a template. Returns each layout's
    placeholders with idx, name, type, left, top, width, height.
    Use idx as the "placeholder" value in shape JSON.
    """
    return _get(f"/template/{template_name}/layouts")


def tool_list_schemes() -> str:
    """List available color scheme names. Only returns names — call
    get_color_scheme(name) to see the actual hex palette."""
    return _get("/agent/schemes")


def tool_get_color_scheme(name: str) -> str:
    """Show a color scheme's full palette with all 12 color hex values.
    Call AFTER list_schemes() or when you need to pick colors that
    fit the chosen visual theme."""
    return _get(f"/agent/schemes/{name}")


def tool_list_images() -> str:
    """List all uploaded images with filename, size, mtime."""
    return _get("/agent/images")


def tool_upload_image(filename: str, data: str) -> str:
    """Upload an image to the image library.

    filename: file name with extension, e.g. "diagram.png", "photo.jpg"
    data: base64-encoded image data (raw bytes, no JSON wrapper)
    """
    result = _post(f"/agent/images/upload?filename={filename}", {"data": data})
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_upload_image_from_file(filename: str, file_path: str) -> str:
    """Upload a local image file to the image library.

    filename: target name in the library, e.g. "photo.png", "diagram.jpg"
    file_path: path to the local image file on disk
    """
    logger.info("upload_image_from_file filename=%s path=%s", filename, file_path)
    path = os.path.abspath(file_path)
    if not os.path.isfile(path):
        raise RuntimeError(f"File not found: {path}")
    ext = filename.rsplit(".", 1)[-1] if "." in filename else "png"
    with open(path, "rb") as f:
        store_r = httpx.post(
            f"{BASE_URL}/image/store",
            files={"file": (filename, f, f"image/{ext}")},
            timeout=60,
        )
    _check(store_r)
    result = store_r.json()
    logger.info("← upload_image_from_file stored as %s", result.get("filename"))
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_upload_image_from_url(filename: str, url: str) -> str:
    """Download an image from a URL and upload it to the image library.

    filename: file name with extension, e.g. "photo.png", "diagram.jpg"
    url: publicly accessible image URL to download
    """
    logger.info("upload_image_from_url filename=%s url=%s", filename, url)
    r = httpx.get(url, timeout=60, follow_redirects=True)
    r.raise_for_status()
    store_r = httpx.post(
        f"{BASE_URL}/image/store",
        files={"file": (filename, r.content, "image/" + (filename.rsplit(".", 1)[-1] if "." in filename else "png"))},
        timeout=60,
    )
    _check(store_r)
    result = store_r.json()
    logger.info("← upload_image_from_url stored as %s", result.get("filename"))
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_get_example() -> str:
    return _get("/agent/example")


def tool_list_design_styles() -> str:
    return _get("/agent/design-styles/list")


def tool_get_design_style(name: str) -> str:
    return _get(f"/agent/design-styles/{name}")


def tool_download_design_style(name: str) -> str:
    """Download a design style as Markdown text."""
    logger.info("→ GET /agent/design-styles/%s/download", name)
    r = httpx.get(f"{BASE_URL}/agent/design-styles/{name}/download", timeout=30)
    if r.is_error:
        raise RuntimeError(f"HTTP {r.status_code}: {r.text[:2000]}")
    logger.info("← GET /agent/design-styles/%s/download  200  (%d chars)", name, len(r.text))
    return r.text


def tool_save_design_style(name: str, description: str, markdown: str) -> str:
    """Create or update a design style with name, description and Markdown content."""
    r = _post(f"/agent/design-styles/{name}", {"name": name, "description": description, "markdown": markdown})
    return json.dumps(r, ensure_ascii=False, indent=2)


def tool_delete_design_style(name: str) -> str:
    """Delete a named design style."""
    return json.dumps(_delete(f"/agent/design-styles/{name}"), ensure_ascii=False, indent=2)


def tool_sync_presentation(file_path: str) -> str:
    """Upload a local JSON file to the sync area. Reads the file from disk.

    file_path: absolute or relative path to a .json file on the local machine
    """
    logger.info("sync_presentation file_path=%s", file_path)
    path = os.path.abspath(file_path)
    if not os.path.isfile(path):
        raise RuntimeError(f"File not found: {path}")
    with open(path, encoding="utf-8") as f:
        body = json.load(f)
    if "slides" not in body:
        raise RuntimeError("JSON must contain a 'slides' array")
    result = _post("/agent/sync", body)
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_export_pptx(
    output_path: str,
    template_name: str = "",
    color_scheme: str = "",
) -> str:
    """Build PPTX from sync file.

    Design style must be chosen first, which gives you the template and color
    scheme. Extract them from the design style's "## 模板" and "## 配色"
    sections and pass them here. Example: template_name="标准1", color_scheme="深邃蓝".
    """
    if not template_name or not color_scheme:
        raise RuntimeError("template_name and color_scheme are required — extract them from the design style")
    logger.info("export_pptx output=%s template=%s scheme=%s", output_path, template_name, color_scheme)
    params = f"output_path={output_path}&template_name={template_name}&color_scheme={color_scheme}"
    result = _post(f"/agent/export?{params}")
    # Clear sync area after successful export
    try:
        _post("/agent/reset")
    except Exception:
        pass
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_shape_reference() -> str:
    """Return the complete shape type catalog with field descriptions.

    Call this to learn what shapes are available and their exact JSON format.
    """
    return _SHAPE_CATALOG


# ============================================================
#  Remote-mode tool functions
# ============================================================


def tool_read_all() -> str:
    return _get("/sync")


def tool_read_page(page: int) -> str:
    return _get(f"/agent/page/{page}")


def tool_read_element(page: int, id: str) -> str:
    return _get(f"/agent/element?page={page}&id={id}")


def tool_list_ids() -> str:
    return _get("/agent/ids")


def tool_reset_presentation() -> str:
    result = _post("/agent/reset")
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_write_page(page: int, json_str: str) -> str:
    body = json.loads(json_str)
    result = _put(f"/agent/page/{page}", body)
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_write_element(page: int, id: str, json_str: str) -> str:
    body = json.loads(json_str)
    result = _put(f"/agent/element?page={page}&id={id}", body)
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_delete_page(page: int) -> str:
    result = _delete(f"/agent/page/{page}")
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_insert_page(position: int) -> str:
    result = _post(f"/agent/page?position={position}")
    return json.dumps(result, ensure_ascii=False, indent=2)


# ============================================================
#  FastMCP — LOCAL mode
# ============================================================

mcp_local = FastMCP(
    "PPT Render Engine (local mode)",
    instructions=f"""## 本地文件模式

你有本地文件访问权限，在 JSON 文件里编辑内容，MCP 只负责同步和导出。

### 工作流（严格按此顺序执行，不可跳过）

**第 1 步 — 选设计风格**
必须先问用户选哪个设计风格。调用 `list_design_styles()` 展示列表，
等用户选择后调用 `get_design_style(name)` 获取规范。
**在完成此步之前，禁止创建或编辑任何内容。**

**第 2 步 — 熟悉格式、模板布局与配色**
- `shape_reference()` 查看形状类型和 JSON 格式
- 从设计规范中提取模板名（如"标准1"）和配色名（如"深邃蓝"），调用 `list_layouts(name)` 查看各版式占位符的 idx、坐标、类型，调用 `get_color_scheme(name)` 查看 12 个颜色槽的色值
- **必须理解每个版式有哪些占位符（如 idx=0 标题、idx=1 正文），以及配色板有哪些可用色值，然后设计内容时给每个形状配上正确的 placeholder 和色板颜色**
- **注意：`list_layouts()` 列出的占位符中，索引 10/11/12 是日期/页脚/编号等系统槽位，不能用于放内容。形状的 placeholder 只能使用内容占位符（如 0/1/2）**
- **在完成本步之前，禁止写任何 JSON。**

**第 3 步 — 编内容**
在本地编辑 JSON 文件，严格按照设计规范的配色、字号、版式来写。
每页至少用 2 种以上视觉元素（色块/表格/箭头/流程图等），禁止纯文字。
**内容未完成时，禁止跳到第 4 步。**

**第 4 步 — 同步**
`sync_presentation(file_path)` 上传到同步区。
**返回的 warnings 中若含"文本溢出"，建议先修复（增大形状/减小字号/缩短文字），引擎会自动缩小但严重溢出仍可能影响效果。**
**只有 warnings 中没有溢出问题了，才能到第 5 步。**

**第 5 步 — 导出**
`export_pptx(path, template, scheme)` 生成 PPTX。
- 设计规范中有"## 模板"和"## 配色"章节，**必须**从中提取对应的模板名和配色名传入
- 例如规范写"使用标准1模板"和"深邃蓝配色"，就传 `export_pptx(path, template_name="标准1", color_scheme="深邃蓝")`
**导出后提醒用户可预览（{_UI_URL}），告知所用的设计风格/模板/配色，询问是否保存为新的设计风格。**

### 工具

| Tool | 用途 |
|------|------|
| `list_design_styles()` | 设计风格列表，**生成内容前先问用户选哪个** |
| `get_design_style(name)` | 查看指定风格的 Markdown 详情 |
| `download_design_style(name)` | 下载 Markdown 文件 |
| `save_design_style(name, description, markdown)` | 创建或更新风格（需自己写 md） |
| `delete_design_style(name)` | 删除风格 |
| `get_example()` | 示例演示文稿 JSON |
| `shape_reference()` | 形状类型和字段全集参考 |
| `sync_presentation(file_path)` | 上传本地 JSON 文件到同步区（传文件路径，不要传内容） |
| `export_pptx(path, template?, scheme?)` | 生成 PPTX（导出后自动清空同步区） |
| `list_templates()` | 模板名称列表（轻量，先调） |
| `list_layouts(name)` | 指定模板的版式和占位符详情（选后调） |
| `list_schemes()` | 配色方案 |
| `get_color_scheme(name)` | 查看指定配色的全部色值（12 个颜色槽） |
| `list_images()` | 已上传图片素材 |
| `upload_image_from_file(filename, path)` | 上传本地图片文件到素材库 |
| `upload_image_from_url(filename, url)` | 从 URL 下载图片到素材库 |

### 图片使用
- 用户提到图片时，先 `list_images()` 检查素材库已有图片
- 用户给了 URL：`upload_image_from_url(filename, url)` 下载并入库
- 用户给了本地文件：`upload_image_from_file(filename, path)` 上传
- 图片文件名会暴露在 image_content.path 中，请保持简洁

### 使用模板时的规则

#### 占位符与容器
- **顶层形状**（slide.shapes 里的元素）：**必须**有 placeholder
- **同一页内**每个 placeholder 值只能出现一次
- 如果一个占位符区域需要放多个元素（如色块 + 文字），用容器包起来：容器绑定 placeholder，子元素（children）**不能**有 placeholder
- **图片**：不支持 placeholder，手写坐标；模板模式下请放在两栏内容或空白版式上
- **表格**：支持 placeholder，也可手写坐标

#### 坐标
- 有 placeholder 的顶层形状：坐标从版式继承，不要手写 left/top/width/height
- 容器（children）：容器自身坐标从版式继承（不手写），但**子元素的坐标必须手写**（百分比 0-100，相对于容器）
- 图片 / 表格（手写坐标）：用百分比 (0-100)

#### 颜色
- 字体颜色：**不要设**，主题色自动处理
- 形状背景色 background_color：**必须从已选配色方案的色板中取**。用 `get_color_scheme(name)` 查看色值，选 12 个颜色槽中的颜色
- 使用设计风格时，优先按设计规范里的配色指引选择对应色板颜色
- template_name 和 color_scheme 在 export_pptx 时传，不在元素 JSON 里设

### 通用约束

- type 只能是 textbox / image / table / shape，不要用 text
- id 格式 "{{页码}}-{{语义}}", 如 "0-标题"
- 坐标用百分比 (0-100)；字号单位是 EMU（1pt=12700，14pt=177800，18pt=228600）
- **文字溢出**：引擎自动缩小文字适应形状，但严重溢出仍会警告，建议保持文字简洁
- **元素遮挡**：warnings 中若含"重叠"，必须调整位置或缩小尺寸消除遮挡后再导出
- **优先使用丰富的视觉元素**：每页尽量结合表格、色块卡片、箭头、菱形、流程图等多种形状，避免纯文字堆砌
- **逻辑关系用图形表达**：流程用箭头串联，分支用菱形+矩形，对比用表格，层级用堆叠矩形，数据用图表布局
- 导出后提醒用户可去 {_UI_URL} 预览
- 导出后告知用户当前使用的**设计风格**、**模板**和**配色方案**，并询问是否保存为新的设计风格（`save_design_style`)
""",
)

# ============================================================
#  FastMCP — REMOTE mode
# ============================================================

_instructions_remote = f"""## 远程模式

无法访问本地文件，需要用工具逐页或逐个形状构建演示文稿。

### 工作流（严格按此顺序执行，不可跳过）

**第 1 步 — 选设计风格**
必须先问用户选哪个设计风格。调用 `list_design_styles()` 展示列表，
等用户选择后调用 `get_design_style(name)` 获取规范。
**在完成此步之前，禁止创建或编辑任何内容。**

**第 2 步 — 熟悉格式、模板布局与配色**
- `shape_reference()` 查看形状类型和 JSON 格式
- 从设计规范中提取模板名（如"标准2"）和配色名（如"森林绿"），调用 `list_layouts(name)` 查看各版式占位符的 idx、坐标、类型，调用 `get_color_scheme(name)` 查看 12 个颜色槽的色值
- **必须理解每个版式有哪些占位符（如 idx=0 标题、idx=1 正文），以及配色板有哪些可用色值，然后设计内容时给每个形状配上正确的 placeholder 和色板颜色**
- **注意：`list_layouts()` 列出的占位符中，索引 10/11/12 是日期/页脚/编号等系统槽位，不能用于放内容。形状的 placeholder 只能使用内容占位符（如 0/1/2）**
- **在完成本步之前，禁止写任何 JSON。**

**第 3 步 — 逐页构建内容**
用 `write_page(page, json)` 或 `write_element(page, id, json)` 逐页写内容，
严格按照设计规范的配色、字号、版式来写。每页至少用 2 种以上视觉元素。
**内容未完成时，禁止跳到第 4 步。**

**第 4 步 — 导出**
`export_pptx(path, template, scheme)` 生成 PPTX。
- 设计规范中有"## 模板"和"## 配色"章节，**必须**从中提取对应的模板名和配色名传入
- 例如规范写"使用标准2模板"和"森林绿配色"，就传 `export_pptx(path, template_name="标准2", color_scheme="森林绿")`
**导出前确保所有页面已 build，且 write_page 返回的 warnings 中没有严重问题。**
**导出后提醒用户可预览（{_UI_URL}），告知所用的设计风格/模板/配色，询问是否保存为新的设计风格。**

### 工具

| Tool | 用途 |
|------|------|
| `list_design_styles()` | 设计风格列表，**生成内容前先问用户选哪个** |
| `get_design_style(name)` | 查看指定风格的 Markdown 详情 |
| `download_design_style(name)` | 下载 Markdown 文件 |
| `save_design_style(name, description, markdown)` | 创建或更新风格（需自己写 md） |
| `delete_design_style(name)` | 删除风格 |
| `get_example()` | 示例演示文稿 JSON |
| `shape_reference()` | 形状类型和字段全集参考 |
| `reset_presentation()` | 清空同步区，从空白开始 |
| `read_all()` | 读取整个同步文件 |
| `read_page(page)` | 读取某页 |
| `read_element(page, id)` | 读取某个形状 |
| `list_ids()` | 列出所有形状 ID |
| `list_templates()` | 模板名称列表（轻量，先调） |
| `list_layouts(name)` | 指定模板的版式和占位符详情（选后调） |
| `list_schemes()` | 配色方案 |
| `get_color_scheme(name)` | 查看指定配色的全部色值（12 个颜色槽） |
| `list_images()` | 已上传图片素材 |
| `upload_image_from_file(filename, path)` | 上传本地图片文件到素材库 |
| `upload_image_from_url(filename, url)` | 从 URL 下载图片到素材库 |
| `write_page(page, json)` | 合并更新一整页（未指定字段保留原值；page=总页数时追加新页） |
| `write_element(page, id, json)` | 修改一个形状（不存在则追加） |
| `insert_page(position)` | 在指定位置插入空白页 |
| `delete_page(page)` | 删除一页 |
| `sync_presentation(file_path)` | 上传本地 JSON 文件到同步区（传文件路径） |
| `export_pptx(path, template?, scheme?)` | 生成 PPTX（导出后自动清空同步区） |

### 图片使用
- 用户提到图片时，先 `list_images()` 检查素材库已有图片
- 用户给了 URL：`upload_image_from_url(filename, url)` 下载并入库
- 用户给了本地文件：`upload_image_from_file(filename, path)` 上传

### 使用模板时的规则

#### 占位符与容器
- **顶层形状**（slide.shapes 里的元素）：**必须**有 placeholder
- **同一页内**每个 placeholder 值只能出现一次
- 如果一个占位符区域需要放多个元素（如色块 + 文字），用容器包起来：容器绑定 placeholder，子元素（children）**不能**有 placeholder
- **图片**：不支持 placeholder，手写坐标；模板模式下请放在两栏内容或空白版式上
- **表格**：支持 placeholder，也可手写坐标

#### 坐标
- 有 placeholder 的顶层形状：坐标从版式继承，不要手写 left/top/width/height
- 容器（children）：容器自身坐标从版式继承（不手写），但**子元素的坐标必须手写**（百分比 0-100，相对于容器）
- 图片 / 表格（手写坐标）：用百分比 (0-100)

#### 颜色
- 字体颜色：**不要设**，主题色自动处理
- 形状背景色 background_color：**必须从已选配色方案的色板中取**。用 `get_color_scheme(name)` 查看色值，选 12 个颜色槽中的颜色
- 使用设计风格时，优先按设计规范里的配色指引选择对应色板颜色
- template_name 和 color_scheme 在 export_pptx 时传，不在元素 JSON 里设

### 通用约束

- type 只能是 textbox / image / table / shape，不要用 text
- id 格式 "{{页码}}-{{语义}}"
- 坐标用百分比 (0-100)；字号单位是 EMU（1pt=12700，14pt=177800，18pt=228600）
- **文字溢出**：引擎自动缩小文字适应形状，但严重溢出仍会警告，建议保持文字简洁
- **元素遮挡**：warnings 中若含"重叠"，必须调整位置或缩小尺寸消除遮挡后再导出
- **优先使用丰富的视觉元素**：每页尽量结合表格、色块卡片、箭头、菱形、流程图等多种形状，避免纯文字堆砌
- **逻辑关系用图形表达**：流程用箭头串联，分支用菱形+矩形，对比用表格，层级用堆叠矩形，数据用图表布局
- write_page(N) 当 N 等于总页数时自动追加新页
- 插入或删除页后所有 ID 自动重编号
- 导出后提醒用户可去 {_UI_URL} 预览
- 导出后告知用户当前使用的**设计风格**、**模板**和**配色方案**，并询问是否保存为新的设计风格（`save_design_style`)
"""

mcp_remote = FastMCP(
    "PPT Render Engine (remote mode)",
    instructions=_instructions_remote,
)


# ============================================================
#  Register tools on each instance
# ============================================================

def _register_shared(mcp_obj):
    mcp_obj.tool(description="获取完整形状类型参考（4 种形状 + 所有字段说明）。构建 PPT 前优先调用")(tool_shape_reference)
    mcp_obj.tool(description="获取完整示例演示文稿 JSON（含模板引用、placeholder）。优先调用此工具学习格式")(tool_get_example)
    mcp_obj.tool(description="列出所有可用的设计风格（含描述和字符数），选择后可用 get_style 或 download")(tool_list_design_styles)
    mcp_obj.tool(description="查看指定设计风格的完整信息（含 Markdown 内容）")(tool_get_design_style)
    mcp_obj.tool(description="下载设计风格的 Markdown 文件")(tool_download_design_style)
    mcp_obj.tool(description="创建或更新设计风格（Markdown 格式），提供 name、description 和 markdown 内容")(tool_save_design_style)
    mcp_obj.tool(description="删除一个已保存的设计风格")(tool_delete_design_style)
    mcp_obj.tool(description="上传完整 JSON 覆盖同步文件（自动修正 ID 前缀）")(tool_sync_presentation)
    mcp_obj.tool(description="从同步文件生成 PPTX。模板和配色从设计规范中提取并传入")(tool_export_pptx)
    mcp_obj.tool(description="列出所有已加载的模板名称（轻量，先调此工具让用户选）")(tool_list_templates)
    mcp_obj.tool(description="查看指定模板的版式及占位符详情（选定模板后调用）")(tool_list_layouts)
    mcp_obj.tool(description="列出所有可用配色方案")(tool_list_schemes)
    mcp_obj.tool(description="查看指定配色的全部色值（12 个颜色槽）")(tool_get_color_scheme)
    mcp_obj.tool(description="列出已上传的图片素材")(tool_list_images)
    mcp_obj.tool(description="上传本地图片文件到素材库，提供文件名和文件路径")(tool_upload_image_from_file)
    mcp_obj.tool(description="从 URL 下载图片并存入素材库")(tool_upload_image_from_url)


def _register_remote(mcp_obj):
    mcp_obj.tool(description="读取整个同步文件内容")(tool_read_all)
    mcp_obj.tool(description="读取指定页 JSON")(tool_read_page)
    mcp_obj.tool(description="按语义 ID 读取形状")(tool_read_element)
    mcp_obj.tool(description="列出所有形状 ID")(tool_list_ids)
    mcp_obj.tool(description="清空同步区，创建空白演示文稿")(tool_reset_presentation)
    mcp_obj.tool(description="合并更新一整页（body 字段覆盖现有，未指定字段保留原值）；page=总页数时追加新页")(tool_write_page)
    mcp_obj.tool(description="按 ID 修改形状（未找到则追加）")(tool_write_element)
    mcp_obj.tool(description="删除指定页（自动重编号 ID）")(tool_delete_page)
    mcp_obj.tool(description="在指定位置插入空白页（自动重编号 ID）")(tool_insert_page)


_register_shared(mcp_local)
_register_shared(mcp_remote)
_register_remote(mcp_remote)


# ============================================================
#  Fuzzy matching on tool name — applied to whichever MCP is active
# ============================================================

def _install_fuzzy(mcp_obj):
    _orig = mcp_obj._tool_manager.get_tool
    def _fuzzy(name: str):
        t = _orig(name)
        if t:
            return t
        candidates = mcp_obj._tool_manager.list_tools()
        names = [x.name for x in candidates]
        matches = difflib.get_close_matches(name, names, n=3, cutoff=0.4)
        if matches:
            raise ToolError(
                f"Unknown tool: {name!r}. Did you mean: {', '.join(matches)}?"
            )
        raise ToolError(
            f"Unknown tool: {name!r}. Available: {', '.join(sorted(names))}"
        )
    mcp_obj._tool_manager.get_tool = _fuzzy


_install_fuzzy(mcp_local)
_install_fuzzy(mcp_remote)


# ============================================================
#  Entrypoint
# ============================================================


def main(mode: str = "remote"):
    import uvicorn
    if mode == "local":
        app = mcp_local.sse_app()
    else:
        app = mcp_remote.sse_app()
    uvicorn.run(app, host=MCP_HOST, port=MCP_PORT)
