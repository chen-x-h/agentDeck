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
from pathlib import Path
from typing import Optional
import httpx
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

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
_UI_URL = f"http://127.0.0.1:{8000}"


# ============================================================
#  shape reference catalog
# ============================================================

_SHAPE_CATALOG = """## 形状类型参考 (Shape Reference)

所有形状必须用标准嵌套 JSON 格式，不可用扁平字段。

### 1. textbox — 文本框
无模板（手写坐标）:
{
  "type": "textbox",
  "id": "0-标题",
  "left": 500000, "top": 400000, "width": 8000000, "height": 500000,
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
- placeholder: "0"=标题栏, "1"=正文框. 设了 placeholder 就不要设 left/top/width/height
- text_content: array of paragraph. 每段有 alignment + runs[]
- runs[].font_size: EMU 单位 (1pt=12700, 18pt=228600, 36pt=457200)
- runs[].bold/italic/underline: bool
- runs[].color: **除非必要不要设**，背景未知设了可能看不清。用 null 让主题色处理
- runs[].hyperlink: {"url":"https://..."} 或 null

### 2. image — 图片
{
  "type": "image",
  "id": "0-图片",
  "left": 500000, "top": 1500000, "width": 3000000, "height": 2000000,
  "image_content": {"path": "photo.png"}
}
字段:
- image_content.path: 图片文件名（必须已上传到图片素材库）
- 坐标/尺寸 必须手写——图片没有 placeholder

### 3. table — 表格
{
  "type": "table",
  "id": "0-表格",
  "left": 500000, "top": 3000000, "width": 8000000, "height": 3000000,
  "table_content": {
    "rows": 2, "cols": 3,
    "cells": [["A1","B1","C1"],["A2","B2","C2"]]
  }
}
字段:
- cells: 二维数组, cells[行][列]
- 可以设 placeholder 继承位置, 但表格通常手写坐标

### 4. shape — 图形（形状）
{
  "type": "shape",
  "id": "0-背景",
  "left": 0, "top": 0, "width": 12192000, "height": 6858000,
  "preset": "rect",
  "fill_color": "#E8E8E8",
  "text_content": [{"runs": [{"text": "带形状的文字", "font_size": 240000}]}]
}
字段:
- preset: 形状类型（rect=矩形, roundrect=圆角矩形, oval=椭圆, 等）
- fill_color: hex "#RRGGBB"
- 可包含 text_content 在里面写字

### 通用字段（所有形状共有）
- type: 必填, "textbox"|"image"|"table"|"shape". 永远不要用 "text"
- id: 推荐格式 "{页码}-{语义}", 如 "0-标题", "1-正文". 服务器自动修正前缀
- left/top/width/height: EMU 单位. 整页: 12192000 x 6858000
- rotation: 可选, 旋转角度 0-360
- hidden: 可选 bool, 隐藏形状
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
    r = httpx.get(f"{BASE_URL}{path}", timeout=30)
    _check(r)
    return json.dumps(r.json(), ensure_ascii=False, indent=2)


def _post(path: str, body: dict | None = None) -> dict:
    r = httpx.post(f"{BASE_URL}{path}", json=body, timeout=60)
    _check(r)
    return r.json()


def _put(path: str, body: dict) -> dict:
    r = httpx.put(f"{BASE_URL}{path}", json=body, timeout=30)
    _check(r)
    return r.json()


def _delete(path: str) -> dict:
    r = httpx.delete(f"{BASE_URL}{path}", timeout=30)
    _check(r)
    return r.json()


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
    return _get("/agent/schemes")


def tool_list_images() -> str:
    return _get("/agent/images")


def tool_get_example() -> str:
    return _get("/agent/example")


def tool_sync_presentation(json_str: str) -> str:
    """Upload full JSON to sync file. Use in both modes to apply changes.

    json_str: complete presentation JSON with "slides" array.
    """
    body = json.loads(json_str)
    result = _post("/agent/sync", body)
    return json.dumps(result, ensure_ascii=False, indent=2)


def tool_export_pptx(
    output_path: str,
    template_name: Optional[str] = None,
    color_scheme: Optional[str] = None,
) -> str:
    """Build PPTX from sync file.

    ALWAYS call list_templates() + list_schemes() first and ask user
    which template/color scheme to use. Pass them here — do NOT put
    them in shape JSON.
    """
    params = f"output_path={output_path}"
    if template_name:
        params += f"&template_name={template_name}"
    if color_scheme:
        params += f"&color_scheme={color_scheme}"
    return json.dumps(_post(f"/agent/export?{params}"), ensure_ascii=False, indent=2)


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

### 工作流

1. `get_example()` 熟悉格式，`shape_reference()` 看完整的形状字段说明
2. 在本地编辑 JSON 文件（项目目录内，不要放临时目录）
3. `sync_presentation(content)` 上传到同步区
4. `list_templates()` 查看模板和占位符信息，让用户选一个
5. `export_pptx(path, template, scheme)` 生成 PPTX

### 工具

| Tool | 用途 |
|------|------|
| `get_example()` | 示例 JSON，优先调用 |
| `shape_reference()` | 形状类型和字段全集参考 |
| `sync_presentation(json_str)` | 上传本地 JSON 到同步区 |
| `export_pptx(path, template?, scheme?)` | 生成 PPTX |
| `list_templates()` | 模板名称列表（轻量，先调） |
| `list_layouts(name)` | 指定模板的版式和占位符详情（选后调） |
| `list_schemes()` | 配色方案 |
| `list_images()` | 已上传图片素材 |

### 使用模板时的规则

先调 `list_templates()` 查看可用模板名称让用户选，选后再调 `list_layouts(name)` 确认每个版式有哪些占位符——返回值里每个 placeholder 有 idx、name、位置和大小。用 idx 作为你形状的 `placeholder` 值。

模板的每一页都是从某个版式创建的。所有元素必须有 `placeholder` 标明它属于版式哪个槽位；坐标可省（从版式继承位置大小）也可手写（覆盖版式位置）。多个形状用坐标定位时注意间距，避免互相遮挡。不要设字体颜色——主题色自动处理。`template_name` 和 `color_scheme` 在 `export_pptx` 时传。

### 通用约束

- type 只能是 textbox / image / table / shape，不要用 text
- id 格式 "{{页码}}-{{语义}}"，如 "0-标题"
- 坐标和字号单位是 EMU（1pt=12700，14pt=177800，18pt=228600）
- 文字超出框会被裁剪，控制长度
- 导出后提醒用户可去 {_UI_URL} 预览
""",
)


# ============================================================
#  FastMCP — REMOTE mode
# ============================================================

_instructions_remote = f"""## 远程模式

无法访问本地文件，需要用工具逐页或逐个形状构建演示文稿。

### 工作流

1. `get_example()` 熟悉格式，`shape_reference()` 看完整字段说明
2. 需要模板的话先 `list_templates()` 看名称让用户选，再 `list_layouts(name)` 看占位符详情
3. 用 `write_page(page, json)` 或 `write_element(page, id, json)` 逐页构建
4. `export_pptx(path, template, scheme)` 生成 PPTX

### 工具

| Tool | 用途 |
|------|------|
| `get_example()` | 示例 JSON，优先调用 |
| `shape_reference()` | 形状类型和字段全集参考 |
| `reset_presentation()` | 清空同步区，从空白开始 |
| `read_all()` | 读取整个同步文件 |
| `read_page(page)` | 读取某页 |
| `read_element(page, id)` | 读取某个形状 |
| `list_ids()` | 列出所有形状 ID |
| `list_templates()` | 模板名称列表（轻量，先调） |
| `list_layouts(name)` | 指定模板的版式和占位符详情（选后调） |
| `list_schemes()` | 配色方案 |
| `list_images()` | 已上传图片素材 |
| `write_page(page, json)` | 替换一整页（page=总页数时追加） |
| `write_element(page, id, json)` | 修改一个形状（不存在则追加） |
| `insert_page(position)` | 在指定位置插入空白页 |
| `delete_page(page)` | 删除一页 |
| `sync_presentation(json_str)` | 批量覆盖全部内容 |
| `export_pptx(path, template?, scheme?)` | 生成 PPTX |

### 使用模板时的规则

先调 `list_templates()` 查看可用模板名称让用户选，选后再调 `list_layouts(name)` 确认每个版式有哪些占位符——返回值里每个 placeholder 有 idx、name、位置和大小。用 idx 作为你形状的 `placeholder` 值。

模板的每一页都是从某个版式创建的。所有元素必须有 `placeholder` 标明它属于版式哪个槽位；坐标可省（从版式继承位置大小）也可手写（覆盖版式位置）。多个形状用坐标定位时注意间距，避免互相遮挡。不要设字体颜色——主题色自动处理。`template_name` 和 `color_scheme` 在 `export_pptx` 时传。

### 通用约束

- type 只能是 textbox / image / table / shape，不要用 text
- id 格式 "{{页码}}-{{语义}}"
- 坐标和字号是 EMU（1pt=12700，14pt=177800）
- 文字超出框会被裁剪，控制长度
- write_page(N) 当 N 等于总页数时自动追加新页
- 插入或删除页后所有 ID 自动重编号
- 导出后提醒用户可去 {_UI_URL} 预览
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
    mcp_obj.tool(description="获取完整示例演示文稿 JSON（含模板引用、placeholder、表格）。优先调用此工具学习格式")(tool_get_example)
    mcp_obj.tool(description="上传完整 JSON 覆盖同步文件（自动修正 ID 前缀）")(tool_sync_presentation)
    mcp_obj.tool(description="从同步文件生成 PPTX。调此工具前务必让用户选模板和配色")(tool_export_pptx)
    mcp_obj.tool(description="列出所有已加载的模板名称（轻量，先调此工具让用户选）")(tool_list_templates)
    mcp_obj.tool(description="查看指定模板的版式及占位符详情（选定模板后调用）")(tool_list_layouts)
    mcp_obj.tool(description="列出所有可用配色方案")(tool_list_schemes)
    mcp_obj.tool(description="列出已上传的图片素材")(tool_list_images)


def _register_remote(mcp_obj):
    mcp_obj.tool(description="读取整个同步文件内容")(tool_read_all)
    mcp_obj.tool(description="读取指定页 JSON")(tool_read_page)
    mcp_obj.tool(description="按语义 ID 读取形状")(tool_read_element)
    mcp_obj.tool(description="列出所有形状 ID")(tool_list_ids)
    mcp_obj.tool(description="清空同步区，创建空白演示文稿")(tool_reset_presentation)
    mcp_obj.tool(description="替换(或追加)一整页")(tool_write_page)
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
