# AgentDeck

AI Agent 驱动的幻灯片制作引擎 —— 模型只需编辑 JSON，MCP 协议桥接到渲染引擎，输出高保真 PPTX。

## 功能

- **JSON ↔ PPTX 双向转换** — 解析和构建互为逆操作，100% 往返一致
- **MCP 协议集成**（核心）— 通过 MCP Server 暴露 18+ 工具，可接入 opencode、Cursor、Claude Desktop 等支持 MCP 的客户端
- **两种工作模式** — 有文件访问时用本地模式（模型直接编辑 JSON），没有则用远程模式（逐元素构建）
- **模板/配色系统** — 支持多模板和多配色方案，内容与样式分离，导出时自由切换
- **设计风格系统** — Markdown 格式的设计规范，绑定模板、配色、字号体系和视觉编排规则，一次选定全程遵循
- **图片素材库** — 支持从 URL 下载或本地文件上传到素材库，模型设计时优先复用
- **容器（children）** — 一个占位符区域可放置多个重叠元素（色块+文字），用容器包裹
- **百分比坐标** — 模型输出 0-100 百分比坐标，引擎自动换算为 EMU，降低心智负担
- **文字自动收缩** — 文字超限时引擎自动缩小字号适应形状，无需模型手动调
- **Web UI** — 浏览器预览和微调，模型搞不定时人工兜底
- **Docker 一键部署** — 配置驱动，开箱即用

## MCP 集成（核心能力）

AgentDeck 通过 MCP Server 暴露工具，让 AI 客户端可以直接管理演示文稿：

```
# 启动 MCP Server
cd mcp_server && uv run python -m mcp_server

# 默认监听 http://127.0.0.1:8010/sse
```

**工具示例：**
- `get_example()` / `shape_reference()` — 学习 JSON 格式
- `list_design_styles()` / `get_design_style(name)` — 查看和选用设计风格
- `sync_presentation(file_path)` / `export_pptx(path, template, scheme)` — 同步 JSON 并导出
- `list_templates()` / `list_layouts(name)` — 查看模板和占位符
- `list_schemes()` / `get_color_scheme(name)` — 查看配色板和色值
- `list_images()` / `upload_image_from_url()` / `upload_image_from_file()` — 管理图片素材
- `write_page()` / `write_element()` — 远程模式下逐元素编辑

> 在 opencode 的 `opencode.json` 中配置 `"mcpServers"` 即可让 Agent 直接操作幻灯片。

## 快速启动

### 本地运行

```bash
# 终端 1：启动渲染引擎
cd ppt_render_engine
pip install uv && uv sync --no-dev
uv run python run.py          # http://127.0.0.1:8000

# 终端 2：启动 MCP Server
cd mcp_server
uv sync --no-dev
uv run python -m mcp_server   # http://127.0.0.1:8010/sse
```

或者一键启动（Windows）：

```bash
start.bat [local|remote]
```

### Docker

```bash
cd docker
docker compose up -d
# 引擎 http://127.0.0.1:8000
# MCP  http://127.0.0.1:8010
```

## 工作流

```
┌─────────────────────────────────────────────────────┐
│  Step 1: 选设计风格（list_design_styles → get_style）│
│  Step 2: 熟悉模板布局和配色（list_layouts + scheme）  │
│  Step 3: Agent 编辑 JSON                            │
│         │ sync_presentation(file_path)              │
│         ▼                                           │
│  AgentDeck Engine（OOXML 复杂度隔离区）               │
│  ┌─────────────────────────────────────────┐         │
│  │  JSON ↔ PPTX 双向转换（100% 往返一致）    │         │
│  │  模板/配色/内容各层独立可替换               │         │
│  │  校验：非法字段/溢出/重叠/系统占位符         │         │
│  └─────────────────────────────────────────┘         │
│         │ export_pptx(path, template, scheme)        │
│         ▼                                           │
│  输出 .pptx 文件                                     │
└─────────────────────────────────────────────────────┘
           ↕ Web UI（人工预览和微调）
```

**本地模式**（推荐）：模型在项目目录编辑 JSON 文件，通过 `sync` 和 `export` 两个 MCP 命令完成迭代。

**远程模式**：模型通过 MCP 工具的 `write_page` / `write_element` 逐元素构建，适合无文件访问的客户端。

## 项目结构

```
├── ppt_render_engine/        # 渲染引擎（FastAPI）
│   ├── src/api/              # REST API 路由（设计风格、图片等）
│   ├── src/core/             # PPTX 构建/解析/设计风格/配色/模板核心
│   ├── src/engine/           # 布局预设、校验器（溢出/重叠/容器）
│   ├── src/models/           # JSON Schema（Pydantic，extra=forbid）
│   ├── storage/              # 存储：模板PPTX / 配色 JSON / 设计风格 .md
│   └── web/                  # Web UI
├── mcp_server/               # MCP 协议层（18+ 工具，严格工作流指令）
└── docker/                   # Docker 部署
```

## 技术栈

FastAPI / python-pptx / lxml / Pydantic / MCP Python SDK / Pillow / Docker
