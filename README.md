# AgentDeck

AI Agent 驱动的幻灯片制作引擎 —— 模型只需编辑 JSON，MCP 协议桥接到渲染引擎，输出高保真 PPTX。

## 功能

- **JSON ↔ PPTX 双向转换** — 解析和构建互为逆操作，100% 往返一致
- **MCP 协议集成**（核心）— 通过 MCP Server 暴露 17 个工具，可接入 opencode、Cursor、Claude Desktop 等支持 MCP 的客户端
- **两种工作模式** — 有文件访问时用本地模式（模型直接编辑 JSON），没有则用远程模式（逐元素构建）
- **模板/配色系统** — 支持多模板和多配色方案，内容与样式分离，导出时自由切换
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
- `sync_presentation(json)` / `export_pptx(path)` — 同步 JSON 并导出
- `list_templates()` / `list_layouts(name)` — 查看模板和占位符
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
│  Agent 编辑 JSON（利用自带文件读写能力）              │
│         │ sync_presentation()                       │
│         ▼                                           │
│  AgentDeck Engine（OOXML 复杂度隔离区）               │
│  ┌─────────────────────────────────────────┐         │
│  │  JSON ↔ PPTX 双向转换（100% 往返一致）    │         │
│  │  模板/配色/内容各层独立可替换               │         │
│  └─────────────────────────────────────────┘         │
│         │ export_pptx()                              │
│         ▼                                           │
│  输出 .pptx 文件                                     │
└─────────────────────────────────────────────────────┘
           ↕ Web UI（人工预览和微调）
```

**本地模式**（推荐）：模型在项目目录编辑 JSON 文件，通过 `sync` 和 `export` 两个 MCP 命令完成迭代。

**远程模式**：模型通过 MCP 工具的 `write_page` / `write_element` 逐元素构建，适合无文件访问的客户端。

## 项目结构

```
├── ppt_render_engine/     # 渲染引擎（FastAPI）
│   ├── src/api/           # REST API 路由
│   ├── src/core/          # PPTX 构建/解析核心
│   ├── src/engine/        # 布局预设和预览
│   ├── src/models/        # JSON Schema（Pydantic）
│   └── web/               # Web UI
├── mcp_server/            # MCP 协议层
└── docker/                # Docker 部署
```

## 技术栈

FastAPI / python-pptx / lxml / Pydantic / MCP Python SDK / Pillow / Docker
