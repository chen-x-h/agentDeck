# AgentDeck

**AI Agent 驱动的幻灯片制作引擎。JSON ↔ PPTX 双向转换，让模型像编辑文本一样制作演示文稿。**

## 项目思想

传统 PPT 生成方案是「模型写 JSON → 引擎出 PPTX」的单向流水线，模型无法迭代修改。每次改内容都要全量重建，而且 PPTX 的 OOXML 格式复杂（模板、配色、占位符、表格、图片交织），引擎的解析和构建必须精确对称，否则往返不一致。

这个项目换了一种思路：

**用 JSON 文件做中间媒介。** Agent 用自带的文件读写能力直接编辑 JSON，引擎只负责两件事：同步（sync）和导出（export）。迭代工作流简化为：编辑本地 JSON → sync → 再编辑 → 再 sync → export 出 PPTX。不需要逐元素 API 调用，不需要自定义通信协议。

引擎内部实现 JSON ↔ PPTX 双向转换，解析和构建互为逆操作，100% 往返一致。OOXML 的复杂度封闭在引擎内部，Agent 和用户看到的始终是干净的 JSON。

同时提供 Web UI 作为人工兜底——模型搞不定的时候，用户可以直接拖拽 JSON 文件到浏览器预览和微调。

## 快速启动

### 本地运行

```bash
# 安装依赖
cd ppt_render_engine
pip install uv
uv sync --no-dev

# 启动渲染引擎
uv run python run.py
# 服务监听 http://127.0.0.1:8000
```

新开终端：

```bash
# 启动 MCP Server（本地文件模式，默认）
cd mcp_server
uv sync --no-dev
uv run python -m mcp_server
# 服务监听 http://127.0.0.1:8010
```

### Docker

```bash
cd docker
docker compose up -d
# 引擎: http://127.0.0.1:8000
# MCP:  http://127.0.0.1:8010
```

## 使用方式

### 推荐：本地文件模式（local mode）

模型有文件访问权限时使用。Agent 在工作目录编辑 JSON 文件，通过 MCP 同步和导出：

```
1. get_example()         # 查看格式
2. shape_reference()     # 查看形状类型参考
3. 编辑本地 JSON 文件     # 模型自带文件编辑能力
4. sync_presentation()   # 同步到引擎
5. export_pptx(path)     # 生成 PPTX
```

### 备用：远程模式（remote mode）

模型无文件访问权限时，通过 `--mode remote` 启动 MCP Server，用工具逐页构建：

```
1. get_example()
2. list_templates() → list_layouts(name)
3. write_page() / write_element() 逐页构建
4. export_pptx(path)
```

### Web UI

浏览器打开 `http://127.0.0.1:8000`，拖拽 JSON 或 ZIP 文件即可预览和编辑。

## 架构

```
┌─────────────────┐     ┌──────────────────┐     ┌──────────────┐
│  LLM (AI Agent) │◄───►│  MCP Server      │◄───►│  Render      │
│  - 编辑 JSON    │     │  sync/export     │     │  Engine      │
│  - 调用 MCP 工具│     │  8/17 tools      │     │  FastAPI     │
└─────────────────┘     └──────────────────┘     └──────┬───────┘
        ▲                                               │
        │ 直接编辑 JSON 文件                              │ HTTP
        ▼                                               ▼
┌─────────────────┐                            ┌──────────────┐
│  Local JSON     │                            │  PPTX File   │
│  presentation   │◄──── sync ────────────────►│  (output)    │
└─────────────────┘                            └──────────────┘
```

关键分层：
- **JSON** — Agent 和引擎之间的契约，也是用户可以直接编辑的文本
- **MCP** — 薄通信层，只做 sync/export/查询，不做逐元素操作（local 模式）
- **Engine** — OOXML 复杂度隔离区，双向转换保证往返一致
- **Web UI** — 人工兜底，JSON 可视化预览和编辑

## 亮点

- **零协议开销**：Agent 用自带的文件读写能力编辑 JSON，不需要为每个操作定义 API
- **双向对称**：构建和解析互为逆操作，往返测试 100% 保真率
- **本地 + 远程双模式**：有文件访问用本地模式，没有就用远程模式
- **模板/配色解耦**：内容不依赖特定模板，导出时可自由切换
- **人工兜底**：Web UI 预览和微调，模型搞不定时人上
- **性能**：单页构建 ~14ms，解析 ~5ms
- **Docker 部署**：一键启动，配置驱动
