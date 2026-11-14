# 秩序册名单 · DeepSeek Harness 版

这是一个从零重写的本地网页：输入 PDF、Excel、Word 或图片，输出固定版式的秩序册名单 DOCX。

核心不是“再写一套自研 Agent 循环”，而是：

> 架构说明见 [docs/architecture.md](docs/architecture.md)。

```text
FastAPI 页面
  → 每个连续文件组一个 DeepSeek Harness Session
  → DSH 自己的 turn / step / tool loop
  → roster MCP 工具（读资料、接收分段、保存片段）
  → 应用合并片段并生成 Word
```

多文件任务不再把所有 PDF 塞进一个上下文：普通文件各自一个 DSH Session；连续页（如报名表1/2/3.jpg）合成一个文件组，共用一个 DSH Session，因此可以回看上一页/下一页拼接身份证号。应用只负责文件组调度、片段合并、来源校验和 Word 生成。

这不是把模型上下文限制成 65k。每个 DSH Session 的 `contextWindow` 使用模型真实容量（`ROSTER_CONTEXT_WINDOW`，默认 1,000,000）；DSH 的 compaction 负责处理单个 Session 内的长上下文。文件组只决定“哪些资料必须互看”，不是缩小模型窗口。

应用不再维护模型请求循环、批次数、纠错次数或“同一错误重试三次”的策略。模型请求重试、工具调用、上下文压缩和会话续跑都由 DSH Core 负责。

## 启动

需要 Python 3.11+、uv，以及阿里云百炼的 `DASHSCOPE_API_KEY`。

```sh
uv sync
cp .env.example .env          # 填入 DASHSCOPE_API_KEY
./start.sh                    # 默认 http://127.0.0.1:8878
```

也可以前台运行：

```sh
uv run --env-file .env uvicorn app:app --host 127.0.0.1 --port 8878
```

停止：`./stop.sh`。生成的文件在 `data/tasks/<任务号>/`。

## 保留了什么

- **前端基础页面**：上传、赛事/组别、七种名单版式、进度、核对、下载、生成历史，以及失败任务的“重新整理”。
- **读取 Skill**：`skills/intake/office`、`office-docx`、`office-xlsx`，由 DSH 的 `skill` 工具加载。技能里的 shell 命令不执行，实际读取统一走 roster 工具。
- **名单生成 Skill**：`roster_skill/` 保留原 `event-roster-docx` 的母版、脚本和版式规则。
- **文档读取层**：`documents.py` / `structure.py` 的通用版面、合并单元格、扫描页、长截图和跨页读取能力。

## Harness 工具

DSH Session 中只暴露这些 roster 工具（以及 `skill`）：

| 工具 | 作用 |
|---|---|
| `mcp__roster__list_documents` | 查看文件、页、工作表、正文单元和已读状态 |
| `mcp__roster__read_document` | 按任意页/行读取文字、PDF 文字坐标或原页图；可反复回读 |
| `mcp__roster__append_rows` | 逐人/配对/单位联系人母版分批提交，按 `group` 自动分段 |
| `mcp__roster__append_roster` | 横排/团体/双栏/分组母版分批提交一个组或 section |
| `mcp__roster__finish_roster` | 校验来源；单文件生成 Word，多文件只保存本文件片段 |
| `mcp__roster__submit_roster` | 小名单一次提交，或 `status=review` 请求人工核对 |

大名单使用 `append_rows` / `append_roster` 分批提交，避免模型一次生成过大的工具参数 JSON。工具校验失败会作为 MCP `isError` 返回，DSH 可以在同一会话中继续读取、修正、再次提交。

多文件时普通文件各自一个 DSH Session；文件名去掉末尾数字后相同的连续页（如报名表1/2/3.jpg）自动并成一个文件组，共用一个 Session，必须按顺序读完整组。某个文件组失败时会继续处理其他文件组，失败组可以在页面上单独“重试此文件”，全部完成后由应用合并生成一份 Word。人工核对时，用户答复恢复该文件组原来的 DSH Session，而不是新建一个“重试批次”。

## 目录

```text
app.py                 FastAPI 页面与 API
harness_runner.py      启动 DSH Session；没有应用侧模型循环
mcp_server.py          roster MCP 工具
document_bank.py       资料目录、读取、来源编号、已读进度
roster_validation.py   母版结构、source 和连续范围校验
renderer.py            调用 event-roster-docx 生成 Word
storage.py             任务 JSON、历史、事件
documents.py           PDF/Excel/Word/图片读取层
structure.py           版面、表头、扫描页和跨页表格识别
presets.py             七种母版说明
skills/intake/         保留的 DSH Office 读取 Skill
roster_skill/          保留的名单生成 Skill
static/                前端页面、样式和母版预览
data/tasks/            任务、上传件、DSH home、Word、落位报告
```

## 支持范围

- Excel：`.xlsx`、`.xls`、`.et`
- Word：`.docx`；`.wps` 使用 macOS `textutil`；`.doc` 需要系统有 LibreOffice
- PDF：有文字层或扫描页
- 图片：`.png`、`.jpg`、`.jpeg`、`.webp`
- 单份最多 300 MB，整批最多 3 GB，单任务最多 100 份

当前页面不提供自定义 Word 母版上传；使用内置七种母版。自定义母版可以按 `roster_skill/references/layout.md` 扩展。

## 测试

```sh
.venv/bin/python -m unittest discover -s tests -v
```

测试覆盖来源校验、连续范围、容器误判、Word 读取和 DOCX 生成。
