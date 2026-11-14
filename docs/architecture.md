# 架构：让 Harness 做 Agent，让应用做应用

## 1. 多文件任务：一个连续文件组一个 DSH Session

`harness_runner.run_task()` 按文件组调度：普通文件各自启动一个独立的 DSH Session；文件名去掉末尾数字后相同的连续页（如报名表1/2/3.jpg）合成一个文件组，共用一个 Session，按顺序读取整组并保存片段；所有片段完成后由应用合并、校验、生成 Word。

单文件任务只有一个 Session；多文件任务不会把所有 PDF 放进同一个 Session。连续页同组是为了让模型能回看上一页/下一页，拼接跨页的身份证号、号码和续表行。

每个 Session 都使用模型真实上下文容量：`contextWindow = ROSTER_CONTEXT_WINDOW`，默认 `1,000,000`（qwen3.8-flash）。DSH 的 `compaction-basic` 在该 Session 接近上限时压缩历史；文件分组不是把窗口调小。

实测把 12 份互不相关的 PDF 放进一个 Session 时，模型在上下文尚未接近 1M 前就出现了工具参数结构错误；因此默认按文档组分 Session。DSH 本身支持多 Session、子 Agent 和 compaction，文档分组属于 DSH 的会话编排，不是另写 Agent 循环。

每个文件组的 Session 只做四件事：

1. 写一次性的 DSH patch（禁用无关内置工具，挂载 skill、MCP 和百炼模型路由）。
2. `harness.start_session()` 或 `harness.resume_session()`。
3. 调用一次 `session.run(prompt)`。
4. 读取 MCP 工具写入的 `fragments/<文件>.json` / `review.json`，落文件状态。

没有 `while rounds < N`，没有 `MAX_OUTPUT_CORRECTIONS`，没有应用侧 `tool_call → tool_result` 拼装。DSH 的 Agent Loop 负责请求、工具执行、错误观察、重试、压缩和结束；应用只负责文件级顺序调度和片段合并。

## 2. 工具错误就是观察

MCP 工具通过 `CallToolResult(isError=True)` 返回校验错误。模型在同一个 Session 里看到错误文本，可以：

- 回读前一页/表头；
- 换 `text` / `image` / `layout` 视图；
- 修正 `source`；
- 继续 append 或 finish。

PDF 文字、Word 正文和 Excel 工作表按原始行号分页。`next_start` 指向本次响应之后的第一行，即使结果因体积上限提前分段也是如此。工具记录已返回的行范围，全部覆盖后标记完成。含图像内容的 PDF 页使用原页图完成读取；相邻页的文字预览只登记实际返回的来源。

Office 结构化读取沿用 DSH Office Skill 使用的 `openpyxl` 和 `python-docx`，PDF 文字、坐标和页图由 PyMuPDF 提取。MCP 提供固定的页、表、行编号，提交时校验来源，并将多文件会话的来源限制在本组。

只有 Harness 运行失败、或 Agent 结束但没有提交名单时，任务才进入 `error`。

## 3. 分段提交

大名单的工具参数 JSON 很容易在模型侧被截断或写坏。因此写操作拆成：

```text
append_rows / append_roster   （多次，小参数）
        ↓
finish_roster                 （一次，校验 + 渲染）
```

`append_rows` 的每一行可以带 `group`；工具按 `group` 合并分段，所以“部门组（一）”“部门组（二）”不会变成两个 Word 分段。

来源校验支持精确行/页编号，也支持连续范围（如 `段落 4-5`、`第 2-3 页`），但要求范围内每个编号都能在读取记录中找到。

## 4. 读取 Skill 与输出 Skill 分离

- `skills/intake/*` 是 DSH `skill` 工具加载的读取说明。它们教模型怎么理解 Office/PDF，但实际读取由 roster 工具执行。
- `roster_skill/*` 是确定性的 DOCX 母版与落位脚本。模型不直接改 Word，只提交结构化名单。

这保证模型负责“理解资料”，程序负责“可重复的排版和来源校验”。

## 5. 人工核对

`submit_roster(status=review)` 把问题写入 `review.json`。应用关闭本轮 DSH 运行时后发布待核对状态。用户答复后，`harness_runner` 用同一个 `dsh_home`、同一个 `session_id` 调用 `resume_session()`，所以模型保留此前的读取记录和工具结果。

## 6. 失败分类

| 情况 | 行为 |
|---|---|
| MCP 工具参数/来源/姓名校验失败 | `isError` 返回，DSH 继续 |
| Word 渲染失败 | 合并阶段报错，保留各文件片段 |
| 模型请求超时/5xx | DSH `llm-retry` 按配置重试 |
| 模型响应格式错误（`MALFORMED_RESPONSE`，包括非法工具参数 JSON） | 百炼路由将该错误加入 DSH `retryableCodes`，DSH 在同一回合重跑失败步骤，最多重试两次 |
| Agent 完成但没有片段 | 只标记该文件 `error`，继续其他文件 |
| 合并时缺少文件片段 | 停止生成 Word 并报告缺少的文件编号 |
| 某文件失败 | 其他文件继续；完成后可单独“重试此文件” |
| DSH 会话恢复失败 | 该文件 `error`，可单独重试 |
| 用户核对 | 任务 `review`，恢复该文件会话，不生成总 Word |

## 7. 为什么不再需要六个 Engine

旧应用为 custom / Pydantic AI / Pi / DSH / LangGraph / OpenAI 各写一套循环和纠错语义，导致同一份资料在不同引擎上的失败行为不一致。新应用只保留一个实际的 Agent Core：DSH。模型适配器、工具协议和会话协议由 DSH 提供，应用只保留一个工具边界。
