# 配置与母版

运行需要 Python 3.10+ 和 `docxtpl`、`openpyxl`。以下命令假定当前目录是 Skill 目录，且安装了 uv。已有依赖时也可用 `python scripts/roster.py`。

## 提取 Excel 和 Word 表格

```sh
uv run --script scripts/roster.py extract --config sources.json --output roster.json
```

`sources.json` 中路径相对于配置文件。行号和 Word 顶层表格序号从 1 开始。每个来源指定 `columns`，值为源表头的准确文本。单队或单角色表可用 `constants` 提供固定值，姓名始终取自列。可用 `first_data_row` 和 `last_data_row` 排除附注或落款。Excel 的队伍、角色列若是单列纵向合并，读取合并区域左上角的值。未合并的空格保持空值。

```json
{
  "sources": [
    {
      "path": "报名汇总.xlsx",
      "sheet": "运动员",
      "header_row": 2,
      "columns": {"team": "代表队", "name": "姓名"},
      "constants": {"role": "运动员"}
    },
    {
      "path": "教练名单.docx",
      "table": 1,
      "header_row": 1,
      "columns": {"team": "队伍", "name": "教练姓名"},
      "constants": {"role": "教练"}
    }
  ]
}
```

整行空白会跳过。映射字段部分为空时报告原行位置。提取适用于单行表头、每行一人。多级表头、名单分布在多列、Word 段落名单或图片，先据原件转成下面的统一 JSON，维持原来的姓名和顺序。Excel 公式使用文件中已缓存的结果。

```json
{
  "people": [
    {"team": "甲队", "role": "运动员", "name": "陈一", "source": "报名汇总.xlsx / 运动员 / 第3行"},
    {"team": "甲队", "role": "教练", "name": "张二", "source": "教练名单.docx / 表1 / 第2行"},
    {"team": "乙队", "role": "运动员", "name": "李三", "source": "报名截图.png / 第6行", "needs_review": true}
  ]
}
```

待复核记录会阻止生成。对照原件确认后更新姓名并清除该标记。多个名单文件可按顺序合并各自的 `people` 数组，也可用多个 `--roster` 参数。脚本按来源顺序排人，保留同组同名记录，在报告中附来源。

## 固定格位

下面示例的甲队运动员每行两个名字，有两行，先横向填，再到下一行。模板中保留实际的格线、边距和字体，并放入四个占位符。

| 位置 | 第 1 格 | 第 2 格 |
|---|---|---|
| 第 1 行 | `{{ players_a.r1c1 }}` | `{{ players_a.r1c2 }}` |
| 第 2 行 | `{{ players_a.r2c1 }}` | `{{ players_a.r2c2 }}` |

```json
{
  "sections": [
    {
      "id": "players_a", "team": "甲队", "role": "运动员",
      "mode": "fixed", "columns": 2, "rows": 2, "direction": "row",
      "max_name_chars": 8
    },
    {
      "id": "coach_a", "team": "甲队", "role": "教练",
      "mode": "fixed", "columns": 1, "rows": 1, "direction": "row"
    }
  ]
}
```

教练使用 `{{ coach_a.r1c1 }}`。`direction` 取 `row`（从左向右，再向下）或 `column`（从上向下，再向右）。矩阵坐标来自样稿，可把下一页的固定格接在同一矩阵后。未用格位填空。`max_name_chars` 可选，按 Unicode 字符计数，实际显示宽度仍以渲染核对为准。每个占位符必须单独放在一个 Word 文字片段（run）中，且使用统一格式。文字片段可与标签、其他姓名片段、制表符位于同一段落。确认模板里的占位符覆盖了所有配置格位。

## 重复样板行

在布局配置中改用以下 section。示例最多三行，每行两格，空名单也保留一行。用样板行定义列宽和姓名字号。`min_rows` 默认为 0。

```json
{
  "id": "players_b", "team": "乙队", "role": "运动员",
  "mode": "repeat", "columns": 2, "min_rows": 1, "max_rows": 3,
  "direction": "row"
}
```

在母版的目标表格中放三行。前后控制行会被模板引擎移除。

| 作用 | 第 1 格 | 第 2 格 |
|---|---|---|
| 开始控制行 | `{%tr for row in players_b %}` | |
| 样板行 | `{{ row.c1 }}` | `{{ row.c2 }}` |
| 结束控制行 | `{%tr endfor %}` | |

重复模式由 Word 决定实际跨页位置。固定页、指定节开始或指定续页位置用预排好的固定格位。保留母版的分页与分节属性。

## 生成与检查

```sh
uv run --script scripts/roster.py render --template 母版.docx --layout layout.json \
  --roster roster.json --output 秩序册.docx --audit 名单落位.json
```

布局覆盖所有要输出的队伍与角色，一组只能路由到一个区块。母版上的变量仅用于已配置的名单区块。其他章节可保留为 Word 原生内容。脚本遇到未知组别、容量不足、待复核姓名、未填充格位或版式设置变化时退出，并指出原因。报告包括每个区块的行列数、实填人数、每人的来源及落位。同组同名作为提醒。
