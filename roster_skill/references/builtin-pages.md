# 内置名单母版

按名单结构选择 `uv run --script scripts/roster.py pages --preset …`。同一结构按报名字段、姓名长度和页宽排版。

| 母版 | `--preset` | 排版 |
|---|---|---|
| [双栏编号表](../assets/two-column-numbered.docx) | `two-column-numbered` | 先排左栏再排右栏，每栏有序号、姓名和号码。 |
| [横排名单](../assets/inline-roster/default.docx) | `inline-roster` | 每队显示资料中提供的岗位、电话及运动员。姓名可附号码，按姓名宽度自动换行，可加队伍索引。 |
| [团体宽表](../assets/team-wide/default.docx) | `team-wide` | 队伍和运动员同表，有电话时加联系方式列。按页宽和姓名宽度决定每行格数，超出后续行。 |
| [逐人名单表](../assets/individual-table/default.docx) | `individual-table` | 每人一行，按单位、姓名、项目、签号内容分配列宽。 |
| [搭档配对表](../assets/paired-table/default.docx) | `paired-table` | 一对两行，共用序号、单位和签号，列宽随内容调整。 |
| [单位联系人表](../assets/unit-contact-table/default.docx) | `unit-contact-table` | 每单位一行，按内容分配列宽和选择横竖版，可指定方向。 |
| [分组队伍表](../assets/grouped-team-table/default.docx) | `grouped-team-table` | 每队一行，同组队伍合并组别列。 |

名单表（`team-wide`、`individual-table`、`paired-table`、`unit-contact-table`）的姓名格按参考样张排版：2 字姓名在中间空一个字宽，与 3 字姓名等宽而不撑满整格；团体宽表放不下 4 字姓名时按两字一行折行；其余母版按内容加宽姓名列。

分段 JSON 的结构见[分段名单字段](imported-pages.md)。双栏编号表和横排名单也可用报名 Excel 或合并清单。

## Excel 与多份报名资料

第 1 行应有「单位」「姓名」各一列。队伍名称只在组首填写时，后续空白沿用同队。可识别「球衣」「球衣号码」「号码」和岗位列。号码有前导零时，Excel 应保存为文本。

```sh
uv run --script scripts/roster.py pages --preset inline-roster \
  --xlsx 报名汇总.xlsx --sheet '部门=部门组' --sheet '街镇=街镇组' \
  --output 人员名单.docx --audit 名单落位.json
```

多份资料按使用顺序写进 `--manifest`。Excel 直接引用，PDF、Word 和图片先提取为组队 JSON。每人可在 `source` 写原文件和页、表或行号。续表按组名和队名合并，保留同名人员和报名顺序。同队多份完整版本冲突时先核对。

```json
{
  "sources": [
    {"xlsx": "报名汇总.xlsx", "sheets": ["部门=部门组"]},
    {"data": "PDF补充名单.json"}
  ]
}
```

组队 JSON 示例：

```json
{
  "groups": [
    {
      "title": "部门组",
      "teams": [
        {
          "name": "甲队",
          "leader": "王平",
          "leader_phone": "000-0000",
          "coach": "李明",
          "players": [
            {"name": "周晓", "number": "07", "source": "报名表.pdf / 第2页第1格"}
          ]
        }
      ]
    }
  ]
}
```

双栏编号表可填领队、教练及对应电话，另有 `assistant_1`、`assistant_2`、`assistant_3` 和 `medic` 格位。横排名单可通过分段 JSON 的 `staff` 数组填任意岗位。旧版的母版名称和分段 `layout` 可继续读取，排版由对应母版按内容计算。
