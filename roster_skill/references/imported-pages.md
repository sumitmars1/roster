# 内置母版的分段名单字段

按 [内置母版目录](builtin-pages.md) 选择 `--preset`，再用本页的字段说明整理名单。母版提供名册结构、合并关系和字体，生成时按资料内容与页宽分配列宽和续行。

## 输入和命令

将本次报名资料整理为 UTF-8 JSON。一个文件可以有多个分段。同一 `--preset` 下可在一个文件中提供多个分段，也可按顺序重复传入 `--data`。同一张表或同一组队伍的续表应在整理时并到对应数组。分段按输入顺序排列。

```json
{
  "sections": [
    {
      "data": {
        "groups": [{
          "group": "部门组",
          "teams": [{
            "team": "示例联队",
            "affiliation": "（甲单位、乙单位）",
            "staff": [
              {"role": "联络员", "name": "林春禾", "phone": "000-0000", "source": "报名表.xlsx / 第3行"}
            ],
            "players": [
              {"name": "苏远山", "source": "报名表.xlsx / 第4行"}
            ]
          }]
        }]
      }
    }
  ]
}
```

```sh
uv run --script scripts/roster.py pages --preset inline-roster \
  --data 名单.json --output 人员名单.docx --audit 名单落位.json
```

下表中的列表字段按原资料顺序填，人员、电话、号码和签号使用文本，保留前导零。序号由脚本按所在列表生成。每条记录可带 `source`，优先取具体人的原文件与行号，缺省沿用上层来源。加 `needs_review: true` 的记录须核清后才能生成。有值但无法落位的字段会报告对应记录，不会默默丢弃。

| 母版 | `data` 的字段和行结构 |
|---|---|
| 横排名单（多岗位） | `groups: [{group, teams}]`。每队 `team`、`players: [{name}]`，可填 `affiliation` 和 `staff: [{role, name, phone}]`。按姓名宽度和可用页宽续行。岗位标签按来源填写，同一职务可重复。可设 `show_directory: true` 加队伍索引。 |
| 单位联系人表 | `rows: [{team, contact, phone}]`，`team` 是参赛单位，`contact` 是领队或联络员。可填 `title` 和 `orientation`（`auto`、`portrait`、`landscape`），默认根据内容自动选择横竖版。 |
| 分组队伍表 | `groups: [{group, teams: [{team}]}]`，队伍签位及现场队伍字样如在来源中出现，放入 `team`，保留原顺序。 |
| 团体宽表（含联系方式） | `title`、`teams: [{team, leader, phone, players: [{name}]}]`。可填 `player_heading`。有电话时显示联系方式列，按姓名宽度和页宽计算每行格数。 |
| 逐人名单表 | `title`、`rows: [{team, name, event, sign}]`。各组使用同一母版分别建段，按资料填比赛项目与签号。 |
| 搭档配对表 | `title`、`pairs: [{team, name, partner, event, sign}]`。一个对象是一对搭档，按两行排，二人共用序号、单位和签号。同一个 `event` 原样显示在两行。 |

电话、领队等缺项可省略，参赛人姓名、参赛单位或队名、组名按对应表结构填写。逐人名单和搭档配对表按资料填写比赛项目，资料未提供时项目栏留空。直接使用三份旧式 DOC 作为报名数据时，先转为 DOCX 或人工核读后整理 JSON。需要组委会、裁判或其他赛务名单时，使用用户指定的对应版面。
