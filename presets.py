import json


PRESETS = {
    "individual-table": ("逐人名单", "每人一行，列出单位、姓名、项目及签号。", '{"sections":[{"data":{"title":"标题","rows":[{"team":"单位","name":"姓名","event":"项目","sign":"签号","source":"定位"}]}}]}'),
    "inline-roster": ("横排名单", "按队列出岗位和运动员，姓名横向排列。", '{"sections":[{"data":{"groups":[{"group":"组别","teams":[{"team":"队名","source":"定位","staff":[{"role":"岗位","name":"姓名","phone":"电话","source":"定位"}],"players":[{"name":"姓名","number":"号码","source":"定位"}]}]}]}}]}'),
    "two-column-numbered": ("双栏编号", "先排左栏，再排右栏，保留运动员号码。", '{"groups":[{"title":"组别","teams":[{"name":"队名","source":"定位","leader":"领队","coach":"教练","leader_phone":"领队电话","coach_phone":"教练电话","players":[{"name":"姓名","number":"号码","source":"定位"}]}]}]}'),
    "paired-table": ("搭档配对", "一对两行，共用单位、序号和签号。", '{"sections":[{"data":{"title":"标题","pairs":[{"team":"单位","name":"姓名","partner":"搭档","event":"项目","sign":"签号","source":"定位"}]}}]}'),
    "team-wide": ("团体宽表", "按队列出领队、联系方式及运动员。", '{"sections":[{"data":{"title":"标题","teams":[{"team":"队名","source":"定位","leader":"领队","phone":"电话","players":[{"name":"姓名","source":"定位"}]}]}}]}'),
    "unit-contact-table": ("单位联系人", "每单位一行，列出联系人及电话。", '{"sections":[{"data":{"title":"标题","rows":[{"team":"单位","contact":"联系人","phone":"电话","source":"定位"}]}}]}'),
    "grouped-team-table": ("分组队伍", "每队一行，同组共用组别栏。", '{"sections":[{"data":{"groups":[{"group":"组别","teams":[{"team":"队名","source":"定位"}]}]}}]}'),
}


def roster_output_schema(preset, source_ids):
    def schema(example):
        if isinstance(example, dict):
            properties = {key: schema(value) for key, value in example.items()}
            if "source" in properties and source_ids is not None:
                properties["source"]["enum"] = list(source_ids)
            if "group_ref" in properties and source_ids is not None:
                properties["group_ref"]["enum"] = ["", *source_ids]
            return {"type": "object", "properties": properties, "required": list(example),
                    "additionalProperties": False}
        if isinstance(example, list):
            return {"type": "array", "items": schema(example[0])}
        if isinstance(example, bool):
            return {"type": "boolean"}
        return {"type": "string"}

    roster = json.loads(PRESETS[preset][2])
    if preset == "two-column-numbered":
        group = roster["groups"][0]
        group["group_ref"] = ""
        group["teams"][0]["assignment_only"] = False
    output = schema({"status": "", "roster": roster, "note": "", "options": [""]})
    output["properties"]["status"]["enum"] = ["ok", "empty", "review"]
    return output
