from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pymupdf
from docx import Document
from docx.shared import Pt
from fastapi.testclient import TestClient
from PIL import Image

import app as web_app
import book_api
import config
import cover_art
import draw_runner
import projects
import template_parser
import templates_store
from book_assembler import assemble_book
from book_pdf import export_pdf
from draw_update import apply_draw_patches, inspect_target_book


class BookPipelineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.orig_tpl_dir = templates_store.TEMPLATES_DIR
        self.orig_proj_dir = projects.PROJECTS_DIR
        self.orig_draw_dir = draw_runner.DRAW_JOBS_DIR
        self.orig_data_dir = config.DATA
        templates_store.TEMPLATES_DIR = self.root / "templates"
        projects.PROJECTS_DIR = self.root / "projects"
        draw_runner.DRAW_JOBS_DIR = self.root / "draw_jobs"
        config.DATA = self.root
        templates_store.ensure_builtin_templates()

    def tearDown(self) -> None:
        templates_store.TEMPLATES_DIR = self.orig_tpl_dir
        projects.PROJECTS_DIR = self.orig_proj_dir
        draw_runner.DRAW_JOBS_DIR = self.orig_draw_dir
        config.DATA = self.orig_data_dir
        self.tmp.cleanup()

    def _make_sample_book(self, path: Path) -> Path:
        doc = Document()
        doc.add_paragraph("样本封面")
        doc.add_section()
        doc.add_paragraph("目录")
        doc.add_paragraph("一、竞赛规程\t1")
        doc.add_paragraph("二、各代表队名单\t3")
        doc.add_paragraph("三、竞赛日程表\t5")
        doc.add_section()
        doc.add_paragraph("一、竞赛规程")
        doc.add_paragraph("主办单位：测试某单位。本样例中的历史队员王小明不应泄漏。")
        doc.add_paragraph("二、各代表队名单")
        tbl = doc.add_table(rows=2, cols=6)
        for i, h in enumerate(["序号", "姓名", "号码", "序号", "姓名", "号码"]):
            tbl.cell(0, i).text = h
        for i, v in enumerate(["1", "王小明", "7", "2", "李小红", "9"]):
            tbl.cell(1, i).text = v
        doc.add_paragraph("三、竞赛日程表")
        doc.save(path)
        return path

    def _make_cover_img(self, path: Path) -> Path:
        im = Image.new("RGB", (600, 848), (30, 80, 150))
        im.save(path)
        return path

    def _make_schedule_xlsx(self, path: Path) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "赛程"
        ws.append(["竞赛日程表", None, None, None, None])
        ws.merge_cells("A1:E1")
        ws.append(["比赛时间", "场序", "组别", "对阵", "赛别"])
        ws.append([46281, 1, "男子A组", "A1 VS A2", "小组赛"])
        ws.append([46281, 2, "男子A组", "A3 VS A4", "小组赛"])
        ws.append([46282, 3, "半决赛", "68胜——69胜", "淘汰赛"])
        ws.append([46282, 4, "半决赛", "A组第一名 VS B组第二名", "交叉赛"])
        wb.save(path)
        return path

    def _make_source_docx(self, path: Path) -> Path:
        d = Document()
        p = d.add_paragraph("2026年测试赛事竞赛规程")
        p.alignment = 1
        d.add_paragraph("一、主办单位")
        d.add_paragraph("测试主办单位，包含完整文本与附件图片。")
        img_p = self._make_cover_img(self.root / "qr.png")
        d.add_paragraph().add_run().add_picture(str(img_p), width=Pt(90))
        d.save(path)
        return path

    def test_builtin_templates_and_sample_parse_strips_historical_data(self) -> None:
        tpls = templates_store.list_templates()
        ids = {t["id"] for t in tpls}
        self.assertTrue({"official-inline", "official-football", "team-wide-event", "fun-games-contacts"} <= ids)

        sample = self._make_sample_book(self.root / "sample.docx")
        parsed = template_parser.parse_sample_book_to_template(sample, name="足球模板样例", use_llm=False)
        self.assertEqual(parsed["roster_preset"], "two-column-numbered")
        self.assertFalse(parsed["builtin"])
        serialized = str(parsed)
        self.assertNotIn("王小明", serialized)
        self.assertNotIn("李小红", serialized)

        fetched = templates_store.get_template(parsed["id"])
        self.assertIsNotNone(fetched)
        self.assertTrue(templates_store.delete_custom_template(parsed["id"]))
        self.assertIsNone(templates_store.get_template(parsed["id"]))

    def test_assemble_book_cover_toc_sync_serial_dates_and_post_draw_update(self) -> None:
        tpl = templates_store.get_template("official-inline")
        cover_bg = self._make_cover_img(self.root / "bg.png")
        back_img = self._make_cover_img(self.root / "back.png")
        reg_docx = self._make_source_docx(self.root / "reg.docx")
        sched_xlsx = self._make_schedule_xlsx(self.root / "sched.xlsx")

        plan = {
            "event_title": "动也测试杯篮球邀请赛",
            "subtitle": "秩  序  册",
            "footer_lines": ["重庆市测试体育协会", "2026年9月"],
            "cover": {"mode": "background", "path": str(cover_bg)},
            "back_cover": {"path": str(back_img)},
            "chapters": [
                {
                    "id": "regulations",
                    "title": "竞赛规程",
                    "kind": "richtext",
                    "toc_visible": True,
                    "sources": [{"type": "docx", "path": str(reg_docx)}],
                },
                {
                    "id": "schedule",
                    "title": "竞赛日程表",
                    "kind": "schedule",
                    "toc_visible": True,
                    "landscape": False,
                    "sources": [{"type": "xlsx", "path": str(sched_xlsx)}],
                },
                {"id": "notes", "title": "记事", "kind": "notes", "toc_visible": False},
            ],
        }

        out_docx = self.root / "out_book.docx"
        out_pdf = self.root / "out_book.pdf"
        res = assemble_book(plan, tpl, out_docx, out_pdf)

        self.assertTrue(out_docx.is_file())
        self.assertTrue(out_pdf.is_file())
        self.assertGreaterEqual(res["pdf_pages"], 5)
        self.assertEqual([t["page"] for t in res["toc"]], [1, 2])

        import re

        with pymupdf.open(out_pdf) as pdf:
            # Page 1: Cover, Page 2: TOC, Page 3: Body page 1, Page 4: Body page 2
            toc_txt = re.sub(r"\s+", "", pdf[1].get_text("text"))
            self.assertIn("竞赛规程", toc_txt)
            self.assertIn("竞赛日程表", toc_txt)
            body2_txt = re.sub(r"\s+", "", pdf[3].get_text("text"))
            self.assertIn("9月16日", body2_txt)
            self.assertIn("A1VSA2", body2_txt)
            self.assertIn("68胜——69胜", body2_txt)
            # Back cover has no body page number
            last_txt = pdf[-1].get_text("text").strip()
            self.assertEqual(last_txt, "")

        # Post-draw XML inspection & scoped replacement
        insp = inspect_target_book(out_docx)
        sched_tbl = next(t for t in insp["tables"] if t["has_candidate_slot"])
        row1 = next(r for r in sched_tbl["rows"] if any("A1 VS A2" in c["text"] for c in r["cells"]))
        col_match = next(c["col"] for c in row1["cells"] if "A1 VS A2" in c["text"])

        upd_docx = self.root / "updated.docx"
        upd_pdf = self.root / "updated.pdf"
        upd_res = apply_draw_patches(
            out_docx,
            [
                {
                    "target": "cell",
                    "table": sched_tbl["table_index"],
                    "row": row1["row"],
                    "col": col_match,
                    "find": "A1 VS A2",
                    "replace": "动也文化 VS 钡栩传媒",
                    "reason": "小组赛第1场签位替换",
                },
                # Attempting to patch win/loss ref must be rejected cleanly by safety guard
                {
                    "target": "cell",
                    "table": sched_tbl["table_index"],
                    "row": row1["row"] + 2,
                    "col": col_match,
                    "find": "68胜——69胜",
                    "replace": "错误队——错误队",
                    "reason": "安全护栏测试",
                },
            ],
            upd_docx,
            upd_pdf,
        )
        self.assertEqual(upd_res["replaced_count"], 1)
        self.assertEqual(upd_res["skipped_count"], 1)

        with pymupdf.open(upd_pdf) as updf:
            full = re.sub(r"\s+", "", "".join(updf[i].get_text("text") for i in range(updf.page_count)))
            self.assertIn("动也文化VS钡栩传媒", full)
            self.assertIn("68胜——69胜", full)
            self.assertIn("A组第一名VSB组第二名", full)
            self.assertNotIn("错误队", full)

    def test_cover_art_generation_and_fastapi_routes(self) -> None:
        cov_png = self.root / "cov.png"
        cover_art.generate_cover_from_background(None, cov_png, "2026年测试篮球比赛", footer_lines=["组委会"])
        self.assertTrue(cov_png.is_file())
        with Image.open(cov_png) as im:
            self.assertEqual(im.size, cover_art.A4_PORTRAIT_PX)

        client = TestClient(web_app.app)
        resp = client.get("/api/templates")
        self.assertEqual(resp.status_code, 200)
        self.assertGreaterEqual(len(resp.json()["templates"]), 4)

        sample = self._make_sample_book(self.root / "api_sample.docx")
        with sample.open("rb") as fh:
            up_tpl = client.post(
                "/api/templates/parse",
                files={"file": ("sample.docx", fh, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={"name": "API入库模板", "use_llm": "false"},
            )
        self.assertEqual(up_tpl.status_code, 201)
        tpl_id = up_tpl.json()["id"]

        p_resp = client.post("/api/projects", data={"name": "测试联赛", "template_id": tpl_id})
        self.assertEqual(p_resp.status_code, 201)
        pid = p_resp.json()["id"]

        reg_docx = self._make_source_docx(self.root / "api_reg.docx")
        with reg_docx.open("rb") as fh:
            m_resp = client.post(
                f"/api/projects/{pid}/materials",
                files=[("files", ("竞赛规程.docx", fh, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))],
                data={"role_hint": "auto"},
            )
        self.assertEqual(m_resp.status_code, 201)
        mid = m_resp.json()["added"][0]["id"]

        # Replace & remove material
        with reg_docx.open("rb") as fh:
            rep = client.post(
                f"/api/projects/{pid}/materials",
                files=[("files", ("竞赛规程v2.docx", fh, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))],
                data={"replace_material_id": mid},
            )
        self.assertEqual(rep.status_code, 201)
        self.assertEqual(len(rep.json()["project"]["materials"]), 1)
        self.assertEqual(rep.json()["project"]["materials"][0]["name"], "竞赛规程v2.docx")

        del_m = client.delete(f"/api/projects/{pid}/materials/{mid}")
        self.assertEqual(del_m.status_code, 200)
        self.assertEqual(len(del_m.json()["project"]["materials"]), 0)

        # Cover preview endpoint
        c_resp = client.post("/api/cover/preview", data={"event_title": "封面生成测试", "footer": "测试单位"})
        self.assertEqual(c_resp.status_code, 200)
        img_resp = client.get(c_resp.json()["preview_url"])
        self.assertEqual(img_resp.status_code, 200)
        self.assertEqual(img_resp.headers["content-type"], "image/png")

        # Cleanup temp cover previews created inside root
        shutil.rmtree(self.root / "cover_previews", ignore_errors=True)

    def test_book_mcp_stdio_and_project_trace(self) -> None:
        import asyncio
        import json
        import os
        import sys
        from mcp import ClientSession, StdioServerParameters
        from mcp.client.stdio import stdio_client
        import documents

        proj = projects.create_project("Trace与MCP测试", "official-inline")
        projects.append_project_trace(proj["id"], {"kind": "step", "detail": "hello"})
        traces = projects.read_project_trace(proj["id"])
        self.assertEqual(len(traces), 1)
        self.assertEqual(traces[0]["detail"], "hello")

        doc_path = self._make_source_docx(self.root / "mcp_src.docx")
        d = Document(str(doc_path))
        t = d.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "场次"
        t.cell(0, 1).text = "对阵"
        t.cell(1, 0).text = "1"
        t.cell(1, 1).text = "A1 VS A2"
        d.save(str(doc_path))

        data_dir = self.root / "mcp_data"
        tdir = data_dir / "tasks" / "mcp_t1"
        tdir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(doc_path, tdir / "f01.docx")
        (tdir / "state.json").write_text(
            json.dumps(
                {
                    "id": "mcp_t1",
                    "files": [
                        {
                            "id": "F1",
                            "name": "赛程原件.docx",
                            "stored_name": "f01.docx",
                            "suffix": ".docx",
                            "size": (tdir / "f01.docx").stat().st_size,
                            "units": documents.inventory(tdir / "f01.docx", bulk=True),
                        }
                    ],
                    "template_spec": {"id": "official-inline", "name": "test"},
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        (tdir / "book_inspect.json").write_text(
            json.dumps(inspect_target_book(doc_path), ensure_ascii=False),
            encoding="utf-8",
        )

        async def _verify() -> None:
            env = dict(os.environ, EVENT_BOOK_DATA_DIR=str(data_dir))
            for mode in ("plan", "draw"):
                params = StdioServerParameters(
                    command=sys.executable,
                    args=[str(Path(__file__).resolve().parents[1] / "book_mcp.py"), "--task", "mcp_t1", "--mode", mode],
                    env=env,
                )
                async with stdio_client(params) as (r, w):
                    async with ClientSession(r, w) as sess:
                        await sess.initialize()
                        tools = await sess.list_tools()
                        names = {tl.name for tl in tools.tools}
                        self.assertIn("list_documents", names)
                        self.assertIn("read_document", names)
                        rd = await sess.call_tool("read_document", {"file_id": "F1", "mode": "text"})
                        self.assertFalse(rd.is_error)
                        if mode == "plan":
                            self.assertIn("submit_book_plan", names)
                            ins = await sess.call_tool("inspect_docx_outline", {"file_id": "F1"})
                            self.assertFalse(ins.is_error)
                            self.assertEqual(json.loads(ins.content[0].text)["file_name"], "赛程原件.docx")
                            sub = await sess.call_tool(
                                "submit_book_plan",
                                {
                                    "event_title": "MCP测试赛",
                                    "roster_scope": {"sources": [], "groups": []},
                                    "chapters": [{"id": "c1", "title": "竞赛规程", "kind": "word", "sources": [{"file_id": "F1"}]}],
                                    "missing_items": [],
                                },
                            )
                            self.assertFalse(sub.is_error)
                        else:
                            self.assertIn("submit_draw_patches", names)
                            itb = await sess.call_tool("inspect_target_book", {})
                            self.assertFalse(itb.is_error)
                            sp = await sess.call_tool(
                                "submit_draw_patches",
                                {
                                    "summary": "更新A1 VS A2",
                                    "patches": [
                                        {
                                            "target": "cell",
                                            "table": 0,
                                            "row": 1,
                                            "col": 1,
                                            "find": "A1 VS A2",
                                            "replace": "甲队 VS 乙队",
                                        }
                                    ],
                                },
                            )
                            self.assertFalse(sp.is_error)

        asyncio.run(_verify())

    def test_official_formatting_rules_and_layout(self) -> None:
        from docx.oxml import OxmlElement, parse_xml
        from docx.oxml.ns import nsdecls, qn

        wm = self.root / "wm.png"
        Image.new("RGBA", (300, 300), (200, 200, 200, 40)).save(wm)

        src_docx = self.root / "official_src.docx"
        sdoc = Document()
        sdoc.add_paragraph("各相关单位：")
        sdoc.add_paragraph("现将2026年竞赛有关事项通知如下。")
        sdoc.add_paragraph("一、总体要求与竞赛安排")
        sdoc.add_paragraph("(一)、组织原则")
        sdoc.add_paragraph("1、各参赛代表队必须严格遵守赛风赛纪。")
        sdoc.add_paragraph("(1)，参赛队员须持有二级以上医院体检证明。")
        sdoc.add_paragraph("二、日程安排与成绩统计")
        sdoc.add_paragraph("1、比赛时间地点另行通知。")
        sdoc.add_paragraph("附件：1.关于举办2026年重庆市青少年体育联赛暨校园特色展示活动报名表及汇总清单")
        sdoc.add_paragraph("2、参赛代表队安全承诺书及健康自查声明以及相关证明材料汇总。")
        sdoc.add_paragraph("附件1")
        sdoc.add_paragraph("活动报名表填写说明")
        sdoc.add_paragraph("一、填写规范")
        sdoc.add_paragraph("请各单位按要求逐项填写完整，确保信息真实准确。")
        mixed_p = sdoc.add_paragraph()
        mixed_p.add_run("（本页无需提交）")
        br_r = mixed_p.add_run()
        br_el = OxmlElement("w:br")
        br_el.set(qn("w:type"), "page")
        br_r._r.append(br_el)
        mixed_p.add_run("附件2")
        sdoc.add_paragraph("报名二维码")
        sdoc.add_paragraph("请使用微信扫码填报。")
        sdoc.save(src_docx)

        tpl = templates_store.get_template("official-inline")
        self.assertIsNotNone(tpl)
        out_docx = self.root / "official_book.docx"
        out_pdf = self.root / "official_book.pdf"
        plan = {
            "event_title": "2026年官方格式测试赛",
            "subtitle": "秩序册",
            "chapters": [
                {
                    "title": "一、关于举办2026年重庆市青少年体育联赛的通知——兼论竞赛规程及报名补充要求。",
                    "kind": "richtext",
                    "sources": [{"type": "docx", "path": str(src_docx)}],
                },
                {"title": "记事", "kind": "notes"},
            ],
        }
        res = assemble_book(plan, tpl, out_docx, out_pdf, watermark_override=wm)
        self.assertTrue(out_docx.is_file())
        self.assertTrue(out_pdf.is_file())
        self.assertEqual(len(res["toc"]), 1)
        self.assertEqual(res["toc"][0]["page"], 1)

        # Verify DOCX XML rules
        bdoc = Document(str(out_docx))
        self.assertTrue(bdoc.settings.odd_and_even_pages_header_footer)
        self.assertEqual(bdoc.styles["Normal"].paragraph_format.space_before, Pt(0))
        self.assertEqual(bdoc.styles["Normal"].paragraph_format.space_after, Pt(0))

        body_sec = bdoc.sections[2]  # cover, toc, first body chapter
        self.assertAlmostEqual(body_sec.top_margin.mm, 30.0, places=1)
        self.assertAlmostEqual(body_sec.bottom_margin.mm, 25.0, places=1)
        self.assertAlmostEqual(body_sec.left_margin.mm, 26.0, places=1)
        self.assertAlmostEqual(body_sec.right_margin.mm, 25.0, places=1)
        self.assertIsNotNone(body_sec._sectPr.find(qn("w:docGrid")))
        self.assertTrue(len(body_sec.header._element.findall(".//" + qn("a:blip"))) > 0)
        self.assertTrue(len(body_sec.even_page_header._element.findall(".//" + qn("a:blip"))) > 0)

        p_texts = [p.text for p in bdoc.paragraphs]
        self.assertIn("（一）组织原则", p_texts)
        self.assertIn("1．各参赛代表队必须严格遵守赛风赛纪。", p_texts)
        self.assertIn("（1）参赛队员须持有二级以上医院体检证明。", p_texts)
        self.assertIn("——兼论竞赛规程及报名补充要求", p_texts)
        self.assertNotIn("——兼论竞赛规程及报名补充要求。", p_texts)
        self.assertIn("（本页无需提交）", p_texts)
        self.assertIn("附件2", p_texts)
        self.assertIn("2．参赛代表队安全承诺书及健康自查声明以及相关证明材料汇总", p_texts)

        # Rule 1 & Rule 2 non-bold + keepNext + Rule 3 twip-only indents
        for p in bdoc.paragraphs:
            t = p.text.strip()
            if t in {"一、总体要求与竞赛安排", "（一）组织原则", "1．各参赛代表队必须严格遵守赛风赛纪。", "（1）参赛队员须持有二级以上医院体检证明。", "附件2"}:
                for r in p.runs:
                    self.assertNotEqual(r.bold, True, f"Heading unexpectedly bold: {t}")
            if t == "一、总体要求与竞赛安排":
                self.assertIsNotNone(p._p.find("./w:pPr/w:keepNext", p._p.nsmap))
            if t.startswith("附件：1．") or t.startswith("2．参赛代表队"):
                ind = p._p.find("./w:pPr/w:ind", p._p.nsmap)
                self.assertIsNotNone(ind)
                self.assertIsNone(ind.get(qn("w:leftChars")))
                self.assertIsNone(ind.get(qn("w:hangingChars")))

        # Verify PDF geometry, page numbers, trapezoid title split, and alignment
        pdf = pymupdf.open(str(out_pdf))
        p3 = pdf[2]  # first body page (-1-)
        p4 = pdf[3]  # 附件1 (-2-)
        p5 = pdf[4]  # 附件2 (-3-)
        self.assertIn("-1-", p3.get_text("text"))
        self.assertIn("-2-", p4.get_text("text"))
        self.assertIn("-3-", p5.get_text("text"))
        self.assertNotIn("（本页无需提交）", p5.get_text("text"))
        self.assertIn("报名二维码", p5.get_text("text"))

        # Check odd-right vs even-left footer positions and non-split 2026年
        p3_words = p3.get_text("words")
        p4_words = p4.get_text("words")
        fn1 = [w for w in p3_words if "-1-" in w[4]][0]
        fn2 = [w for w in p4_words if "-2-" in w[4]][0]
        self.assertGreater(fn1[0], 450.0)
        self.assertLess(fn2[0], 130.0)
        p3_lines = [l.strip() for l in p3.get_text("text").splitlines() if l.strip()]
        self.assertTrue(any("2026" in ln and "年" in ln for ln in p3_lines[:5]))

        # Check multi-item attachment description hanging alignment on p3
        att_x: dict[str, float] = {}
        for b in p3.get_text("rawdict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                chars = [c for sp in l["spans"] for c in sp.get("chars", [])]
                s = "".join(c["c"] for c in chars).strip()
                if s.startswith("附件：1．"):
                    att_x["att"] = next(c for c in chars if c["c"] == "附")["bbox"][0]
                    att_x["item1_text"] = next(c for c in chars if c["c"] == "关")["bbox"][0]
                elif s.startswith("色展示"):
                    att_x["wrap1"] = next(c for c in chars if c["c"] == "色")["bbox"][0]
                elif s.startswith("2．参赛"):
                    att_x["num2"] = next(c for c in chars if c["c"] == "2")["bbox"][0]
                    att_x["item2_text"] = next(c for c in chars if c["c"] == "参")["bbox"][0]
                elif "材料汇总" in s:
                    att_x["wrap2"] = next(c for c in chars if c["c"].strip())["bbox"][0]
        self.assertEqual(set(att_x), {"att", "item1_text", "wrap1", "num2", "item2_text", "wrap2"})
        self.assertLess(att_x["att"], att_x["num2"])
        self.assertAlmostEqual(att_x["wrap1"], att_x["item1_text"], delta=2.0)
        self.assertAlmostEqual(att_x["wrap2"], att_x["item2_text"], delta=2.0)
        pdf.close()

    def test_template_parser_handles_start_jc_and_official_footers(self) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls, qn
        from docx.shared import Mm

        sample = self.root / "official_sample.docx"
        doc = Document()
        doc.settings.odd_and_even_pages_header_footer = True
        sec = doc.sections[0]
        sec.top_margin = Mm(30)
        sec.bottom_margin = Mm(25)
        sec.left_margin = Mm(26)
        sec.right_margin = Mm(25)
        odd_p = sec.footer.paragraphs[0]
        odd_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        odd_p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="28"/></w:rPr><w:t>-</w:t></w:r>'))
        odd_p._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr=" PAGE "/>'))
        odd_p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="28"/></w:rPr><w:t>-</w:t></w:r>'))
        even_p = sec.even_page_footer.paragraphs[0]
        even_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        even_p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="28"/></w:rPr><w:t>-</w:t></w:r>'))
        even_p._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr=" PAGE "/>'))
        even_p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:sz w:val="28"/></w:rPr><w:t>-</w:t></w:r>'))

        p = doc.add_paragraph("测试起跑对齐段落")
        p_pr = p._p.get_or_add_pPr()
        p_pr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="start"/>'))
        for i in range(4):
            bp = doc.add_paragraph(f"一、章节{i + 1}")
            for r in bp.runs:
                r.font.name = "方正黑体_GBK"
                r.font.size = Pt(16)
        doc.save(sample)

        parsed = template_parser.parse_sample_book_to_template(sample, name="公文样例", use_llm=False)
        fmt = parsed["formatting"]
        margins = fmt["page"]["margins_mm"]
        self.assertAlmostEqual(margins["top"], 30.0, delta=0.5)
        self.assertAlmostEqual(margins["bottom"], 25.0, delta=0.5)
        self.assertAlmostEqual(margins["left"], 26.0, delta=0.5)
        self.assertAlmostEqual(margins["right"], 25.0, delta=0.5)
        self.assertEqual(fmt["page_numbers"]["style"], "hyphen")
        self.assertEqual(fmt["page_numbers"]["prefix"], "-")
        self.assertEqual(fmt["page_numbers"]["suffix"], "-")
        self.assertEqual(fmt["page_numbers"]["align"], "odd_right_even_left")
        self.assertTrue(fmt["page_numbers"]["odd_even"])

