"""assemble 的測試（unittest，全部使用合成資料，不碰網路）。

直接跑：
    uv run --with python-pptx --with pymupdf --with google-api-python-client python tools/tests/test_assemble.py
"""
import importlib.util
import io
import sys
import tempfile
import unittest
import warnings
from pathlib import Path

import pymupdf
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

_TOOLS = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_TOOLS))
for _name in ("bni25s", "assemble"):
    _spec = importlib.util.spec_from_file_location(_name, _TOOLS / f"{_name}.py")
    _mod = importlib.util.module_from_spec(_spec)
    sys.modules[_name] = _mod
    _spec.loader.exec_module(_mod)
assemble = sys.modules["assemble"]

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def S(*lines, notes=""):
    return (lines, notes)


def build(spec, path: Path, sections=True) -> Path:
    """spec = [(章節名, [S(...), ...])]；sections=False 就不寫章節清單。"""
    prs = Presentation()
    groups, idx = [], 0
    for name, slides in spec:
        idxs = []
        for lines, notes in slides:
            s = prs.slides.add_slide(prs.slide_layouts[6])
            for i, line in enumerate(lines):
                s.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(4), Inches(1)).text_frame.text = line
            if notes:
                s.notes_slide.notes_text_frame.text = notes
            idxs.append(idx)
            idx += 1
        groups.append((name, idxs))
    if sections:
        pres = prs.part._element
        ids = [e.get("id") for e in prs.slides._sldIdLst.sldId_lst]
        ext_lst = pres.find(qn("p:extLst"))
        if ext_lst is None:
            ext_lst = etree.SubElement(pres, qn("p:extLst"))
        ext = etree.SubElement(ext_lst, qn("p:ext"), uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
        sec_lst = etree.SubElement(ext, f"{{{P14}}}sectionLst")
        for n, (name, idxs) in enumerate(groups):
            sec = etree.SubElement(sec_lst, f"{{{P14}}}section", name=name, id=f"{{00000000-0000-0000-0000-{n:012d}}}")
            lst = etree.SubElement(sec, f"{{{P14}}}sldIdLst")
            for i in idxs:
                etree.SubElement(lst, f"{{{P14}}}sldId", id=ids[i])
    prs.save(str(path))
    return path


def guest_deck(path: Path, *names) -> Path:
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rows = max(1, len(names))
    table = s.shapes.add_table(rows, 6, Inches(0.5), Inches(1), Inches(9), Inches(1)).table
    for r in range(rows):
        table.cell(r, 0).text = f"來賓{r + 1}"
        table.cell(r, 1).text = names[r] if names else ""
    prs.save(str(path))
    return path


MASTER = [
    ("BNI簡介", [S("BNI 簡介", notes="#固定"), S("4,195 會員", notes="#BNI台灣發展現況")]),
    ("表揚優秀會員", [S("表揚優秀會員", notes="#固定"), S("舊表揚")]),
    ("金質獎章會員", [S("舊金質")]),
    ("副主席報告-YYYYMMDD副主席", [S("副主席報告"), S("舊副主席內容")]),
    ("教育培訓時間-3分鐘教育", [S("教育培訓時間", notes="教育培訓Header #固定"), S("舊教育1"), S("舊教育2")]),
    ("活動協調報告", [S("活動協調報告", notes="#固定"), S("舊活動")]),
    ("夥伴活動分享", [S("夥伴活動分享", notes="#固定"), S("舊夥伴")]),
    ("來賓20秒自我介紹", [S("來賓介紹", notes="#固定"), S("許願池", notes="#固定"), S("舊來賓名單")]),
    ("25秒簡報-甲組", [S("甲組 請上台", notes="甲組 #固定"), S("甲組 組長：王", "230 王小明", "115 李大華"),
                   S("230", "王小明"), S("115 李大華")]),
    ("25秒簡報-乙組", [S("乙組 請上台", notes="乙組 #固定"), S("乙組 組長：簡", "206 簡均安"), S("206 簡均安")]),
    ("25秒簡報-丙組", [S("丙組 請上台", notes="丙組 #固定"), S("丙組 組長：林", "301 林"), S("301 林")]),
]


def first_lines(prs) -> dict[str, list[str]]:
    by_id = {str(s.slide_id): s for s in prs.slides}
    out = {}
    for sec in prs.part._element.iter(f"{{{P14}}}section"):
        out[sec.get("name")] = [assemble.slide_text(by_id[e.get("id")]).split("\n")[0]
                                for e in sec.iter(f"{{{P14}}}sldId")]
    return out


class Case(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.up = self.tmp / "up"
        self.up.mkdir()
        self.mem = self.tmp / "mem"
        self.mem.mkdir()
        self.master = build(MASTER, self.tmp / "master.pptx")
        self.guest = guest_deck(self.tmp / "guest.pptx", "陳庭儀")

    def run_assemble(self, guest=None):
        out = self.tmp / "out.pptx"
        uploads = sorted(p for p in self.up.iterdir() if p.suffix in (".pptx", ".pdf"))
        deck, rep = assemble.assemble(self.master, uploads, assemble.member_files(self.mem),
                                      guest or self.guest, out)
        with warnings.catch_warnings():
            warnings.simplefilter("error")  # zipfile Duplicate name 警告＝壞檔
            prs = Presentation(str(out))
        return deck, rep, prs


class ClassifyTest(unittest.TestCase):
    def test_names(self):
        self.assertEqual(assemble.classify("副主席20261028.pptx"), ("role", "副主席"))
        self.assertEqual(assemble.classify("3分鐘20261028.pptx"), ("role", "三分鐘"))
        self.assertEqual(assemble.classify("秘財20261028 (1).pptx"), ("role", "秘財"))
        self.assertEqual(assemble.classify("25秒-女性魅力經濟20261028.pptx"), ("group", "女性魅力經濟"))
        self.assertEqual(assemble.classify("會員_230.pptx"), (None, None))

    def test_date_first(self):
        self.assertEqual(assemble.classify("20261028副主席.pptx"), ("role", "副主席"))
        self.assertEqual(assemble.classify("20261028_25秒-乙組.pptx"), ("group", "乙組"))
        self.assertEqual(assemble.classify("秘財1151028.pptx"), ("role", "秘財"))

    def test_match_section_with_suffix(self):
        owned = assemble.ROLES["副主席"]
        self.assertEqual(assemble.match_section("副主席報告-20261028副主席", owned), "副主席報告")
        self.assertIsNone(assemble.match_section("秘書財務報告", owned))


class SelectUploadsTest(unittest.TestCase):
    def pick(self, names, master="20261021.pptx"):
        rep = assemble.Report()
        got = assemble.select_uploads([Path(n) for n in names], Path(master), rep)
        return [p.name for p in got], rep

    def test_file_date(self):
        self.assertEqual(assemble.file_date("秘財1151028.pptx"), "20261028")
        self.assertEqual(assemble.file_date("20261028副主席.pptx"), "20261028")
        self.assertEqual(assemble.file_date("25秒-乙組20261028 (1).pptx"), "20261028")
        self.assertIsNone(assemble.file_date("活動組.pptx"))

    def test_latest_per_role_and_older_than_master_dropped(self):
        got, rep = self.pick(["三分鐘20261021.pptx", "三分鐘20261026.pptx", "秘財20261021.pptx",
                              "秘財1151028.pptx", "25秒-乙組20261014.pptx", "副主席20261028.pptx"])
        self.assertEqual(got, ["三分鐘20261026.pptx", "副主席20261028.pptx", "秘財1151028.pptx"])
        self.assertEqual(rep.problems, [])

    def test_newer_wins_without_master_date(self):
        got, _ = self.pick(["三分鐘20261021.pptx", "3分鐘20261028.pptx"], master="m.pptx")
        self.assertEqual(got, ["3分鐘20261028.pptx"])

    def test_undated_and_same_day_reported(self):
        got, rep = self.pick(["活動組.pptx", "副主席20261028.pptx", "20261028副主席.pptx"])
        self.assertEqual(got, [])
        self.assertTrue(any("活動組.pptx" in p and "沒有日期" in p for p in rep.problems))
        self.assertTrue(any("同一天有多份" in p for p in rep.problems))


class MemberNumberTest(unittest.TestCase):
    def slide(self, *lines):
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        for i, line in enumerate(lines):
            s.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(4), Inches(1)).text_frame.text = line
        return s

    def test_single_number(self):
        self.assertEqual(assemble.member_number(self.slide("230", "王小明")), "230")

    def test_lowercase_no_prefix(self):
        self.assertEqual(assemble.member_number(self.slide("no.230 王小明")), "230")

    def test_overview_with_many_numbers_is_not_member(self):
        self.assertIsNone(assemble.member_number(self.slide("230 王小明", "115 李大華")))

    def test_leader_own_page_is_member(self):
        self.assertEqual(assemble.member_number(self.slide("120 吳昇洋", "本組組長，服務 16 年")), "120")

    def test_single_member_overview_is_not_member(self):
        self.assertIsNone(assemble.member_number(self.slide("丙組 組長：林", "301 林")))

    def test_tag_wins(self):
        s = self.slide("只有圖片")
        s._element.cSld.set("name", "BNI25S-230")
        self.assertEqual(assemble.member_number(s), "230")


class RoleTest(Case):
    def test_vice_president(self):
        build([("BNI簡介", [S("4,300 會員", notes="#BNI台灣發展現況")]),
               ("表揚優秀會員", [S("表揚優秀會員"), S("新表揚")]),
               ("副主席報告-20261028副主席", [S("副主席報告"), S("新副主席")])],
              self.up / "副主席20261028.pptx")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["BNI簡介"], ["BNI 簡介", "4,300 會員"])
        self.assertEqual(got["表揚優秀會員"], ["表揚優秀會員", "新表揚"])
        self.assertEqual(got["副主席報告-YYYYMMDD副主席"], ["副主席報告", "新副主席"])
        self.assertEqual(got["金質獎章會員"], ["舊金質"])
        self.assertTrue(any("沒有章節「金質獎章會員」" in p for p in rep.problems))
        # 主檔副主席報告標題頁沒標 #固定 → 整段換並提醒
        self.assertTrue(any("副主席報告" in p and "#固定" in p for p in rep.problems))

    def test_vice_president_header_kept_when_upload_lacks_it(self):
        build([("表揚優秀會員", [S("新表揚A"), S("新表揚B")]), ("金質獎章會員", [S("新金質")])],
              self.up / "副主席20261028.pptx")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["表揚優秀會員"], ["表揚優秀會員", "新表揚A", "新表揚B"])
        self.assertEqual(got["金質獎章會員"], ["新金質"])
        self.assertFalse(any("金質獎章會員" in p and "#固定" in p for p in rep.problems))
        head = next(s for s in prs.slides if assemble.slide_text(s) == "表揚優秀會員")
        self.assertIn("#固定", head.notes_slide.notes_text_frame.text)

    def test_empty_section_kept_empty(self):
        build([("活動協調報告", []), ("夥伴活動分享", [S("新夥伴")])], self.up / "活動組20261028.pptx")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["活動協調報告"], [])
        self.assertEqual(got["夥伴活動分享"], ["新夥伴"])

    def test_no_sections_skipped(self):
        build([("x", [S("新專講")])], self.up / "秘財20261028.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        self.assertTrue(any("沒有任何章節" in p for p in rep.problems))

    def test_three_minute_without_sections(self):
        build([("x", [S("新教育1"), S("新教育2"), S("新教育3")])], self.up / "三分鐘20261028.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["教育培訓時間-3分鐘教育"], ["教育培訓時間", "新教育1", "新教育2", "新教育3"])
        self.assertFalse(rep.problems and any("三分鐘" in p for p in rep.problems))

    def test_events_without_sections_split_by_title(self):
        build([("x", [S("活動協調報告"), S("新活動1"), S("新活動2"), S("夥伴活動分享"), S("新夥伴1")])],
              self.up / "活動組20261028.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["活動協調報告"], ["活動協調報告", "新活動1", "新活動2"])
        self.assertEqual(got["夥伴活動分享"], ["夥伴活動分享", "新夥伴1"])
        heads = [s for s in prs.slides if assemble.slide_text(s) in ("活動協調報告", "夥伴活動分享")]
        self.assertEqual(len(heads), 2)  # 標題頁用主檔的，上傳檔的不重複放
        self.assertTrue(all("#固定" in s.notes_slide.notes_text_frame.text for s in heads))

    def test_events_without_sections_missing_title_skipped(self):
        build([("x", [S("活動協調報告"), S("新活動1")])], self.up / "活動組20261028.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["活動協調報告"], ["活動協調報告", "舊活動"])
        self.assertEqual(got["夥伴活動分享"], ["夥伴活動分享", "舊夥伴"])
        self.assertTrue(any("夥伴活動分享" in p and "找到 0 頁" in p for p in rep.problems))

    def test_notes_carried(self):
        build([("副主席報告", [S("副主席報告"), S("新內容", notes="備註要照抄")])], self.up / "副主席20261028.pptx")
        _, _, prs = self.run_assemble()
        notes = [s.notes_slide.notes_text_frame.text for s in prs.slides if s.has_notes_slide]
        self.assertIn("備註要照抄", notes)


def scaled_upload(path: Path, k_w: float, k_h: float) -> Path:
    """預設尺寸 ×(k_w, k_h) 的三分鐘檔：一個文字框在 (1.5in, 1.5in)，字級 30pt。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = int(prs.slide_width * k_w), int(prs.slide_height * k_h)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(Inches(1.5), Inches(1.5), Inches(3), Inches(1.5))
    tb.text_frame.text = "Canva 教育"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(30)
    prs.save(str(path))
    return path


def pdf_upload(path: Path, *pages, size=(720, 540)) -> Path:
    """每頁一行文字的 PDF（預設 10x7.5 吋，和測試主檔同比例）。"""
    doc = pymupdf.open()
    for text in pages:
        doc.new_page(width=size[0], height=size[1]).insert_text((72, 72), text, fontname="china-t", fontsize=24)
    doc.save(str(path))
    return path


def pictures(slide) -> int:
    return sum(1 for sh in slide.shapes if sh.shape_type == 13)


class PdfTest(Case):
    def test_three_minute_pdf(self):
        pdf_upload(self.up / "三分鐘20261028.pdf", "教育一", "教育二")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)["教育培訓時間-3分鐘教育"]
        self.assertEqual(got, ["教育培訓時間", "", ""])
        imgs = [s for s in prs.slides if assemble.slide_text(s) == "" and pictures(s) == 1]
        self.assertEqual(len(imgs), 2)
        self.assertTrue(any("PDF 2 頁轉成圖片投影片" in line for line in rep.lines))
        self.assertEqual(rep.problems, [])

    def test_events_pdf_split_by_pdf_text(self):
        pdf_upload(self.up / "活動組20261028.pdf", "活動協調報告", "新活動", "夥伴活動分享", "新夥伴1", "新夥伴2")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["活動協調報告"], ["活動協調報告", ""])
        self.assertEqual(got["夥伴活動分享"], ["夥伴活動分享", "", ""])

    def test_group_pdf(self):
        pdf_upload(self.up / "25秒-乙組20261028.pdf", "乙組 組長：簡", "乙組 目標客群")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)["25秒簡報-乙組"]
        self.assertEqual(got, ["乙組 請上台", "", "", "206 簡均安"])

    def test_pdf_fits_master_keeping_ratio(self):
        pdf_upload(self.up / "三分鐘20261028.pdf", "寬版", size=(960, 540))  # 16:9 放進 4:3 主檔
        _, _, prs = self.run_assemble()
        pic = next(sh for s in prs.slides for sh in s.shapes if sh.shape_type == 13)
        self.assertEqual(pic.width, prs.slide_width)
        self.assertAlmostEqual(pic.height / pic.width, 540 / 960, places=3)
        self.assertEqual(pic.top, (prs.slide_height - pic.height) // 2)


class SizeTest(Case):
    def test_same_ratio_scaled(self):
        scaled_upload(self.up / "三分鐘20261028.pptx", 1.5, 1.5)
        _, rep, prs = self.run_assemble()
        s = next(s for s in prs.slides if assemble.slide_text(s) == "Canva 教育")
        tb = s.shapes[0]
        self.assertEqual((tb.left, tb.top, tb.width, tb.height), (Inches(1), Inches(1), Inches(2), Inches(1)))
        self.assertEqual(tb.text_frame.paragraphs[0].runs[0].font.size, Pt(20))
        self.assertTrue(any("等比例縮放" in line for line in rep.lines))

    def test_different_ratio_skipped(self):
        scaled_upload(self.up / "三分鐘20261028.pptx", 1.5, 1.0)
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["教育培訓時間-3分鐘教育"], ["教育培訓時間", "舊教育1", "舊教育2"])
        self.assertTrue(any("比例和主檔不同" in p for p in rep.problems))

    def test_same_size_untouched(self):
        self.assertIsNone(assemble.fit_to_master(Presentation(), Presentation().slide_width,
                                                 Presentation().slide_height))


class GroupTest(Case):
    def test_group_pages_replaced_stage_kept(self):
        build([("x", [S("乙組 請上台"), S("乙組 組長：簡", "206 簡均安"), S("乙組 目標客群")])],
              self.up / "25秒-乙組20261028.pptx")
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["25秒簡報-乙組"], ["乙組 請上台", "乙組 組長：簡", "乙組 目標客群", "206 簡均安"])

    def test_imported_group_pages_marked_fixed(self):
        build([("x", [S("乙組 組長：簡", "206 簡均安")])], self.up / "25秒-乙組20261028.pptx")
        _, _, prs = self.run_assemble()
        s = next(s for s in prs.slides if assemble.slide_text(s).startswith("乙組 組長"))
        self.assertIn("#固定", s.notes_slide.notes_text_frame.text)


class MemberTest(Case):
    def test_replace_by_slide_number(self):
        build([("x", [S("no.230 新內容")])], self.mem / "王小明.pptx", sections=False)
        deck, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["25秒簡報-甲組"], ["甲組 請上台", "甲組 組長：王", "no.230 新內容", "115 李大華"])
        self.assertEqual(list(deck.member_sid), ["230"])
        s = next(s for s in prs.slides if assemble.slide_text(s).startswith("no.230"))
        self.assertEqual(s._element.cSld.get("name"), "BNI25S-230")

    def test_duplicate_number_rejected(self):
        build([("x", [S("230 甲")])], self.mem / "a_230.pptx", sections=False)
        build([("x", [S("230 乙")])], self.mem / "b_230.pptx", sections=False)
        deck, rep, _ = self.run_assemble()
        self.assertEqual(deck.member_sid, {})
        self.assertTrue(any("同一編號" in p for p in rep.problems))

    def test_problem_message_has_no_filename(self):
        build([("x", [S("只有圖片")])], self.mem / "王小明.pptx", sections=False)
        _, rep, _ = self.run_assemble()
        self.assertTrue(rep.problems)
        self.assertFalse(any("王小明" in p for p in rep.problems))


class MoveTest(Case):
    def test_move_by_roster(self):
        build([("x", [S("甲組 組長：王", "230 王小明")])], self.up / "25秒-甲組20261028.pptx")
        build([("x", [S("乙組 組長：簡", "206 簡均安", "115 李大華")])], self.up / "25秒-乙組20261028.pptx")
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        self.assertEqual(got["25秒簡報-甲組"][-1], "230")
        self.assertEqual(got["25秒簡報-乙組"][-2:], ["206 簡均安", "115 李大華"])  # 第一位 206 不動
        self.assertEqual(rep.moves, ["  115：甲組 → 乙組"])

    def test_no_move_when_still_listed(self):
        build([("x", [S("乙組 組長：簡", "206 簡均安", "115 李大華")])], self.up / "25秒-乙組20261028.pptx")
        _, rep, _ = self.run_assemble()
        self.assertEqual(rep.moves, [])

    def test_ambiguous_not_moved(self):
        build([("x", [S("甲組 組長：王", "230 王小明")])], self.up / "25秒-甲組20261028.pptx")
        build([("x", [S("乙組 組長：簡", "206 簡均安", "115 李大華")])], self.up / "25秒-乙組20261028.pptx")
        build([("x", [S("丙組 組長：林", "301 林", "115 李大華")])], self.up / "25秒-丙組20261028.pptx")
        _, rep, _ = self.run_assemble()
        self.assertEqual(rep.moves, [])
        self.assertTrue(any(p.startswith("115：") for p in rep.problems))


class GuestTest(Case):
    def test_guests_shown(self):
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)["來賓20秒自我介紹"]
        self.assertEqual(got[:2], ["來賓介紹", "許願池"])
        self.assertEqual(len(got), 3)
        self.assertTrue(all(s._element.get("show") is None for s in list(prs.slides)[:3]))

    def test_no_guest_hides_first_two(self):
        _, _, prs = self.run_assemble(guest_deck(self.tmp / "g0.pptx"))
        hidden = [assemble.slide_text(s).split("\n")[0] for s in prs.slides if s._element.get("show") == "0"]
        self.assertEqual(hidden, ["來賓介紹", "許願池"])


class RotateTest(Case):
    def test_first_group_moves_last(self):
        _, rep, prs = self.run_assemble()
        groups = [n for n in first_lines(prs) if n.startswith("25秒簡報-")]
        self.assertEqual(groups, ["25秒簡報-乙組", "25秒簡報-丙組", "25秒簡報-甲組"])

    def test_slide_order_follows_sections(self):
        _, _, prs = self.run_assemble()
        flat = [line for lines in first_lines(prs).values() for line in lines]
        self.assertEqual(flat, [assemble.slide_text(s).split("\n")[0] for s in prs.slides])


if __name__ == "__main__":
    unittest.main()
