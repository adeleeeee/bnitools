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
    ("新會員/續約會員", [S("舊新會員"), S("續約宣示", notes="#續約會員")]),
    ("5分鐘主題專講(秘財)", [S("五分鐘主題專講"), S("舊專講")]),
    ("秘書財務報告", [S("秘書財務報告"), S("舊財務")]),
    ("來賓20秒自我介紹", [S("來賓介紹", notes="#固定"), S("許願池", notes="#固定"), S("舊來賓名單")]),
    ("來賓感想", [S("來賓感想頁", notes="#固定")]),
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

    def test_same_week_keeps_files_dated_as_master(self):
        names = ["秘財20261021.pptx", "秘財1151028.pptx", "三分鐘20261028.pptx"]
        got, _ = self.pick(names, master="20261028.pptx")
        self.assertEqual(got, [])  # 一般組裝：和主檔同一天＝已組過
        rep = assemble.Report()
        got = assemble.select_uploads([Path(n) for n in names], Path("20261028.pptx"), rep, same_week=True)
        self.assertEqual([p.name for p in got], ["三分鐘20261028.pptx", "秘財1151028.pptx"])
        self.assertTrue(any("秘財20261021.pptx" in line and "早於" in line for line in rep.lines))

    def test_undated_and_same_day_reported(self):
        got, rep = self.pick(["活動組.pptx", "副主席20261028.pptx", "20261028副主席.pptx"])
        self.assertEqual(got, [])
        self.assertTrue(any("活動組.pptx" in p and "沒有日期" in p for p in rep.problems))
        self.assertTrue(any("同一天有多份" in p for p in rep.problems))


class DriveHelpersTest(unittest.TestCase):
    PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    def f(self, name, t):
        return {"name": name, "modifiedTime": t, "mimeType": self.PPTX}

    def test_pick_master_latest_date_then_modified(self):
        files = [self.f("20261021.pptx", "2026-10-22T05:51:00Z"),
                 self.f("20261028_秘財更新.pptx", "2026-10-27T14:17:00Z"),
                 self.f("20261028.pptx", "2026-10-27T16:02:00Z"),
                 self.f("報告.pptx", "2026-11-01T00:00:00Z")]
        f, date = assemble.pick_master(files)
        self.assertEqual((f["name"], date), ("20261028.pptx", "20261028"))
        files[1]["modifiedTime"] = "2026-10-28T01:00:00Z"
        self.assertEqual(assemble.pick_master(files)[0]["name"], "20261028_秘財更新.pptx")

    def test_is_same_week(self):
        from datetime import date
        self.assertTrue(assemble.is_same_week("20261111", date(2026, 11, 10)))   # 週二，本週三的檔已經在了
        self.assertTrue(assemble.is_same_week("20261111", date(2026, 11, 11)))   # 例會當天
        self.assertFalse(assemble.is_same_week("20261104", date(2026, 11, 10)))  # 上週的檔＝一般組裝
        self.assertFalse(assemble.is_same_week("20261111", date(2026, 11, 12)))  # 開完會，下一次是下週

    def test_uploaded_at_uses_later_of_created_and_modified(self):
        # 從電腦上傳：modifiedTime 是對方電腦上的修改時間，比上傳時間早
        f = {"name": "a.pptx", "createdTime": "2026-11-09T04:40:02.396Z", "modifiedTime": "2026-11-03T07:18:03Z"}
        self.assertEqual(assemble.newer_than([f], "2026-11-08T00:00:00Z"), [f])
        self.assertEqual(assemble.newer_than([f], "2026-11-09T05:00:00Z"), [])
        self.assertEqual(assemble.newer_than([f], None), [f])
        # 同一秒內，有沒有毫秒都要比對
        g = {"name": "b.pptx", "modifiedTime": "2026-11-10T13:21:25.643Z"}
        self.assertEqual(assemble.newer_than([g], "2026-11-10T13:21:25Z"), [g])

    def test_latest_by_name_keeps_last_upload(self):
        old = {"name": "秘財.pptx", "createdTime": "2026-11-10T07:38:00Z", "modifiedTime": "2026-11-10T07:38:00Z", "id": "old"}
        new = {"name": "秘財.pptx", "createdTime": "2026-11-10T14:00:00Z", "modifiedTime": "2026-11-10T09:00:00Z", "id": "new"}
        other = {"name": "活動組.pptx", "modifiedTime": "2026-11-10T01:00:00Z", "id": "x"}
        for order in ([old, new, other], [new, old, other]):
            kept, dup = assemble.latest_by_name(order)
            self.assertEqual(sorted(f["id"] for f in kept), ["new", "x"])
            self.assertEqual(dup, ["秘財.pptx"])

    def test_pick_master_none(self):
        with self.assertRaises(SystemExit):
            assemble.pick_master([self.f("報告.pptx", "2026-11-01T00:00:00Z")])

    def test_folder_id_from(self):
        fid = "A" * 33
        self.assertEqual(assemble.folder_id_from(f"https://drive.google.com/drive/folders/{fid}?usp=sharing"), fid)
        self.assertEqual(assemble.folder_id_from(fid), fid)
        with self.assertRaises(SystemExit):
            assemble.folder_id_from("https://example.com/x")


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
        build([("x", [S("新副主席")])], self.up / "副主席20261028.pptx", sections=False)
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

    def renewal(self, prs):
        return [(assemble.slide_text(s), s._element.get("show")) for s in prs.slides
                if "續約宣示" in assemble.slide_text(s) or "新會員" in assemble.slide_text(s)]

    def test_finance_without_sections_split_by_title(self):
        build([("x", [S("五分鐘主題專講 Weekly Themed Talks"), S("新專講1"), S("新專講2"),
                      S("秘書財務報告 Secretary"), S("新財務")])], self.up / "秘財1151104.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)
        # 主檔標題頁沒標 #固定 → 連標題頁一起換成上傳檔的
        self.assertEqual(got["5分鐘主題專講(秘財)"], ["五分鐘主題專講 Weekly Themed Talks", "新專講1", "新專講2"])
        self.assertEqual(got["秘書財務報告"], ["秘書財務報告 Secretary", "新財務"])
        # 沒有新會員／續約＝本週沒有：只留宣示頁並隱藏，不算問題
        self.assertEqual(got["新會員/續約會員"], ["續約宣示"])
        self.assertEqual(self.renewal(prs), [("續約宣示", "0")])
        self.assertFalse(any("新會員" in p for p in rep.problems))

    def test_finance_sections_without_new_members(self):
        build([("5分鐘主題專講(秘財)", [S("五分鐘主題專講"), S("新專講")]), ("秘書財務報告", [S("秘書財務報告")])],
              self.up / "秘財20261104.pptx")
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["新會員/續約會員"], ["續約宣示"])
        self.assertEqual(self.renewal(prs), [("續約宣示", "0")])
        self.assertEqual(rep.problems, [])

    def test_finance_empty_new_members_section(self):
        build([("新會員/續約會員", []), ("5分鐘主題專講(秘財)", [S("新專講")]), ("秘書財務報告", [S("新財務")])],
              self.up / "秘財20261104.pptx")
        _, _, prs = self.run_assemble()
        self.assertEqual(self.renewal(prs), [("續約宣示", "0")])

    def test_finance_with_new_members_replaced(self):
        build([("新會員/續約會員", [S("新會員A")]), ("5分鐘主題專講(秘財)", [S("新專講")]), ("秘書財務報告", [S("新財務")])],
              self.up / "秘財20261104.pptx")
        _, _, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["新會員/續約會員"], ["新會員A"])

    def test_empty_upload_reported_and_others_still_applied(self):
        (self.up / "秘財20261028.pptx").write_bytes(b"")
        build([("教育培訓時間-3分鐘教育", [S("新三分鐘")])], self.up / "三分鐘20261028.pptx")
        _, rep, prs = self.run_assemble()
        self.assertTrue(any("秘財20261028.pptx" in p and "0 bytes" in p for p in rep.problems))
        self.assertIn("新三分鐘", first_lines(prs)["教育培訓時間-3分鐘教育"])

    def test_corrupt_upload_reported(self):
        (self.up / "秘財20261028.pptx").write_bytes(b"not a zip")
        _, rep, _ = self.run_assemble()
        self.assertTrue(any("秘財20261028.pptx" in p and "打不開" in p for p in rep.problems))

    def test_roles_and_groups_without_new_file_listed(self):
        build([("教育培訓時間-3分鐘教育", [S("新三分鐘")])], self.up / "三分鐘20261028.pptx")
        build([("x", [S("乙組 請上台"), S("乙組 組長：簡 新總覽", "206 簡均安")])],
              self.mem / "25秒-乙組20261028.pptx", sections=False)
        deck, rep, prs = self.run_assemble()
        self.assertEqual(rep.unchanged, ["副主席", "秘財", "活動組", "甲組", "丙組"])
        self.assertIn("副主席、秘財、活動組、甲組、丙組", assemble.render_report(deck, rep, self.tmp / "out.pptx"))
        self.assertFalse(any("副主席" in p for p in rep.problems))  # 沒交是附註，不是問題

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


class MemberFolderGroupTest(Case):
    def test_group_file_in_member_folder(self):
        build([("x", [S("乙組 請上台"), S("乙組 組長：簡 新總覽", "206 簡均安")])],
              self.mem / "25秒-乙組 PowerTeam_20261104.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["25秒簡報-乙組"], ["乙組 請上台", "乙組 組長：簡 新總覽", "206 簡均安"])
        self.assertFalse(any("PowerTeam" in p for p in rep.problems))

    def test_mixed_file_split_into_group_and_member(self):
        build([("x", [S("NOW 丙組 請上台 NEXT 甲組"), S("丙組 新介紹"), S("No.", "301", "林新")])],
              self.mem / "301-林某_PowerTeam.pptx", sections=False)
        deck, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["25秒簡報-丙組"], ["丙組 請上台", "丙組 新介紹", "No."])
        self.assertIn("301", deck.member_sid)
        self.assertFalse(any("林某" in line for line in rep.lines + rep.problems))  # 會員檔名不進報告

    def test_member_file_scaled_to_master(self):
        src = build([("x", [S("230", "王新")])], self.tmp / "big.pptx", sections=False)
        prs = Presentation(str(src))
        prs.slide_width, prs.slide_height = prs.slide_width * 3 // 2, prs.slide_height * 3 // 2
        prs.save(str(self.mem / "230-王.pptx"))
        deck, rep, out = self.run_assemble()
        self.assertIn("230", deck.member_sid)
        self.assertTrue(any("230：會員檔投影片尺寸" in line for line in rep.lines))

    def test_subset_copy_with_link_to_dropped_slide(self):
        src = build([("x", [S("請上台頁"), S("組別頁"), S("個人頁")])], self.tmp / "linked.pptx", sections=False)
        prs = Presentation(str(src))
        slides = list(prs.slides)
        slides[2].shapes[0].click_action.target_slide = slides[0]  # 動作按鈕連回第 1 頁
        prs.save(str(src))
        out = self.tmp / "one.pptx"
        assemble.subset_copy(src, [2], out)
        got = Presentation(str(out))
        self.assertEqual([assemble.slide_text(s) for s in got.slides], ["個人頁"])
        got.save(str(out))  # 再存一次再開，仍是同一頁
        self.assertEqual([assemble.slide_text(s) for s in Presentation(str(out)).slides], ["個人頁"])

    def test_renamed_group_not_matched_to_next_group(self):
        # 組長把組名改了（主檔沒有「戊組」），請上台頁的 NEXT 寫著下一組「甲組」→ 不能換到甲組
        build([("x", [S("NOW 戊組 請上台 NEXT 甲組 請準備"), S("戊組 介紹")])], self.mem / "x.pptx", sections=False)
        before = first_lines(Presentation(str(self.master)))["25秒簡報-甲組"]
        _, rep, prs = self.run_assemble()
        self.assertTrue(any("「戊組」對不到主檔的組別" in p for p in rep.problems))
        self.assertCountEqual(first_lines(prs)["25秒簡報-甲組"], before)

    def test_group_name_read_between_now_and_stage(self):
        # NOW 前面的字（頁首、上一組）不算；只看 NOW 到「請上台」之間
        build([("x", [S("甲組 回顧 NOW 乙組 請上台 NEXT 丙組"), S("乙組 新介紹")])], self.mem / "x.pptx", sections=False)
        _, rep, prs = self.run_assemble()
        self.assertEqual(first_lines(prs)["25秒簡報-乙組"][:2], ["乙組 請上台", "乙組 新介紹"])

    def test_unknown_group_reported(self):
        build([("x", [S("NOW 戊組 請上台"), S("戊組 介紹")])], self.mem / "x.pptx", sections=False)
        _, rep, _ = self.run_assemble()
        self.assertTrue(any("對不到主檔的組別" in p for p in rep.problems))


class GuestTest(Case):
    def test_guests_shown(self):
        _, rep, prs = self.run_assemble()
        got = first_lines(prs)["來賓20秒自我介紹"]
        self.assertEqual(got[:2], ["來賓介紹", "許願池"])
        self.assertEqual(len(got), 3)
        self.assertTrue(all(s._element.get("show") is None for s in list(prs.slides)[:3]))

    def hidden(self, prs):
        return [assemble.slide_text(s).split("\n")[0] for s in prs.slides if s._element.get("show") == "0"]

    def test_guest_not_updated_hides_guest_sections(self):
        out = self.tmp / "out.pptx"
        _, rep = assemble.assemble(self.master, [], [], None, out)
        prs = Presentation(str(out))
        self.assertEqual(first_lines(prs)["來賓20秒自我介紹"], ["來賓介紹", "許願池", "舊來賓名單"])
        self.assertEqual([h for h in self.hidden(prs) if h != "續約宣示"], ["來賓介紹", "許願池", "舊來賓名單", "來賓感想頁"])

    def test_guests_shown_feedback_shown(self):
        _, _, prs = self.run_assemble()
        self.assertNotIn("來賓感想頁", self.hidden(prs))

    def test_missing_feedback_section_reported(self):
        master = build([sec for sec in MASTER if sec[0] != "來賓感想"], self.tmp / "m2.pptx")
        _, rep = assemble.assemble(master, [], [], None, self.tmp / "o2.pptx")
        self.assertTrue(any("來賓感想" in p for p in rep.problems))

    def test_guest_is_new(self):
        master = {"modifiedTime": "2026-10-27T16:02:00Z"}
        self.assertFalse(assemble.guest_is_new({"modifiedTime": "2026-10-27T10:47:00Z"}, master))
        self.assertTrue(assemble.guest_is_new({"modifiedTime": "2026-11-03T09:00:00Z"}, master))

    def test_empty_guest_list_hides_all(self):
        _, _, prs = self.run_assemble(guest_deck(self.tmp / "g0.pptx"))
        self.assertEqual(self.hidden(prs), ["來賓介紹", "許願池", "來賓1", "來賓感想頁"])


class PatchTest(Case):
    """同週補件：主檔已經是本週組過的，重跑不能再輪替，也不能動沒重傳的部分。"""

    def test_patch_does_not_rotate_and_only_changes_reuploaded(self):
        first = self.tmp / "20261028.pptx"
        build([("教育培訓時間-3分鐘教育", [S("第一版三分鐘")])], self.up / "三分鐘20261028.pptx")
        assemble.assemble(self.master, sorted(self.up.iterdir()), [], self.guest, first, stamp="2026-10-27T12:00:00Z")
        order = [n for n in first_lines(Presentation(str(first))) if n.startswith("25秒簡報-")]
        self.assertEqual(order, ["25秒簡報-乙組", "25秒簡報-丙組", "25秒簡報-甲組"])  # 第一次：輪替一次

        up2 = self.tmp / "up2"
        up2.mkdir()
        build([("秘書財務報告", [S("重傳的財務")]), ("5分鐘主題專講(秘財)", [S("重傳的專講")]),
               ("新會員/續約會員", [S("重傳的新會員")])], up2 / "秘財20261028.pptx")
        out = self.tmp / "patched.pptx"
        deck, rep = assemble.assemble(first, sorted(up2.iterdir()), [], None, out, patch=True)
        got = first_lines(Presentation(str(out)))
        self.assertEqual([n for n in got if n.startswith("25秒簡報-")], order)  # 沒有再輪替
        self.assertIn("重傳的財務", got["秘書財務報告"])
        self.assertIn("第一版三分鐘", got["教育培訓時間-3分鐘教育"])  # 沒重傳的不動
        text = assemble.render_report(deck, rep, out)
        self.assertIn("同週補件", text)
        self.assertNotIn("本週沒交新檔", text)

    def test_patch_without_new_guest_leaves_guest_sections_shown(self):
        first = self.tmp / "20261028.pptx"
        assemble.assemble(self.master, [], [], self.guest, first)
        out = self.tmp / "patched.pptx"
        assemble.assemble(first, [], [], None, out, patch=True)
        prs = Presentation(str(out))
        hidden = [assemble.slide_text(s).split("\n")[0] for s in prs.slides if s._element.get("show") == "0"]
        self.assertEqual([h for h in hidden if h != "續約宣示"], [])  # 來賓章節維持顯示
        self.assertIn("陳庭儀", "".join(assemble.slide_text(s) for s in prs.slides))

    def test_stamp_written_and_read_back(self):
        out = self.tmp / "out.pptx"
        self.assertIsNone(assemble.read_stamp(self.master))
        assemble.assemble(self.master, [], [], self.guest, out, stamp="2026-11-10T12:00:00Z")
        self.assertEqual(assemble.read_stamp(out), "2026-11-10T12:00:00Z")
        again = self.tmp / "again.pptx"
        assemble.assemble(out, [], [], None, again, patch=True)  # 本機補件沒給新時間 → 沿用原本的記錄
        self.assertEqual(assemble.read_stamp(again), "2026-11-10T12:00:00Z")


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
