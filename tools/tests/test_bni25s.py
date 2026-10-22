"""bni25s 的測試（unittest，全部使用合成資料，不碰網路）。

直接跑：
    uv run --with python-pptx --with google-api-python-client --with XlsxWriter python tools/tests/test_bni25s.py
"""
import importlib.util
import io
import struct
import sys
import tempfile
import unittest
import warnings
import zlib
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Inches

_TOOLS = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_TOOLS))  # bni25s 會 import 同目錄的 bni_config
_spec = importlib.util.spec_from_file_location("bni25s", _TOOLS / "bni25s.py")
bni25s = importlib.util.module_from_spec(_spec)
sys.modules["bni25s"] = bni25s
_spec.loader.exec_module(bni25s)


def png_bytes(rgb=(255, 0, 0)) -> bytes:
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    raw = b"\x00" + bytes(rgb)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


# 1x1 GIF
GIF = bytes.fromhex("47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b")


def reload(prs):
    buf = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("error")  # zipfile 的 Duplicate name 警告 = 產出損毀檔
        prs.save(buf)
    buf.seek(0)
    return Presentation(buf)


def set_layout_name(layout, name):
    layout._element.find(qn("p:cSld")).set("name", name)


def add_text(slide, *lines):
    for i, line in enumerate(lines):
        box = slide.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(4), Inches(1))
        box.text_frame.text = line


def make_master():
    """0 歡迎頁／1 分組總覽（非會員樣板，也有編號＋姓名）／2、3 會員頁。"""
    prs = Presentation()
    member_layout = prs.slide_layouts[6]
    set_layout_name(member_layout, "會員樣板")
    add_text(prs.slides.add_slide(prs.slide_layouts[5]), "歡迎")
    add_text(prs.slides.add_slide(prs.slide_layouts[5]), "230 王小明", "115 李大華")
    add_text(prs.slides.add_slide(member_layout), "230", "王小明")
    add_text(prs.slides.add_slide(member_layout), "115 李大華 Mason")
    return reload(prs)


def make_member(*lines, slides=1, link=None, layout_name="會員樣板", rgb=(255, 0, 0)):
    prs = Presentation()
    layout = prs.slide_layouts[6]
    set_layout_name(layout, layout_name)
    for _ in range(slides):
        s = prs.slides.add_slide(layout)
        add_text(s, *lines)
        s.shapes.add_picture(io.BytesIO(png_bytes(rgb)), Inches(5), Inches(1))
        if link:
            s.shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address = link
    return reload(prs)


def texts(slide):
    return bni25s.slide_text(slide)


class DateTest(unittest.TestCase):
    def test_tuesday_run(self):
        w = bni25s.meeting_wednesday(date(2026, 10, 13))
        self.assertEqual(w, date(2026, 10, 14))
        self.assertEqual(bni25s.deck_name(w), "BNI25S-20261014")

    def test_on_wednesday_is_same_day(self):
        self.assertEqual(bni25s.meeting_wednesday(date(2026, 10, 14)), date(2026, 10, 14))

    def test_thursday_goes_to_next_week(self):
        self.assertEqual(bni25s.meeting_wednesday(date(2026, 10, 15)), date(2026, 10, 21))

    def test_year_boundary(self):
        w = bni25s.meeting_wednesday(date(2027, 1, 5))
        self.assertEqual(w, date(2027, 1, 6))
        self.assertEqual(bni25s.deck_name(w - bni25s.timedelta(days=7)), "BNI25S-20261230")

    def test_last_wednesday(self):
        self.assertEqual(bni25s.last_wednesday(date(2026, 10, 14)), date(2026, 10, 14))
        self.assertEqual(bni25s.last_wednesday(date(2026, 10, 15)), date(2026, 10, 14))
        self.assertEqual(bni25s.last_wednesday(date(2026, 10, 13)), date(2026, 10, 7))


class FilenameTest(unittest.TestCase):
    parse = staticmethod(bni25s.parse_upload_number)

    def test_real_world_shapes(self):
        """會員實際檔名的形狀（姓名換成假的）。"""
        cases = {
            "示範_230王小明.pptx": "230",
            "示範230＿王小明.PPTX": "230",
            "20261007_25秒＿173王小明律師.pptx": "173",
            "示範_238王小明名片(小馬哥)1014.pptx": "238",
            "示範會員25秒簡報_288_王小明_20261014W12.pptx": "288",
            "王小明 115": "115",
        }
        for name, number in cases.items():
            self.assertEqual(self.parse(name), number, name)

    def test_leading_zero_kept(self):
        self.assertEqual(self.parse("示範008_王小明.pptx"), "008")

    def test_same_number_twice_is_fine(self):
        self.assertEqual(self.parse("230_王小明_230.pptx"), "230")

    def test_no_number(self):
        for name in ("王小明.pptx", "20261007_王小明.pptx", "王小明_25秒_W12.pptx", "王小明1014.pptx"):
            with self.assertRaises(bni25s.MemberProblem, msg=name):
                self.parse(name)

    def test_two_numbers_rejected(self):
        with self.assertRaises(bni25s.MemberProblem):
            self.parse("230_王小明_115.pptx")

    def test_error_message_has_no_personal_name(self):
        with self.assertRaises(bni25s.MemberProblem) as ctx:
            self.parse("示範_王小明.pptx")
        self.assertNotIn("王小明", str(ctx.exception))


class ResolveNumberTest(unittest.TestCase):
    def setUp(self):
        self.prs = make_master()

    def resolve(self, filename, *lines):
        return bni25s.resolve_upload_number(make_member(*lines), filename, self.prs)

    def test_slide_number_when_filename_has_none(self):
        self.assertEqual(self.resolve("華榮會員25秒簡報王小明.pptx", "某某公司", "230", "王小明"), "230")

    def test_number_and_name_in_one_line(self):
        self.assertEqual(self.resolve("王小明（20261020).pptx", "115 李大華 Nick"), "115")

    def test_slide_wins_over_filename(self):
        self.assertEqual(self.resolve("示範_115王小明.pptx", "230", "王小明"), "230")

    def test_number_not_in_master_ignored(self):
        """「100% 滿意」行首也是 3 位數，但主檔沒有 100 號 → 不算。"""
        self.assertEqual(self.resolve("王小明.pptx", "100% 滿意", "230", "王小明"), "230")

    def test_no_prefix_any_case(self):
        self.assertEqual(self.resolve("王小明.pptx", "NO.230 王小明"), "230")
        self.assertEqual(self.resolve("王小明.pptx", "no. 230 王小明"), "230")

    def test_filename_fallback(self):
        self.assertEqual(self.resolve("示範_230王小明.pptx", "只有圖片"), "230")

    def test_two_master_numbers_on_slide_uses_filename(self):
        self.assertEqual(self.resolve("示範_115.pptx", "230 王小明", "115 李大華"), "115")

    def test_nothing_found(self):
        with self.assertRaises(bni25s.MemberProblem) as ctx:
            self.resolve("示範_王小明.pptx", "只有圖片")
        self.assertNotIn("王小明", str(ctx.exception))


class FindSlideTest(unittest.TestCase):
    def setUp(self):
        self.prs = make_master()

    def test_skips_overview_slide(self):
        self.assertEqual(bni25s.find_member_slide(self.prs, "230"), 2)

    def test_number_and_name_in_one_box(self):
        self.assertEqual(bni25s.find_member_slide(self.prs, "115"), 3)

    def test_no_prefix_number(self):
        add_text(self.prs.slides.add_slide(self.prs.slides[2].slide_layout), "No. 246", "張三")
        self.assertEqual(bni25s.find_member_slide(self.prs, "246"), 4)

    def test_not_found(self):
        with self.assertRaises(bni25s.MemberProblem):
            bni25s.find_member_slide(self.prs, "999")

    def test_number_boundary(self):
        with self.assertRaises(bni25s.MemberProblem):
            bni25s.find_member_slide(self.prs, "23")

    def test_number_inside_phone_ignored(self):
        """別人會員頁上的電話含 230，不能被當成 230 的頁。"""
        add_text(self.prs.slides.add_slide(self.prs.slides[2].slide_layout), "246", "電話 0912-230-456")
        self.assertEqual(bni25s.find_member_slide(self.prs, "230"), 2)

    def test_multiple_candidates_rejected(self):
        add_text(self.prs.slides.add_slide(self.prs.slides[2].slide_layout), "230", "王小明")
        with self.assertRaises(bni25s.MemberProblem):
            bni25s.find_member_slide(self.prs, "230")

    def test_tag_wins(self):
        self.prs.slides[3]._element.find(qn("p:cSld")).set("name", "BNI25S-230")
        self.assertEqual(bni25s.find_member_slide(self.prs, "230"), 3)


class ReplaceTest(unittest.TestCase):
    def replace(self, prs, number, member):
        importer = bni25s.DeckImporter(prs)
        bni25s.check_member_deck(member, number)
        idx = bni25s.find_member_slide(prs, number)
        bni25s.replace_member_slide(prs, importer, idx, member, number)
        bni25s.prune_unused_imports(prs)
        return reload(prs)

    def test_replace_in_place(self):
        prs = make_master()
        before = [s.part.blob for s in prs.slides]
        out = self.replace(prs, "230", make_member("230", "王小明 新內容"))

        self.assertEqual(len(out.slides), 4)
        new = out.slides[2]
        self.assertIn("新內容", texts(new))
        self.assertEqual(bni25s.slide_tag(new), "BNI25S-230")
        self.assertRegex(new.slide_layout.name, r"^會員樣板 \[bni25s:[0-9a-f]{12}\]$")
        pics = [sh for sh in new.shapes if sh.shape_type == 13]
        self.assertEqual(pics[0].image.blob, png_bytes())
        for i in (0, 1, 3):
            self.assertEqual(out.slides[i].part.blob, before[i], f"第 {i} 頁不該變動")
        self.assertEqual(len(out.slide_masters), 2)

    def test_empty_relationship_id_ignored(self):
        """動作按鈕常留 r:id=""，那不是關聯，不能當成「無法對應的關聯」拒收。"""
        member = make_member("230", "王小明 新內容")
        run = member.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        rpr = run._r.get_or_add_rPr()
        link = rpr.makeelement(qn("a:hlinkClick"), {})
        link.set(qn("r:id"), "")
        rpr.append(link)
        out = self.replace(make_master(), "230", reload(member))
        self.assertIn("新內容", texts(out.slides[2]))

    def test_no_case_insensitive_partname_collision(self):
        """會員圖片副檔名大寫（image1.GIF），主檔已有 image1.gif → 不能配出只差大小寫的名字。"""
        from pptx.opc.packuri import PackURI
        prs = make_master()
        prs.slides[0].shapes.add_picture(io.BytesIO(GIF), Inches(1), Inches(1))
        prs = reload(prs)
        self.assertIn("/ppt/media/image1.gif", {str(p.partname) for p in prs.part.package.iter_parts()})
        member = make_member("230", "王小明")
        pic = [sh for sh in member.slides[0].shapes if sh.shape_type == 13][0]
        member.slides[0].shapes._spTree.remove(pic._element)
        member.slides[0].shapes.add_picture(io.BytesIO(GIF), Inches(5), Inches(1))
        for part in member.part.package.iter_parts():
            if str(part.partname).endswith(".gif"):
                part.partname = PackURI(str(part.partname)[:-4] + ".GIF")
        out = self.replace(prs, "230", reload(member))
        names = [str(p.partname).lower() for p in out.part.package.iter_parts()]
        self.assertEqual(len(names), len(set(names)))

    def test_retagged_slide_found_next_week_even_without_text(self):
        prs = self.replace(make_master(), "230", make_member("只有圖片"))
        self.assertEqual(bni25s.find_member_slide(prs, "230"), 2)

    def test_same_template_imported_once(self):
        prs = self.replace(make_master(), "230", make_member("230", "王小明"))
        prs = self.replace(prs, "115", make_member("115 李大華 Mason"))
        self.assertEqual(len(prs.slide_masters), 2)

    def test_unused_imported_master_pruned(self):
        prs = self.replace(make_master(), "230", make_member("230", "王小明"))
        prs = self.replace(prs, "230", make_member("230", "王小明", layout_name="會員樣板 新版"))
        self.assertEqual(len(prs.slide_masters), 2)
        self.assertIn("新版", prs.slides[2].slide_layout.name)

    def test_original_master_never_pruned(self):
        prs = self.replace(make_master(), "230", make_member("230", "王小明"))
        prs = self.replace(prs, "115", make_member("115 李大華 Mason"))
        # 原母片的會員樣板已沒人用，但不是本程式匯入的，必須保留
        self.assertEqual(prs.slide_masters[0].slide_layouts[6].name, "會員樣板")

    def test_external_hyperlink_kept(self):
        out = self.replace(make_master(), "230", make_member("230", "王小明", link="https://example.invalid/x"))
        run = out.slides[2].shapes[0].text_frame.paragraphs[0].runs[0]
        self.assertEqual(run.hyperlink.address, "https://example.invalid/x")

    def test_two_slides_rejected(self):
        with self.assertRaises(bni25s.MemberProblem):
            bni25s.check_member_deck(make_member("230", slides=2), "230")

    def test_slightly_different_slide_size_accepted(self):
        """會員檔常見 12179300 vs 主檔 12192000，使用者決定不管尺寸。"""
        member = make_member("230", "王小明 新內容")
        member.slide_width = member.slide_width - 12700
        out = self.replace(make_master(), "230", reload(member))
        self.assertIn("新內容", texts(out.slides[2]))


P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def add_sections(prs, groups):
    """在 presentation.xml 加上 PowerPoint 章節：groups = [(名稱, [slide index...])]。"""
    from lxml import etree
    pres = prs.part._element
    ids = [e.get("id") for e in prs.slides._sldIdLst.sldId_lst]
    ext_lst = pres.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, qn("p:extLst"))
    ext = etree.SubElement(ext_lst, qn("p:ext"), uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
    sec_lst = etree.SubElement(ext, f"{{{P14}}}sectionLst")
    for n, (name, idxs) in enumerate(groups):
        sec = etree.SubElement(sec_lst, f"{{{P14}}}section", name=name, id=f"{{00000000-0000-0000-0000-00000000000{n}}}")
        lst = etree.SubElement(sec, f"{{{P14}}}sldIdLst")
        for i in idxs:
            etree.SubElement(lst, f"{{{P14}}}sldId", id=ids[i])


def section_map(prs):
    pos = {e.get("id"): i for i, e in enumerate(prs.slides._sldIdLst.sldId_lst)}
    out = {}
    for sec in prs.part._element.iter(f"{{{P14}}}section"):
        out[sec.get("name")] = [pos.get(e.get("id")) for e in sec.iter(f"{{{P14}}}sldId")]
    return out


class SectionTest(unittest.TestCase):
    def test_replaced_slide_stays_in_its_section(self):
        prs = make_master()
        add_sections(prs, [("開場", [0, 1]), ("25秒", [2, 3])])
        prs = reload(prs)
        importer = bni25s.DeckImporter(prs)
        bni25s.replace_member_slide(prs, importer, 2, make_member("230", "新內容"), "230")
        out = reload(prs)
        self.assertEqual(section_map(out), {"開場": [0, 1], "25秒": [2, 3]})


class RiskyContentTest(unittest.TestCase):
    def _member_with_part(self, partname, content_type, reltype, on_slide=True):
        from pptx.opc.package import Part
        from pptx.opc.packuri import PackURI
        prs = make_member("230", "王小明")
        part = Part(PackURI(partname), content_type, prs.part.package, b"payload")
        (prs.slides[0].part if on_slide else prs.part).relate_to(part, reltype)
        return reload(prs)

    def test_clean_member(self):
        self.assertEqual(bni25s.find_risky_content(make_member("230", "王小明")), set())

    def test_macro_rejected(self):
        prs = self._member_with_part(
            "/ppt/vbaProject.bin", "application/vnd.ms-office.vbaProject",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject", on_slide=False)
        self.assertIn("巨集", bni25s.find_risky_content(prs))
        with self.assertRaises(bni25s.MemberProblem):
            bni25s.check_member_deck(prs, "230")

    def test_ole_object_rejected(self):
        prs = self._member_with_part(
            "/ppt/embeddings/oleObject1.bin", "application/vnd.openxmlformats-officedocument.oleObject",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject")
        self.assertIn("嵌入物件", bni25s.find_risky_content(prs))

    def test_embedded_package_rejected(self):
        prs = self._member_with_part(
            "/ppt/embeddings/Microsoft_Word_Document1.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package")
        self.assertIn("嵌入檔案", bni25s.find_risky_content(prs))

    def test_local_file_link_rejected(self):
        prs = make_member("230", "王小明", link="file:///C:/payload.exe")
        self.assertIn("連到本機檔案的連結", bni25s.find_risky_content(prs))

    def test_chart_workbook_allowed(self):
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE
        prs = make_member("230", "王小明")
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("S", (1, 2))
        prs.slides[0].shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0, Inches(4), Inches(3), data)
        self.assertEqual(bni25s.find_risky_content(reload(prs)), set())


class AbusiveFileErrorTest(unittest.TestCase):
    class FakeErr:
        def __init__(self, status, content):
            self.resp = type("R", (), {"status": status})()
            self.content = content

    def test_abusive(self):
        err = self.FakeErr(403, b'{"error":{"errors":[{"reason":"cannotDownloadAbusiveFile"}]}}')
        self.assertTrue(bni25s.is_abusive_file_error(err))

    def test_other_403_not_abusive(self):
        self.assertFalse(bni25s.is_abusive_file_error(self.FakeErr(403, b'{"reason":"insufficientFilePermissions"}')))
        self.assertFalse(bni25s.is_abusive_file_error(self.FakeErr(500, b"cannotDownloadAbusiveFile")))


class ApplyUploadsTest(unittest.TestCase):
    def test_mixed_batch(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            make_master().save(d / "master.pptx")
            files = {
                "20261007_25秒＿230王小明.pptx": make_member("230", "王小明 新內容"),
                "王小明.pptx": make_member("x"),
                "示範115_李大華 Mason.pptx": make_member("115", slides=2),
                "示範999_陳大文.pptx": make_member("999"),
            }
            uploads = []
            for i, (name, prs) in enumerate(files.items(), 1):
                path = d / f"u{i}.pptx"
                prs.save(path)
                label = bni25s._label(i, name)
                uploads.append(bni25s.UploadFile(label=label, filename=name, path=path))
            (d / "broken.pptx").write_bytes(b"not a zip")
            uploads.append(bni25s.UploadFile(label="777", filename="示範777_張三.pptx", path=d / "broken.pptx"))

            done, problems = bni25s.apply_uploads(d / "master.pptx", uploads, d / "out.pptx")

            self.assertEqual(done, ["230"])
            self.assertEqual(len(problems), 4)
            joined = "\n".join(problems)
            self.assertIn("第 2 個檔案", joined)
            for personal in ("王小明", "李大華", "陳大文", "張三"):
                self.assertNotIn(personal, joined)
            self.assertIn("新內容", texts(Presentation(str(d / "out.pptx")).slides[2]))

    def test_filename_without_number_replaced(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            make_master().save(d / "master.pptx")
            make_member("某某公司", "230", "王小明 新內容").save(d / "u1.pptx")
            name = "華榮會員25秒簡報王小明.pptx"
            uploads = [bni25s.UploadFile(label=bni25s._label(1, name), filename=name, path=d / "u1.pptx")]
            done, problems = bni25s.apply_uploads(d / "master.pptx", uploads, d / "out.pptx")
            self.assertEqual((done, problems), (["230"], []))
            self.assertIn("新內容", texts(Presentation(str(d / "out.pptx")).slides[2]))

    def test_duplicate_number_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            make_master().save(d / "master.pptx")
            uploads = []
            for i, name in enumerate(("示範230_王小明.pptx", "示範_230王小明.pptx"), 1):
                make_member("230").save(d / f"u{i}.pptx")
                uploads.append(bni25s.UploadFile(label="230", filename=name, path=d / f"u{i}.pptx"))
            done, problems = bni25s.apply_uploads(d / "master.pptx", uploads, d / "out.pptx")
            self.assertEqual(done, [])
            self.assertEqual(len(problems), 2)
            self.assertFalse((d / "out.pptx").exists())


if __name__ == "__main__":
    unittest.main()
