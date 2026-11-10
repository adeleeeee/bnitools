#!/usr/bin/env python3
"""BNI 週會簡報組裝：上週主檔＋各職務／各組上傳檔＋會員個人檔＋來賓名單 → 本週主檔。

主簡報「分屍」成各自負責的檔，大家編輯後上傳到 Drive「Dismember-PPT」，程式依章節組回去。

上傳區：
  Dismember-PPT/                職務檔 <職務>YYYYMMDD.pptx、組別檔 25秒-<組名>YYYYMMDD.pptx
                                （日期在前也可以，如 YYYYMMDD副主席.pptx）
  會員資料夾（每週不同）        會員個人 25 秒單頁檔；dry-run 啟動時用 --members 給 Drive 資料夾連結
主檔：PPT 資料夾裡檔名 YYYYMMDD 開頭、日期最新的簡報（如 20261028_秘財更新.pptx）；同一天多份取 Drive 修改時間最新的

組裝規則（內容一律以「章節」為準，Mac PowerPoint 會弄丟備忘稿，章節比較可靠）：
  挑檔     上傳區新舊週混放：每個職務／組別只用檔名日期最新的一份（西元 8 碼或民國 7 碼）；
           日期不晚於主檔（主檔名 YYYYMMDD）＝已組過的舊檔不用；沒日期或同一天多份 → 略過並報告
           空檔（0 bytes）或打不開的檔 → 略過並報告，其他檔照常處理
  PDF      上傳檔是 PDF → 每頁轉成一張圖片投影片（主檔尺寸、等比例置中）再照下面規則處理；
           沒有章節，所以適用「沒有章節」的規則，切段找標題頁用 PDF 裡的文字
  尺寸     上傳檔（職務檔、組別檔）投影片尺寸和主檔不同但比例相同（如 Canva 匯出）→ 物件與字級等比例縮放；
           比例不同 → 整份略過並報告
  職務檔   上傳檔的章節名對應主檔章節，整段替換。章節名允許「主名-附註」。
           上傳檔有章節但沒有投影片 → 本週沒內容，保留空章節
           上傳檔缺少負責的章節     → 主檔原內容不動並報告
           上傳檔完全沒有章節       → 三分鐘：整份當作唯一章節；活動組：依各章節名的標題頁切段；
                                      秘財：依「五分鐘主題專講」「秘書財務報告」標題頁切段；
                                      主檔標題頁沒標 #固定 時連標題頁一起換成上傳檔的；
                                      其他職務整份略過並報告
           秘財檔沒有「新會員/續約會員」（或該章節是空的）＝本週沒有 → 只留 #續約會員 宣示頁並隱藏
           副主席的 BNI簡介 只換備忘稿含 #BNI台灣發展現況 的那一頁
           副主席其他章節、沒有章節的三分鐘／活動組：主檔章節第一頁是 #固定 標題頁就保留，
           上傳檔開頭同一張標題頁不重複放（金質獎章會員本來就沒有標題頁）
  組別檔   只換組別頁＝請上台頁之後、第一張會員個人頁之前的所有頁。請上台頁沿用主檔，
           上傳檔第一頁若是請上台頁就丟掉。主檔有標 #固定 的組會交叉比對，不一致就不換。
  會員檔   投影片上的編號優先，其次檔名；換掉該會員在組別章節裡的個人頁。
           會員資料夾裡也會有組長上傳的組別檔（檔名 25秒-<組名>，或第一頁是請上台頁）→ 當組別檔替換；
           檔裡夾帶的會員個人頁拆出來當會員檔。組名看檔名，檔名沒有就看請上台頁上 NEXT 之前最先出現的組名；
           對不到主檔的組名（例如組長自己改了組名）→ 略過並報告，不會換到 NEXT 寫的下一組。
  換組     組長上傳的組別頁上列有組員編號：會員個人頁目前所在的組沒有列他、
           而剛好另一組有列 → 搬到那組（排在第一位之後，其餘依編號）。判斷不了就不搬並報告。
  來賓     來賓名單投影片 Drive 修改時間沒有晚於主檔 → 沒更新＝本週沒有來賓
           有更新：來賓章節保留 #固定 頁，其餘換成來賓名單投影片
           沒有來賓（沒更新或名單空白）→「來賓20秒自我介紹」「來賓感想」兩個章節整段隱藏；有來賓 → 整段顯示
  輪替     上週第一組移到最後
  同週補件 主檔檔名日期＝本週例會日（星期三）＝這週已經組過 → 補件（patch）：不輪替；只處理上次抓檔之後
           才上傳的檔（職務檔、組別檔、會員檔、來賓名單），其餘頁面完全不動，手動改過的內容都留著。
           「上次抓檔時間」寫在成品的檔案屬性（關鍵字 BNI-ASSEMBLE fetched=…）；沒有就用主檔的 Drive 修改時間，
           也可以用 --since 指定。補件時來賓名單沒更新＝來賓章節不動（不會隱藏）。
  同名     上傳區或會員資料夾有同名檔（重傳時 Drive 常會兩份並存）→ 只用最後上傳的那一份
  沒交     職務／組別本週沒有新檔 → 不更新、沿用主檔內容，報告的「本週沒交新檔」列出來（附註，不算問題）
  報告     各步驟結果、問題清單、本週沒交新檔的職務與組別、會員成品頁碼、換組清單、含影片／音訊的頁（最後人工檢查）

會員個人頁＝頁面上剛好一個 3 位數會員編號，或帶有本程式的會員標記（cSld name BNI25S-<編號>）。

🔴 公開 repo：報告與 log 裡會員只寫編號，不寫姓名與會員檔名。職務檔與組別檔的檔名不含個資，可以印。

用法：
    R="uv run --with python-pptx --with pymupdf --with google-api-python-client --with google-auth-oauthlib --with google-auth python tools/assemble.py"
    $R dry-run --work scratch/assemble/run1 --members <本週會員資料夾連結>   # 成品只存本機；同週再跑會自動變補件
    $R dry-run --work scratch/assemble/run2 --members <連結> --since 2026-11-10T12:00:00Z   # 補件時指定「這個時間之後上傳的才處理」
    $R local --master m.pptx --uploads 上傳區/ --members 會員/ [--guest g.pptx] [--patch] --out o.pptx --report r.txt
"""
from __future__ import annotations

import argparse
import io
import re
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import pymupdf  # PDF 上傳檔轉圖片投影片
from lxml import etree
from pptx import Presentation
from pptx.exc import PackageNotFoundError
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.parts.slide import SlidePart

import bni25s
import bni_config
from bni25s import MemberProblem, slide_text

P14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
MEDIA_SUFFIXES = {"/video": "影片", "/audio": "音訊"}
GROUP_PREFIX = "25秒簡報-"
GROUP_FILE_PREFIX = "25秒-"

# 檔名開頭 → 這份檔負責的主檔章節
ROLES = {
    "副主席": ["BNI簡介", "表揚優秀會員", "金質獎章會員", "副主席報告"],
    "三分鐘": ["教育培訓時間-3分鐘教育"],
    "秘財": ["新會員/續約會員", "5分鐘主題專講(秘財)", "秘書財務報告"],
    "活動組": ["活動協調報告", "夥伴活動分享"],
}
VP_MONTHLY_TAG = "#BNI台灣發展現況"
HEADER_ROLES = {"副主席"}  # 有章節時也保留主檔 #固定 標題頁
NO_HEADER_SECTIONS = {"金質獎章會員"}  # 本來就沒有標題頁，不提醒
TITLE_SPLIT_ROLES = {"活動組", "秘財"}  # 沒有章節時依標題頁切段（簡報連續、各段開頭都有標題頁）
TITLE_ALIASES = {"5分鐘主題專講(秘財)": ("五分鐘主題專講", "5分鐘主題專講")}  # 標題頁上的字和章節名不同
NEW_RENEW = "新會員/續約會員"
RENEWAL_TAG = "#續約會員"  # 續約宣示頁：本週沒有新會員／續約時隱藏
RATIO_TOLERANCE = 0.01
PDF_RENDER_WIDTH = 1920  # PDF 轉圖的寬度（像素）
PDF_MIME = "application/pdf"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
GUEST_SECTION = "來賓20秒自我介紹"
GUEST_FEEDBACK_SECTION = "來賓感想"  # 分享與回饋後面那頁，跟著來賓有無顯示／隱藏
FIXED = ("#固定", "＃固定")
DATE_RE = re.compile(r"^\d{8}$")
MEMBER_NO_RE = re.compile(r"(?:^|NO\.\s*|\s)(\d{3})(?=\s|$)", re.MULTILINE | re.IGNORECASE)
LEADER_RE = re.compile(r"組長[：:]")
MEMBER_TAG_RE = re.compile(re.escape(bni25s.SLIDE_TAG_PREFIX) + r"(\d{3})$")


@dataclass
class Report:
    lines: list[str] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)
    moves: list[str] = field(default_factory=list)
    unchanged: list[str] = field(default_factory=list)  # 本週沒交新檔、沿用上週的職務與組別
    patch: bool = False  # 同週補件：沒輪替、只換新上傳的檔


# ---------------------------------------------------------------- 投影片判斷

def notes_of(slide) -> str:
    return slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""


def is_fixed(slide) -> bool:
    return any(t in notes_of(slide) for t in FIXED)


def member_number(slide) -> str | None:
    """會員個人頁的編號；不是個人頁回傳 None。
    #固定 頁與有「組長：」的頁是組別頁（組員總覽只有一位時，頁面上也剛好一個編號）。
    只看「組長：」不看「組長」：組長本人的個人頁上常寫著「組長」。"""
    m = MEMBER_TAG_RE.match(bni25s.slide_tag(slide) or "")
    if m:
        return m.group(1)
    text = slide_text(slide)
    if is_fixed(slide) or LEADER_RE.search(text):
        return None
    numbers = set(MEMBER_NO_RE.findall(text))
    return numbers.pop() if len(numbers) == 1 else None


def roster_numbers(slide) -> set[str]:
    return set(MEMBER_NO_RE.findall(slide_text(slide)))


def image_hashes(slide) -> frozenset:
    return frozenset(hash(r.target_part.blob) for r in slide.part.rels.values()
                     if not r.is_external and r.reltype.endswith("/image"))


def media_kinds(slide) -> list[str]:
    kinds = []
    for rel in slide.part.rels.values():
        for suf, label in MEDIA_SUFFIXES.items():
            if rel.reltype.endswith(suf) and label not in kinds:
                kinds.append(label)
    return kinds


def sections_of(prs) -> list[tuple[str, list]]:
    """[(章節名, [slide, ...])]，依章節清單順序。"""
    by_id = {str(s.slide_id): s for s in prs.slides}
    return [(sec.get("name"), [by_id[e.get("id")] for e in sec.iter(P14 + "sldId")])
            for sec in prs.part._element.iter(P14 + "section")]


def same_slide(a, b) -> bool:
    """同一張標題頁：文字相同（不計空白），或圖片完全相同。"""
    ta, tb = ("".join(slide_text(s).split()) for s in (a, b))
    if ta or tb:
        return ta == tb
    ia, ib = image_hashes(a), image_hashes(b)
    return bool(ia) and ia == ib


def _scale_part(el, kx: float, ky: float) -> None:
    """spTree 最上層物件的位置／大小（群組內部座標跟著群組走，不動）與所有字級。"""
    tree = el.find(f"{P_NS}cSld/{P_NS}spTree")
    for shape in tree:
        xfrm = next((x for x in shape.iter(A_NS + "xfrm", P_NS + "xfrm")), None)
        if xfrm is None:
            continue  # 位置繼承自版型／母片，版型母片也會縮
        off, ext = xfrm.find(A_NS + "off"), xfrm.find(A_NS + "ext")
        if off is not None:
            off.set("x", str(round(int(off.get("x")) * kx)))
            off.set("y", str(round(int(off.get("y")) * ky)))
        if ext is not None:
            ext.set("cx", str(round(int(ext.get("cx")) * kx)))
            ext.set("cy", str(round(int(ext.get("cy")) * ky)))
    for tag in ("rPr", "defRPr", "endParaRPr"):
        for r in el.iter(A_NS + tag):
            if r.get("sz"):
                r.set("sz", str(max(100, round(int(r.get("sz")) * kx))))


def fit_to_master(up, width: int, height: int) -> str | None:
    """上傳檔尺寸和主檔不同但比例相同 → 等比例縮放（記憶體內），回傳說明；比例不同就報錯。"""
    pw, ph = up.slide_width, up.slide_height
    if (pw, ph) == (width, height):
        return None
    kx, ky = width / pw, height / ph
    if abs(kx - ky) / kx > RATIO_TOLERANCE:
        raise MemberProblem(f"投影片比例和主檔不同（{pw}x{ph} vs {width}x{height}），整份略過")
    for s in up.slides:
        _scale_part(s._element, kx, ky)
    for m in up.slide_masters:
        for lay in m.slide_layouts:
            _scale_part(lay._element, kx, ky)
        _scale_part(m._element, kx, ky)
    up.slide_width, up.slide_height = width, height
    return f"投影片尺寸 {pw}x{ph} 和主檔不同，已等比例縮放 ×{kx:.3f}（請檢查版面）"


# ---------------------------------------------------------------- 主檔模型

class Deck:
    """主檔＋章節模型。搬移都在模型上做，最後一次寫回 sldIdLst 與 p14 章節清單。"""

    def __init__(self, path: Path, report: Report):
        self.prs = Presentation(str(path))
        self.report = report
        self.importer = bni25s.DeckImporter(self.prs)
        self.lst = self.prs.slides._sldIdLst
        self.entry = {e.get("id"): e for e in self.lst}
        self.slide = {str(s.slide_id): s for s in self.prs.slides}
        self.sections = []  # [章節名, 章節元素, [slide id, ...]]
        for sec in self.prs.part._element.iter(P14 + "section"):
            self.sections.append([sec.get("name"), sec, [e.get("id") for e in sec.iter(P14 + "sldId")]])
        ids = [i for _, _, ids in self.sections for i in ids]
        if sorted(ids) != sorted(self.entry):
            raise RuntimeError("主檔章節清單和投影片對不上（有投影片不屬於任何章節）")
        self.origin = {}  # 新匯入 slide id → (來源說明, 來源頁碼)
        self.member_sid = {}  # 本次換過的會員編號 → slide id

    def section(self, name: str):
        """完全同名優先；沒有就比對「-」前的主名（主檔可能叫「副主席報告-YYYYMMDD副主席」）。"""
        hits = [s for s in self.sections if s[0] == name] or \
            [s for s in self.sections if s[0].split("-")[0] == name]
        if len(hits) != 1:
            raise MemberProblem(f"主檔找不到（或有多個）章節「{name}」")
        return hits[0]

    def group_sections(self):
        return [s for s in self.sections if s[0].startswith(GROUP_PREFIX)]

    def group_split(self, sec) -> int:
        """組別章節裡第一張會員個人頁的位置；請上台頁是 0，組別頁是 1 到這個位置之前。"""
        ids = sec[2]
        return next((n for n in range(1, len(ids)) if member_number(self.slide[ids[n]])), len(ids))

    def drop(self, sid: str):
        e = self.entry.pop(sid)
        self.lst.remove(e)
        self.prs.part.drop_rel(e.get(qn("r:id")))

    def import_slide(self, src_slide, origin, tag="SEC", force_fixed=False) -> str:
        part = self.importer.import_slide(src_slide, tag)
        self.lst.add_sldId(self.prs.part.relate_to(part, RT_SLIDE))
        e = self.lst.sldId_lst[-1]
        sid = e.get("id")
        self.entry[sid] = e
        # DeckImporter 不帶備忘稿；#固定 等標記在備忘稿，下週還要用，照抄文字
        note = notes_of(src_slide)
        if force_fixed and not any(t in note for t in FIXED):
            note = (note + " #固定").strip()
        if note:
            part.slide.notes_slide.notes_text_frame.text = note
        self.slide[sid] = part.slide
        self.origin[sid] = origin
        return sid

    def import_many(self, items, force_fixed=False) -> list[str]:
        """items: [(slide, origin)]。全部成功才回傳；中途失敗就撤回已匯入的頁，主檔不留半套。"""
        done = []
        try:
            for s, origin in items:
                done.append(self.import_slide(s, origin, force_fixed=force_fixed))
        except MemberProblem:
            for sid in done:
                self.drop(sid)
                self.origin.pop(sid)
            raise
        return done

    def finish(self, out: Path) -> None:
        order = [i for _, _, ids in self.sections for i in ids]
        extra, missing = set(order) - set(self.entry), set(self.entry) - set(order)
        if extra or missing or len(order) != len(set(order)):
            raise RuntimeError(f"章節模型和投影片對不上：多 {len(extra)}、少 {len(missing)}、"
                               f"重複 {len(order) - len(set(order))}")
        for i in order:
            self.lst.append(self.entry[i])  # append 已存在的元素＝移到最後
        sec_lst = self.sections[0][1].getparent()
        for _, sec, _ in self.sections:
            sec_lst.append(sec)  # 章節清單順序跟著模型（輪替會改順序）
        for _, sec, ids in self.sections:
            sl = sec.find(P14 + "sldIdLst")
            for e in list(sl):
                sl.remove(e)
            for i in ids:
                etree.SubElement(sl, P14 + "sldId").set("id", i)
        pruned = bni25s.prune_unused_imports(self.prs)
        self.report.lines.append(f"清掉 {pruned} 個不再使用的匯入母片")
        # prs.slides 只在第一次讀取時依順序改名；這裡主動改名，避免 slideN 與順序不符
        self.prs.part.rename_slide_parts([e.get(qn("r:id")) for e in self.lst])
        names = [str(p.partname).lower() for p in self.prs.part.package.iter_parts()]
        if len(names) != len(set(names)):
            raise RuntimeError("產出檔內部有重複的 partname，停止")
        self.prs.save(str(out))


# ---------------------------------------------------------------- 上傳檔分類

def classify(name: str) -> tuple[str | None, str | None]:
    """('role', 職務) / ('group', 組名) / (None, None)。"""
    stem = re.sub(r"^\d{7,8}[\s_-]*", "", Path(name).stem)  # 日期在前：20261028副主席
    stem = re.sub(r"\d{8}.*$", "", stem).strip(" _-")
    if stem.startswith(GROUP_FILE_PREFIX):
        return "group", stem.removeprefix(GROUP_FILE_PREFIX)
    for role in ROLES:
        if stem.startswith(role) or (role == "三分鐘" and stem.startswith("3分鐘")):
            return "role", role
    return None, None


def file_date(name: str) -> str | None:
    """檔名裡的日期 → YYYYMMDD。認西元 8 碼（20261028）與民國 7 碼（1151028）。"""
    stem = Path(name).stem
    m = re.search(r"(?<!\d)(20\d{6})(?!\d)", stem)
    if m:
        return m.group(1)
    m = re.search(r"(?<!\d)(1\d{2})(\d{4})(?!\d)", stem)
    return f"{int(m.group(1)) + 1911}{m.group(2)}" if m else None


def select_uploads(uploads: list[Path], master: Path, rep: Report, same_week: bool = False) -> list[Path]:
    """上傳區新舊週混放：每個職務／組別只取檔名日期最新的一份。
    主檔檔名是 YYYYMMDD 時，日期不晚於主檔的視為已組過的舊檔，不用。沒有日期的檔無法判斷新舊，略過。
    same_week＝同週補件：主檔就是本週的，檔名日期和主檔同一天的是重傳檔，要用；更早的才是舊檔。"""
    after = master.stem if DATE_RE.match(master.stem) else None
    keep, groups = [], {}
    for f in sorted(uploads):
        key = classify(f.name)
        if key == (None, None):
            keep.append(f)  # 交給後面報「對不到任何職務或組別」
            continue
        d = file_date(f.name)
        if d is None:
            rep.problems.append(f"{f.name}：檔名沒有日期，無法判斷新舊，略過（請在檔名加 YYYYMMDD）")
            continue
        if after and (d < after if same_week else d <= after):
            rep.lines.append(f"[{f.name}] 日期{'早於' if same_week else '不晚於'}主檔 {after}，視為舊檔不用")
            continue
        groups.setdefault(key, []).append((d, f))
    for items in groups.values():
        top = max(d for d, _ in items)
        newest = [f for d, f in items if d == top]
        for d, f in items:
            if d != top:
                rep.lines.append(f"[{f.name}] 有更新的檔，不用")
        if len(newest) > 1:
            rep.problems.append("同一天有多份：" + "、".join(f.name for f in newest) + "，無法判斷用哪份，全部略過")
            continue
        keep.append(newest[0])
    return sorted(keep)


def match_section(upload_name: str, owned: list[str]) -> str | None:
    if upload_name in owned:
        return upload_name
    base = upload_name.split("-")[0]
    return base if base in owned else None


# ---------------------------------------------------------------- 各步驟

def pdf_to_deck(path: Path, width: int, height: int):
    """PDF 每頁轉成一張圖片投影片（主檔尺寸，等比例置中）。回傳 (簡報, slide→PDF 該頁文字)。"""
    prs = Presentation()
    prs.slide_width, prs.slide_height = width, height
    texts = {}
    with pymupdf.open(str(path)) as pdf:
        if pdf.page_count == 0:
            raise MemberProblem("PDF 沒有任何頁面，整份略過")
        for page in pdf:
            zoom = PDF_RENDER_WIDTH / page.rect.width
            png = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes("png")
            k = min(width / page.rect.width, height / page.rect.height)
            w, h = round(page.rect.width * k), round(page.rect.height * k)
            s = prs.slides.add_slide(prs.slide_layouts[6])
            s.shapes.add_picture(io.BytesIO(png), (width - w) // 2, (height - h) // 2, w, h)
            texts[s.slide_id] = page.get_text()
    return prs, lambda s: texts.get(s.slide_id, "")


def open_upload(deck: Deck, path: Path):
    """回傳 (上傳簡報, 取頁面文字的函式)。PDF 轉成圖片投影片，文字從 PDF 讀（切段找標題頁用）。"""
    w, h = deck.prs.slide_width, deck.prs.slide_height
    if path.stat().st_size == 0:
        raise MemberProblem("檔案是空的（0 bytes），上傳沒有成功，請重新上傳")
    if path.suffix.lower() == ".pdf":
        up, text_of = pdf_to_deck(path, w, h)
        deck.report.lines.append(f"  {path.name}：PDF {len(up.slides)} 頁轉成圖片投影片")
        return up, text_of
    try:
        up = Presentation(str(path))
    except (PackageNotFoundError, zipfile.BadZipFile, KeyError) as exc:
        raise MemberProblem("檔案打不開（不是正常的 pptx），請重新上傳") from exc
    note = fit_to_master(up, w, h)
    if note:
        deck.report.lines.append(f"  {path.name}：{note}")
    return up, slide_text


def replace_keep_header(deck: Deck, target: str, items, warn_missing: bool) -> str:
    """主檔章節第一頁是 #固定 標題頁就保留，後面換成 items；items 開頭若是同一張標題頁就不重複放。
    items: [(slide, origin)]。先匯入成功才刪舊頁。回傳報告用說明。"""
    sec = deck.section(target)
    head = sec[2][0] if sec[2] and is_fixed(deck.slide[sec[2][0]]) else None
    if head and items and same_slide(deck.slide[head], items[0][0]):
        items = items[1:]
    new = deck.import_many(items)
    for i in sec[2]:
        if i != head:
            deck.drop(i)
    sec[2] = ([head] if head else []) + new
    if head is None and warn_missing:
        deck.report.problems.append(f"主檔「{target}」第一頁沒有 #固定 標題頁，整段換掉（請在主檔標題頁備註加 #固定）")
    return f"標題頁沿用主檔，其餘換成 {len(new)} 頁" if head else f"整段換成 {len(new)} 頁"


def split_by_titles(slides: list, owned: list[str], text_of=slide_text) -> list[tuple[str, object, list]]:
    """沒有章節的上傳檔：依標題頁（章節主名或 TITLE_ALIASES 開頭）切段 → [(章節, 標題頁, 內容頁)]。
    新會員/續約會員 沒有標題頁＝本週沒有，不算錯（不回傳這段）。"""
    found = []
    for name in owned:
        titles = TITLE_ALIASES.get(name, (name.split("-")[0],))
        hits = [n for n, s in enumerate(slides) if text_of(s).strip().startswith(titles)]
        if not hits and name == NEW_RENEW:
            continue
        if len(hits) != 1:
            raise MemberProblem(f"沒有章節，標題頁「{titles[0]}」找到 {len(hits)} 頁，無法切段，整份略過")
        found.append((name, hits[0]))
    starts = [n for _, n in found]
    if starts[0] != 0 or starts != sorted(starts):
        raise MemberProblem("沒有章節，標題頁順序不對（第一頁要是第一段的標題頁），整份略過")
    ends = starts[1:] + [len(slides)]
    return [(name, slides[a], slides[a + 1:b]) for (name, a), b in zip(found, ends)]


def set_no_new_members(deck: Deck) -> str:
    """本週沒有新會員／續約：章節只留 #續約會員 宣示頁並隱藏。"""
    sec = deck.section(NEW_RENEW)
    keep = [i for i in sec[2] if RENEWAL_TAG in notes_of(deck.slide[i])]
    for i in sec[2]:
        if i not in keep:
            deck.drop(i)
    for i in keep:
        deck.slide[i]._element.set("show", "0")
    sec[2] = keep
    return f"  {NEW_RENEW}：本週沒有，續約宣示頁隱藏"


def apply_role_without_sections(deck: Deck, role: str, path: Path, up, text_of) -> None:
    owned = ROLES[role]
    slides = list(up.slides)
    pages = {id(s): i for i, s in enumerate(slides, 1)}
    rep = deck.report
    if len(owned) == 1:
        ids = deck.section(owned[0])[2]
        if not ids or not is_fixed(deck.slide[ids[0]]):
            raise MemberProblem(f"沒有章節，主檔「{owned[0]}」第一頁不是 #固定 標題頁，無法判斷，整份略過")
        done = replace_keep_header(deck, owned[0], [(s, (path.name, pages[id(s)])) for s in slides], True)
        rep.lines.append(f"  {owned[0]}：沒有章節，依整份；{done}")
        return
    if role not in TITLE_SPLIT_ROLES:
        raise MemberProblem("上傳檔沒有任何章節，整份略過（請用 PowerPoint 確認章節還在）")
    pieces = split_by_titles(slides, owned, text_of)
    for target, _, _ in pieces:  # 先全部檢查，避免換了一半
        deck.section(target)
    for target, title, body in pieces:
        ids = deck.section(target)[2]
        if ids and is_fixed(deck.slide[ids[0]]):  # 主檔標題頁有 #固定 → 用主檔的
            done = replace_keep_header(deck, target, [(s, (path.name, pages[id(s)])) for s in body], True)
        else:  # 主檔標題頁沒標 → 連標題頁一起換成上傳檔的
            sec = deck.section(target)
            new = deck.import_many([(s, (path.name, pages[id(s)])) for s in [title] + body])
            for i in sec[2]:
                deck.drop(i)
            sec[2] = new
            done = f"整段換成上傳檔 {len(new)} 頁（含標題頁）"
        rep.lines.append(f"  {target}：沒有章節，依標題頁切段；{done}")
    if role == "秘財" and NEW_RENEW not in [t for t, _, _ in pieces]:
        rep.lines.append(set_no_new_members(deck))


def apply_role(deck: Deck, role: str, path: Path) -> None:
    rep = deck.report
    up, text_of = open_upload(deck, path)
    secs = sections_of(up)
    if not secs:
        apply_role_without_sections(deck, role, path, up, text_of)
        return
    owned = ROLES[role]
    pages = {id(s): i for i, s in enumerate(up.slides, 1)}
    used = set()
    for up_name, slides in secs:
        target = match_section(up_name or "", owned)
        if target is None:
            rep.problems.append(f"{path.name}：章節「{up_name}」不屬於{role}，略過")
            continue
        if target != up_name:
            rep.lines.append(f"  {path.name}：章節「{up_name}」視為主檔「{target}」")
        used.add(target)
        try:
            sec = deck.section(target)
        except MemberProblem as exc:
            rep.problems.append(f"{path.name}：{exc}，這段略過")
            continue
        if target == "BNI簡介":
            old = [i for i in sec[2] if VP_MONTHLY_TAG in notes_of(deck.slide[i])]
            if len(old) != 1 or len(slides) != 1:
                rep.problems.append(f"{path.name}：BNI簡介 應各有 1 頁（主檔 {len(old)}、上傳 {len(slides)}），略過")
                continue
            pos = sec[2].index(old[0])
            [new] = deck.import_many([(slides[0], (path.name, pages[id(slides[0])]))])
            sec[2][pos] = new
            deck.drop(old[0])
            rep.lines.append(f"  {target}：換掉每月數據頁")
            continue
        if target == NEW_RENEW and not slides:
            rep.lines.append(set_no_new_members(deck))
            continue
        items = [(s, (path.name, pages[id(s)])) for s in slides]
        empty = "（空章節，本週沒有內容）" if not slides else ""
        if role in HEADER_ROLES:
            try:
                done = replace_keep_header(deck, target, items, target not in NO_HEADER_SECTIONS)
            except MemberProblem as exc:
                rep.problems.append(f"{path.name}：章節「{target}」匯入失敗：{exc}（這段沒換）")
                continue
            rep.lines.append(f"  {target}：{done}{empty}")
            continue
        try:
            new = deck.import_many(items)
        except MemberProblem as exc:
            rep.problems.append(f"{path.name}：章節「{target}」匯入失敗：{exc}（這段沒換）")
            continue
        for i in sec[2]:
            deck.drop(i)
        sec[2] = new
        rep.lines.append(f"  {target}：整段換成上傳檔 {len(slides)} 頁{empty}")
    for o in owned:
        if o in used:
            continue
        if o == NEW_RENEW:  # 秘財檔沒放這段＝本週沒有新會員／續約
            rep.lines.append(set_no_new_members(deck))
        else:
            rep.problems.append(f"{path.name}：沒有章節「{o}」，主檔這段維持不變")


def apply_group(deck: Deck, group: str, path: Path) -> None:
    sec = deck.section(GROUP_PREFIX + group)
    ids = sec[2]
    head = deck.slide[ids[0]]
    k = deck.group_split(sec)
    old = ids[1:k]
    if not old:
        raise MemberProblem(f"主檔「{group}」請上台頁之後直接就是會員個人頁，找不到組別頁")
    fixed_n = sum(1 for i in ids[1:] if is_fixed(deck.slide[i]))
    if fixed_n and fixed_n != len(old):
        raise MemberProblem(f"主檔「{group}」組別頁判斷不一致：依會員編號 {len(old)} 頁、依 #固定 {fixed_n} 頁")
    up, _ = open_upload(deck, path)
    head_key = notes_of(head).split()[:1]
    head_img = image_hashes(head)
    # 上傳檔第一頁若是請上台頁就不放（沿用主檔原頁）：備忘稿第一個詞相同、有「請上台」、或圖片相同
    new_src = [(i, s) for i, s in enumerate(up.slides, 1)
               if not (i == 1 and ((head_key and notes_of(s).split()[:1] == head_key)
                                   or "請上台" in slide_text(s)
                                   or (head_img and image_hashes(s) == head_img)))]
    new = deck.import_many([(s, (path.name, i)) for i, s in new_src], force_fixed=True)
    for i in old:
        deck.drop(i)
    sec[2] = [ids[0]] + new + ids[k:]
    deck.report.lines.append(f"  組別頁 {len(old)} 頁 → {len(new)} 頁（請上台頁沿用主檔）")


def member_index(deck: Deck) -> dict[str, list[tuple[list, str]]]:
    """會員編號 → [(組別章節, slide id)]，只看組別頁之後的會員個人頁。"""
    out = {}
    for sec in deck.group_sections():
        for sid in sec[2][deck.group_split(sec):]:
            n = member_number(deck.slide[sid])
            if n:
                out.setdefault(n, []).append((sec, sid))
    return out


@dataclass
class MemberFile:
    label: str      # 報告用：編號，或「第 N 個會員檔」
    filename: str   # 只用來解析，不得印出
    path: Path


def apply_members(deck: Deck, files: list[MemberFile]) -> None:
    rep = deck.report
    index = member_index(deck)
    parsed, seen = [], {}
    for f in files:
        try:
            mprs = Presentation(str(f.path))
        except (PackageNotFoundError, zipfile.BadZipFile, KeyError):
            rep.problems.append(f"{f.label}：會員檔無法開啟（不是有效的 pptx）")
            continue
        try:
            n = bni25s.resolve_upload_number(mprs, f.filename, exists=lambda x: len(index.get(x, [])) == 1)
        except MemberProblem as exc:
            rep.problems.append(f"{f.label}：{exc}")
            continue
        seen.setdefault(n, []).append(f)
        parsed.append((n, mprs))
    for n, mprs in parsed:
        if len(seen[n]) > 1:
            rep.problems.append(f"{n}：同一編號有 {len(seen[n])} 個會員檔，請只留一個（沒換）")
            continue
        hits = index.get(n, [])
        if len(hits) != 1:
            rep.problems.append(f"{n}：主檔組別章節裡有 {len(hits)} 張這位會員的個人頁，沒換")
            continue
        try:
            bni25s.check_member_deck(mprs, n)
            sec, old = hits[0]
            sid = deck.import_slide(mprs.slides[0], (f"會員 {n} 上傳檔", 1), tag=n)
        except MemberProblem as exc:
            rep.problems.append(f"{n}：{str(exc).removeprefix(f'{n} ')}（沒換）")
            continue
        sec[2][sec[2].index(old)] = sid
        deck.drop(old)
        deck.member_sid[n] = sid
    rep.lines.append(f"  換了 {len(deck.member_sid)} 位會員")


def auto_move(deck: Deck) -> None:
    """依組別頁上的組員編號，把放錯組的會員個人頁搬到新組。"""
    rep = deck.report
    groups = deck.group_sections()
    where, rosters = {}, {}
    for sec in groups:
        k = deck.group_split(sec)
        for sid in sec[2][k:]:
            n = member_number(deck.slide[sid])
            if n:
                where.setdefault(n, []).append((sec, sid))
        rosters[sec[0]] = set().union(*(roster_numbers(deck.slide[i]) for i in sec[2][1:k]))
    for n in sorted(where):
        if len(where[n]) != 1:
            rep.problems.append(f"{n}：有 {len(where[n])} 張個人頁，換組不處理")
            continue
        cur, sid = where[n][0]
        if not rosters[cur[0]] or n in rosters[cur[0]]:
            continue  # 這組組別頁看不出組員，或本來就在這組
        targets = [g for g in groups if g is not cur and n in rosters[g[0]]]
        here = cur[0].removeprefix(GROUP_PREFIX)
        if not targets:
            rep.problems.append(f"{n}：個人頁在「{here}」，但沒有任何一組的組員總覽列了他，請組長確認總覽頁")
            continue
        if len(targets) > 1:
            rep.problems.append(f"{n}：「{here}」組別頁沒有列他，另有 {len(targets)} 組列了他，沒有搬")
            continue
        dst = targets[0]
        cur[2].remove(sid)
        k = deck.group_split(dst)
        members = [i for i in dst[2][k:] if member_number(deck.slide[i])]
        # 第一位不動，其餘依編號：插在第一個編號比他大的人前面，沒有就放最後一位會員之後
        after = [i for i in members[1:] if member_number(deck.slide[i]) > n]
        if after:
            dst[2].insert(dst[2].index(after[0]), sid)
        elif members:
            dst[2].insert(dst[2].index(members[-1]) + 1, sid)
        else:
            dst[2].append(sid)
        rep.moves.append(f"  {n}：{cur[0].removeprefix(GROUP_PREFIX)} → {dst[0].removeprefix(GROUP_PREFIX)}")


def guest_names(guest_prs) -> list[str]:
    names = []
    for s in guest_prs.slides:
        for sh in s.shapes:
            if not sh.has_table:
                continue
            cells = [c.text.strip() for r in sh.table.rows for c in r.cells]
            for i, c in enumerate(cells[:-1]):
                if re.fullmatch(r"來賓\s*\d+", c) and cells[i + 1]:
                    names.append(cells[i + 1])
    return names


def apply_guests(deck: Deck, guest_path: Path) -> None:
    g = Presentation(str(guest_path))
    sec = deck.section(GUEST_SECTION)
    fixed = [i for i in sec[2] if is_fixed(deck.slide[i])]
    new = deck.import_many([(s, ("來賓名單投影片", n)) for n, s in enumerate(g.slides, 1)])
    for i in [i for i in sec[2] if i not in fixed]:
        deck.drop(i)
    sec[2] = fixed + new
    names = guest_names(g)
    if not names:
        deck.report.lines.append(f"  來賓 0 位；來賓名單 {len(new)} 頁")
        set_guest_visible(deck, False)
        return
    deck.report.lines.append(f"  來賓 {len(names)} 位；來賓名單 {len(new)} 頁")
    set_guest_visible(deck, True)


def set_guest_visible(deck: Deck, shown: bool) -> None:
    """有來賓：來賓章節與來賓感想全部顯示；沒有來賓：全部隱藏。"""
    targets = [GUEST_SECTION, GUEST_FEEDBACK_SECTION]
    for name in targets:
        try:
            ids = deck.section(name)[2]
        except MemberProblem:
            deck.report.problems.append(f"主檔沒有「{name}」章節，這段沒有{'顯示' if shown else '隱藏'}（請在主檔加章節）")
            continue
        for i in ids:
            el = deck.slide[i]._element
            if shown:
                el.attrib.pop("show", None)
            else:
                el.set("show", "0")
        deck.report.lines.append(f"  {name}：{'顯示' if shown else '隱藏'}")


def rotate_groups(deck: Deck) -> None:
    idx = [n for n, s in enumerate(deck.sections) if s[0].startswith(GROUP_PREFIX)]
    if not idx:
        raise RuntimeError(f"主檔沒有「{GROUP_PREFIX}」開頭的章節")
    if idx != list(range(idx[0], idx[-1] + 1)):
        raise RuntimeError("25 秒組別章節不連續，無法輪替")
    deck.sections.insert(idx[-1], deck.sections.pop(idx[0]))
    order = [s[0].removeprefix(GROUP_PREFIX) for s in deck.group_sections()]
    deck.report.lines.append("組別輪替（上週第一組改到最後）：" + " → ".join(order))


# ---------------------------------------------------------------- 組裝

def group_of_member_file(deck: Deck, mf: MemberFile, prs) -> str | None:
    """會員資料夾裡的組別檔（組長上傳）：檔名是 25秒-<組名>，或第一頁是請上台頁。回傳組名；不是組別檔回傳 None。
    組名：檔名開頭對主檔組名；檔名對不到就看請上台頁「NOW」到「請上台」之間寫的組名
    （沒有 NOW 就看「NEXT」之前）。NEXT 後面是下一組的名字，不能拿來當這份檔的組別
    （組長自己改了組名時會誤換到下一組）。"""
    kind, key = classify(mf.filename)
    first = slide_text(prs.slides[0]) if len(prs.slides) else ""
    if kind != "group" and "請上台" not in first:
        return None
    names = [s[0].removeprefix(GROUP_PREFIX) for s in deck.group_sections()]
    if kind == "group":
        hits = [g for g in names if key == g or key.startswith(g)]
    else:
        now = re.search(r"NOW(.*?)請上台", first, flags=re.IGNORECASE | re.DOTALL)
        own = now.group(1) if now else re.split(r"NEXT", first, maxsplit=1, flags=re.IGNORECASE)[0]
        hits = sorted((g for g in names if g in own), key=own.index)[:1]
        if not hits and now:
            written = " ".join(own.split())
            raise MemberProblem(f"看起來是組別檔，但請上台頁寫的組名「{written}」對不到主檔的組別，略過")
    if len(hits) != 1:
        raise MemberProblem(f"看起來是組別檔，但對不到主檔的組別（找到 {len(hits)} 組），略過")
    return hits[0]


def subset_copy(src: Path, keep: list[int], out: Path) -> None:
    """只留第 keep 頁（0 起算）另存；章節清單拿掉（組別檔、會員檔都不看章節）。
    被留下的頁用連結（如動作按鈕）指到的頁，即使拿掉也還在檔案裡；重開檔時 python-pptx 會把留下的頁
    改名成 slide1.xml…，跟這些殘留頁撞名、讀到錯的頁 → 殘留頁先改成不會撞的名字。"""
    prs = Presentation(str(src))
    lst = prs.slides._sldIdLst
    for n, e in reversed(list(enumerate(list(lst)))):
        if n not in keep:
            prs.part.drop_rel(e.get(qn("r:id")))
            lst.remove(e)
    kept = {id(prs.part.related_part(e.get(qn("r:id")))) for e in lst}
    leftovers = [p for p in prs.part.package.iter_parts()
                 if isinstance(p, SlidePart) and id(p) not in kept]
    for k, part in enumerate(leftovers, 1):
        part.partname = PackURI(f"/ppt/slides/slide_unlisted{k}.xml")
    for sec_lst in list(prs.part._element.iter(P14 + "sectionLst")):
        ext = sec_lst.getparent()
        ext.getparent().remove(ext)
    prs.save(str(out))


def sort_member_folder(deck: Deck, members: list[MemberFile], work: Path
                       ) -> tuple[list[tuple[str, Path, str]], list[MemberFile]]:
    """會員資料夾裡混了組別檔：拆成 [(組名, 組別檔, 報告用說明)] 和會員個人檔。
    組別檔裡的會員個人頁拆出來當會員檔，其餘頁（含請上台頁）當組別檔。
    拆出來的檔名只用組名與編號——會員檔名含姓名，不能出現在報告。"""
    groups, rest = [], []
    for mf in members:
        try:
            prs = Presentation(str(mf.path))
        except (PackageNotFoundError, zipfile.BadZipFile, KeyError):
            rest.append(mf)  # 交給 apply_members 報「無法開啟」
            continue
        try:
            g = group_of_member_file(deck, mf, prs)
        except MemberProblem as exc:
            deck.report.problems.append(f"{mf.label}：{exc}")
            continue
        work.mkdir(exist_ok=True)
        if g is None:
            try:
                rest.append(fitted_member(deck, mf, prs, work / f"fit_{len(rest)}.pptx"))
            except MemberProblem as exc:
                deck.report.problems.append(f"{mf.label}：{exc}")
            continue
        nums = {n: member_number(s) for n, s in enumerate(prs.slides)}
        gpath = work / f"25秒-{g}（會員資料夾{mf.label}）.pptx"
        subset_copy(mf.path, [n for n, x in nums.items() if not x], gpath)
        groups.append((g, gpath, mf.label))
        for n, x in nums.items():
            if x:
                mpath = work / f"{x}.pptx"
                subset_copy(mf.path, [n], mpath)
                one = MemberFile(label=x, filename=mpath.name, path=mpath)
                try:
                    rest.append(fitted_member(deck, one, Presentation(str(mpath)), mpath))
                except MemberProblem as exc:
                    deck.report.problems.append(f"{x}：{exc}")
    return groups, rest


def fitted_member(deck: Deck, mf: MemberFile, prs, out: Path) -> MemberFile:
    """會員檔尺寸和主檔不同（如 Canva）→ 等比例縮放另存到 out；尺寸相同或差不到 1%（原本就換得好）原檔照用。"""
    w, h = deck.prs.slide_width, deck.prs.slide_height
    if abs(prs.slide_width / w - 1) <= RATIO_TOLERANCE and abs(prs.slide_height / h - 1) <= RATIO_TOLERANCE:
        return mf
    note = fit_to_master(prs, w, h)
    if note is None:
        return mf
    prs.save(str(out))
    deck.report.lines.append(f"  {mf.label}：會員檔{note}")
    return MemberFile(label=mf.label, filename=mf.filename, path=out)


def assemble(master: Path, uploads: list[Path], members: list[MemberFile], guest: Path | None,
             out: Path, patch: bool = False, stamp: str | None = None) -> tuple[Deck, Report]:
    """guest 是 None＝來賓名單投影片沒更新：一般組裝當作本週沒有來賓；補件時來賓章節不動。
    patch＝同週補件：主檔已經是本週組過的，不輪替，沒交新檔的清單也不列（傳進來的本來就只有新上傳的檔）。
    stamp＝這次抓檔的時間，寫進成品的檔案屬性，下次補件用來判斷哪些檔是之後才上傳的。"""
    rep = Report(patch=patch)
    deck = Deck(master, rep)
    if patch:
        rep.lines.append("🔁 同週補件：主檔是本週已經組過的檔 → 不輪替，只換新上傳的檔，其餘頁面不動")
    member_groups, members = sort_member_folder(deck, members, out.parent / "member_split")
    selected = select_uploads(uploads, master, rep, same_week=patch)
    got = {classify(f.name) for f in selected}
    group_names = [s[0].removeprefix(GROUP_PREFIX) for s in deck.group_sections()]
    got_groups = {g for g, _, _ in member_groups} | {
        g for g in group_names for kind, key in got if kind == "group" and (key == g or key.startswith(g))}
    if not patch:
        rep.unchanged = [r for r in ROLES if ("role", r) not in got] + [g for g in group_names if g not in got_groups]
    for f in selected:
        kind, key = classify(f.name)
        rep.lines.append(f"[{f.name}]")
        try:
            if kind == "role":
                apply_role(deck, key, f)
            elif kind == "group":
                apply_group(deck, key, f)
            else:
                rep.problems.append(f"{f.name}：檔名對不到任何職務或組別，略過")
        except MemberProblem as exc:
            rep.problems.append(f"{f.name}：{exc}（這份沒換）")
    for g, gpath, label in member_groups:
        rep.lines.append(f"[會員資料夾的組別檔：{g}（{label}）]")
        try:
            apply_group(deck, g, gpath)
        except MemberProblem as exc:
            rep.problems.append(f"會員資料夾的組別檔 {g}（{label}）：{exc}（這份沒換）")
    rep.lines.append(f"[會員個人檔 {len(members)} 份]")
    apply_members(deck, members)
    auto_move(deck)
    rep.lines.append("[來賓名單投影片]")
    if guest is None and patch:
        rep.lines.append("  來賓名單投影片沒有新的更新，來賓章節不動")
    elif guest is None:
        rep.lines.append("  來賓名單投影片沒有比主檔新＝本週沒有來賓")
        set_guest_visible(deck, False)
    else:
        apply_guests(deck, guest)
    if not patch:
        rotate_groups(deck)
    if stamp:
        deck.prs.core_properties.keywords = STAMP_PREFIX + stamp
    deck.finish(out)
    return deck, rep


def render_report(deck: Deck, rep: Report, out: Path) -> str:
    final = Presentation(str(out))
    page = {str(s.slide_id): n for n, s in enumerate(final.slides, 1)}
    sec_of = {i: name for name, _, ids in deck.sections for i in ids}
    media = []
    for sid, n in sorted(page.items(), key=lambda kv: kv[1]):
        kinds = media_kinds(final.slides[n - 1])
        if kinds:
            src = deck.origin.get(sid)
            where = f"來自 {src[0]} 第 {src[1]} 頁" if src else "主檔原有"
            media.append(f"  第 {n} 頁｜{'、'.join(kinds)}｜{sec_of.get(sid)}｜{where}")
    members = [f"  {num} → 第 {page[sid]} 頁"
               for num, sid in sorted(deck.member_sid.items(), key=lambda kv: page[kv[1]])]
    return "\n".join(
        [f"成品：{out.name}（{len(final.slides)} 頁）", ""] + rep.lines
        + ["", "👤 會員個人頁（成品頁碼）："] + (members or ["  （無）"])
        + ["", "🔀 換組："] + (rep.moves or ["  （無）"])
        + ["", "⚠️ 問題／略過："] + ([f"  {p}" for p in rep.problems] or ["  （無）"])
        + ([] if rep.patch else
           ["", "📌 本週沒交新檔（沿用上週，不算問題）：", "  " + ("、".join(rep.unchanged) or "（無）")])
        + ["", "🎬 含影片／音訊的頁面（最後請逐頁檢查）："] + (media or ["  （無）"])
    )


def member_files(folder: Path | None) -> list[MemberFile]:
    if folder is None:
        return []
    paths = sorted(p for p in folder.iterdir() if p.suffix.lower() == ".pptx")
    return [MemberFile(label=_member_label(i, p.name), filename=p.name, path=p) for i, p in enumerate(paths, 1)]


def _member_label(idx: int, filename: str) -> str:
    try:
        return bni25s.parse_upload_number(filename)
    except MemberProblem:
        return f"第 {idx} 個會員檔"


def run_local(master: Path, uploads_dir: Path, members_dir: Path | None, guest: Path | None,
              out: Path, report_path: Path, patch: bool = False, stamp: str | None = None) -> int:
    uploads = sorted(p for p in uploads_dir.iterdir() if p.suffix.lower() in (".pptx", ".pdf"))
    deck, rep = assemble(master, uploads, member_files(members_dir), guest, out, patch=patch, stamp=stamp)
    text = render_report(deck, rep, out)
    report_path.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 1 if rep.problems else 0


# ---------------------------------------------------------------- Drive

def is_deck(f: dict) -> bool:
    return f["mimeType"] in (bni25s.PPTX_MIME, bni25s.GSLIDES_MIME)


def is_upload(f: dict) -> bool:
    return is_deck(f) or f["mimeType"] == PDF_MIME


MASTER_NAME_RE = re.compile(r"^(\d{8})")
FOLDER_LINK_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")


def pick_master(files: list[dict]) -> tuple[dict, str]:
    """PPT 資料夾裡檔名 YYYYMMDD 開頭的簡報（如 20261028.pptx、20261028_秘財更新.pptx）：
    日期最新的；同一天有多份取 Drive 修改時間最新的。回傳 (檔案, YYYYMMDD)。"""
    dated = [(m.group(1), f) for f in files if is_deck(f) and (m := MASTER_NAME_RE.match(f["name"]))]
    if not dated:
        raise SystemExit("❌ PPT 資料夾找不到 YYYYMMDD 開頭的主檔")
    date, f = max(dated, key=lambda df: (df[0], df[1]["modifiedTime"]))
    return f, date


def folder_id_from(link: str) -> str:
    """Drive 資料夾連結或 ID → ID。"""
    m = FOLDER_LINK_RE.search(link)
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", link):
        return link
    raise SystemExit("❌ --members 要給 Drive 資料夾連結（.../folders/<ID>）或資料夾 ID")


def guest_is_new(guest_f: dict, master_f: dict) -> bool:
    """來賓名單投影片在主檔之後有改過才換。"""
    return bni25s.parse_drive_time(guest_f["modifiedTime"]) > bni25s.parse_drive_time(master_f["modifiedTime"])


STAMP_PREFIX = "BNI-ASSEMBLE fetched="
STAMP_RE = re.compile(re.escape(STAMP_PREFIX) + r"([0-9T:.+\-Z]+)")


def now_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def read_stamp(path: Path) -> str | None:
    """成品檔案屬性裡的上次抓檔時間；沒有（手動組的、舊版程式組的）回傳 None。
    直接讀 zip 裡的 docProps/core.xml，不必為了一個屬性把幾百 MB 的簡報整份打開。"""
    with zipfile.ZipFile(path) as z:
        if "docProps/core.xml" not in z.namelist():
            return None
        m = STAMP_RE.search(z.read("docProps/core.xml").decode("utf-8"))
    return m.group(1) if m else None


def is_same_week(master_date: str, today: date) -> bool:
    """主檔檔名日期就是本次例會日（今天起最近的星期三）＝這週已經組過。"""
    return master_date == f"{bni25s.meeting_wednesday(today):%Y%m%d}"


def uploaded_at(f: dict) -> datetime:
    """檔案最後一次進 Drive 的時間。從電腦上傳的檔，modifiedTime 是檔案在對方電腦上的修改時間，
    可能比上傳時間早好幾天，所以和 createdTime 取較晚的。"""
    return max(bni25s.parse_drive_time(f[k]) for k in ("createdTime", "modifiedTime") if f.get(k))


def latest_by_name(files: list[dict]) -> tuple[list[dict], list[str]]:
    """同名檔只留最後上傳的一份（重傳時 Drive 常會兩份並存）。回傳 (留下的檔, 有重複的檔名)。"""
    best: dict[str, dict] = {}
    dup = []
    for f in files:
        old = best.get(f["name"])
        if old is not None and f["name"] not in dup:
            dup.append(f["name"])
        if old is None or uploaded_at(f) > uploaded_at(old):
            best[f["name"]] = f
    return list(best.values()), dup


def newer_than(files: list[dict], cutoff: str | None) -> list[dict]:
    """cutoff 之後才上傳的檔；cutoff 是 None＝不篩。"""
    if cutoff is None:
        return files
    t = bni25s.parse_drive_time(cutoff)
    return [f for f in files if uploaded_at(f) > t]


def fetch(svc, cfg: dict, work: Path, members_link: str, today: date, since: str | None = None
          ) -> tuple[Path, Path, Path, Path | None, bool]:
    """抓最新主檔、上傳區根目錄的簡報、本週會員資料夾（啟動時給連結）、來賓名單（有更新才抓）。
    主檔是本週已組過的＝補件：上傳區、會員資料夾、來賓名單都只抓上次抓檔之後才上傳的。
    回傳 (主檔, 上傳區, 會員資料夾, 來賓名單或 None, 是不是補件)。"""
    master_f, mdate = pick_master(bni25s.list_folder(svc, cfg["slides25s_ppt_folder_id"]))
    master = work / f"{mdate}.pptx"  # 存成 YYYYMMDD.pptx，挑上傳檔時要比日期
    print(f"主檔：{master_f['name']}（Drive 修改時間 {master_f['modifiedTime'][:16]} UTC）")
    bni25s.download(svc, master_f, master)
    patch = is_same_week(mdate, today)
    cutoff = None
    if patch:
        stamped = read_stamp(master)
        cutoff = since or stamped or master_f["modifiedTime"]
        how = "--since 指定" if since else "成品記錄的上次抓檔時間" if stamped else "主檔的 Drive 修改時間（成品沒有記錄抓檔時間）"
        print(f"🔁 同週補件：主檔日期 {mdate} 就是本週例會日 → 不輪替，只處理 {cutoff} 之後上傳的檔（{how}）")
    elif since:
        raise SystemExit(f"❌ --since 只能用在同週補件；主檔日期 {mdate} 不是本週例會日，這次是一般組裝")

    uploads_dir = work / "uploads"
    uploads_dir.mkdir()
    listed = bni25s.list_folder(svc, cfg["dismember_folder_id"])
    ups, dup = latest_by_name([f for f in listed if is_upload(f)])
    for name in dup:
        print(f"上傳區：{name} 有同名檔，只用最後上傳的那一份")
    ups = newer_than(ups, cutoff)
    for f in ups:
        ext = ".pdf" if f["mimeType"] == PDF_MIME else ".pptx"
        bni25s.download(svc, f, uploads_dir / (f["name"].removesuffix(ext) + ext))
    print(f"上傳區：{len(ups)} 份職務／組別檔（pptx／Google 簡報／PDF）")
    for f in listed:
        if not is_upload(f):
            print(f"上傳區：{f['name']} 不是簡報或 PDF，不處理")

    members_dir = work / "members"
    members_dir.mkdir()
    mid = folder_id_from(members_link)
    mname = svc.files().get(fileId=mid, fields="name", supportsAllDrives=True).execute()["name"]
    mfiles, mdup = latest_by_name([f for f in bni25s.list_folder(svc, mid) if is_deck(f)])
    if mdup:
        print(f"會員資料夾：{len(mdup)} 個檔名有同名檔，各只用最後上傳的那一份")
    mfiles = newer_than(mfiles, cutoff)
    for f in mfiles:
        bni25s.download(svc, f, members_dir / (f["name"].removesuffix(".pptx") + ".pptx"))
    print(f"會員資料夾：{mname}（{len(mfiles)} 份）")

    gf = svc.files().get(fileId=cfg["guest_slides_file_id"], fields="id, mimeType, modifiedTime, createdTime",
                         supportsAllDrives=True).execute()
    guest = None
    if newer_than([gf], cutoff) if patch else guest_is_new(gf, master_f):
        guest = work / "guest.pptx"
        bni25s.download(svc, gf, guest)
    print(f"來賓名單：Drive 修改時間 {gf['modifiedTime'][:16]} UTC → {'有更新，會換' if guest else '沒有新的更新，不換'}")
    return master, uploads_dir, members_dir, guest, patch


def cmd_dry_run(args) -> int:
    cfg = bni_config.load("assemble")
    work = Path(args.work)
    work.mkdir(parents=True, exist_ok=False)
    svc = bni25s.get_drive_service()
    today = date.fromisoformat(args.today) if args.today else datetime.now(ZoneInfo(cfg["timezone"])).date()
    stamp = now_stamp()  # 列檔之前就記時間：列檔之後才上傳的檔，下次補件要抓得到
    master, uploads_dir, members_dir, guest, patch = fetch(svc, cfg, work, args.members, today, args.since)
    return run_local(master, uploads_dir, members_dir, guest, work / "assembled.pptx", work / "report.txt",
                     patch=patch, stamp=stamp)


def cmd_local(args) -> int:
    return run_local(Path(args.master), Path(args.uploads), Path(args.members) if args.members else None,
                     Path(args.guest) if args.guest else None, Path(args.out), Path(args.report), patch=args.patch)


def main() -> None:
    p = argparse.ArgumentParser(description="BNI 週會簡報組裝")
    sub = p.add_subparsers(dest="cmd", required=True)
    dr = sub.add_parser("dry-run", help="從 Drive 抓最新檔案組裝，成品只存本機")
    dr.add_argument("--work", required=True, help="工作資料夾（必須是新的）")
    dr.add_argument("--members", required=True, help="本週會員個人檔的 Drive 資料夾連結（每週不同）")
    dr.add_argument("--since", help="同週補件時：只處理這個時間（UTC，如 2026-11-10T12:00:00Z）之後上傳的檔")
    dr.add_argument("--today", help="指定今天日期 YYYY-MM-DD（測試用）")
    dr.set_defaults(func=cmd_dry_run)
    lo = sub.add_parser("local", help="完全離線：本機檔案組裝")
    lo.add_argument("--master", required=True)
    lo.add_argument("--uploads", required=True, help="職務檔／組別檔所在資料夾")
    lo.add_argument("--members", help="會員個人檔所在資料夾")
    lo.add_argument("--guest", help="來賓名單投影片；不給＝來賓章節沿用主檔")
    lo.add_argument("--patch", action="store_true",
                    help="同週補件：主檔是本週已組過的，不輪替；--uploads／--members 只放這次要補的檔")
    lo.add_argument("--out", required=True)
    lo.add_argument("--report", required=True)
    lo.set_defaults(func=cmd_local)
    args = p.parse_args()
    try:
        code = args.func(args)
    except bni_config.ConfigError as exc:
        sys.exit(f"設定錯誤：{exc}")
    except Exception as exc:
        # Drive API 錯誤訊息含檔案 ID 與網址，公開 log 只印狀態碼
        from googleapiclient.errors import HttpError
        if isinstance(exc, HttpError):
            sys.exit(f"❌ Drive API 錯誤：HTTP {exc.resp.status}")
        raise
    sys.exit(code)


if __name__ == "__main__":
    main()
