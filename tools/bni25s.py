#!/usr/bin/env python3
"""BNI 週會 25 秒會員簡報：把會員上傳的單頁簡報換進主檔的共用邏輯。

週會流程由 tools/assemble.py 呼叫（會員頁比對、匯入、安全把關），手動執行、不排程：
各職務／各組交檔時間不固定。下面的 update／cleanup 指令是舊流程（BNI25S-YYYYMMDD 主檔
＋上傳區），保留供手動使用。

比對規則（找主檔裡「哪一頁是這位會員的」）：
  1. 這支程式換過的頁面帶有隱藏標記（slide 的 cSld name = BNI25S-<編號>），有標記就用標記
  2. 沒標記：版型名稱含「會員樣板」且某一行以編號開頭，必須剛好一頁
  分組總覽頁也有「編號 姓名」，但版型不是會員樣板，因此被排除。
  找不到或找到多頁都報錯，不猜。
上傳檔是哪位會員：先看投影片上「行首是 3 位數」且主檔有那頁的編號（會員常不照檔名規則），
  投影片上判斷不出來才用檔名的「剛好 3 位數」。兩邊不一致時以投影片為準。

替換方式等同 PowerPoint「保留來源格式」：會員頁連同它的版型、母片、佈景主題、圖片一起
匯入主檔。同一份版型只匯入一次（版型名稱帶 [bni25s:指紋]），換完後沒人用的匯入母片會清掉。

🔴 公開 repo：log 只能出現會員編號，不得出現姓名、檔名、Drive 檔案 ID。
   會員檔名格式錯誤時只寫「第 N 個檔案」。Drive API 錯誤只印 HTTP 狀態碼。

用法（本機）：
    R="uv run --with python-pptx --with google-api-python-client --with google-auth-oauthlib --with google-auth python tools/bni25s.py"
    $R update --dry-run --out scratch/bni25s/out.pptx    # 全部跑完但不上傳，產出存本機
    $R update --name-suffix -TEST                         # 上傳成 BNI25S-<日期>-TEST
    $R update --overwrite                                 # 本週檔已存在時就地覆蓋
    $R cleanup --dry-run
    $R apply-local --master m.pptx --upload 001王小明.pptx --out out.pptx   # 完全離線
"""
from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from lxml import etree
from pptx import Presentation
from pptx.exc import PackageNotFoundError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import PartFactory, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

import bni_config

WEDNESDAY = 2
DECK_PREFIX = "BNI25S-"
SLIDE_TAG_PREFIX = "BNI25S-"
MEMBER_LAYOUT_KEYWORD = "會員樣板"
LAYOUT_FP_RE = re.compile(r"\[bni25s:([0-9a-f]{12})\]")
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P14_SLD_ID = "{http://schemas.microsoft.com/office/powerpoint/2010/main}sldId"
# 投影片 / 版型 id 的合法下限（OOXML：sldMasterId 與 sldLayoutId 共用此空間）
MIN_MASTER_OR_LAYOUT_ID = 2147483648

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
GSLIDES_MIME = "application/vnd.google-apps.presentation"
FOLDER_MIME = "application/vnd.google-apps.folder"
SCOPES = ["https://www.googleapis.com/auth/drive"]


class MemberProblem(Exception):
    """單一會員的問題：記錄下來、繼續處理其他人，最後整次執行標為失敗。
    訊息只能含編號，不得含姓名或檔名（公開 log）。"""


# ---------------------------------------------------------------- 日期

def meeting_wednesday(today: date) -> date:
    """今天之後（含今天）最近的星期三 = 本次例會。"""
    return today + timedelta(days=(WEDNESDAY - today.weekday()) % 7)


def last_wednesday(today: date) -> date:
    """今天之前（含今天）最近的星期三 = 剛開完的例會（cleanup 用）。"""
    return today - timedelta(days=(today.weekday() - WEDNESDAY) % 7)


def deck_name(d: date) -> str:
    return f"{DECK_PREFIX}{d:%Y%m%d}"


# ---------------------------------------------------------------- 檔名

def parse_upload_number(filename: str) -> str:
    """從檔名抓會員編號：剛好 3 位數的數字。會員實際檔名很雜
    （`20261007_25秒＿173某某.pptx`、`○○會員25秒簡報_288_某某_20261014W12.pptx`），
    日期 8 位、`25秒`／`W12` 2 位、`1014` 4 位都不算。出現兩組不同的 3 位數 → 報錯不猜。"""
    numbers = set(re.findall(r"(?<!\d)\d{3}(?!\d)", filename))
    if not numbers:
        raise MemberProblem("檔名找不到 3 位數編號")
    if len(numbers) > 1:
        raise MemberProblem(f"檔名有 {len(numbers)} 組 3 位數字，無法判斷哪個是編號")
    return numbers.pop()


# 編號必須在某一行的開頭（「230」「232 某某」「No. 230」），
# 避免別人頁面上的電話 0912-230-456 之類被誤認。會員會寫 No.／NO.／no.，不分大小寫
_LEADING_NUMBER = r"^\s*(?:[Nn][Oo]\.?\s*)?{}(?!\d)"


def resolve_upload_number(member_prs, filename: str, prs=None, *, exists=None) -> str:
    """投影片上行首的 3 位數、且主檔找得到那頁 → 剛好一個就用它；否則退回檔名。
    exists(編號) 判斷主檔有沒有這位會員；不給就用 find_member_slide 在 prs 裡找。"""
    line_re = re.compile(_LEADING_NUMBER.format(r"(\d{3})"), re.MULTILINE)
    on_slide = set()
    for s in member_prs.slides:
        on_slide.update(line_re.findall(slide_text(s)))

    def in_master(number: str) -> bool:
        try:
            find_member_slide(prs, number)
        except MemberProblem:
            return False
        return True

    matched = {n for n in on_slide if (exists or in_master)(n)}
    try:
        from_name = parse_upload_number(filename)
    except MemberProblem as exc:
        if len(matched) == 1:
            return matched.pop()
        raise MemberProblem(f"投影片上沒有可對應主檔的編號，{exc}")
    if len(matched) == 1 and from_name not in matched:
        number = matched.pop()
        print(f"{number} ⚠️ 檔名編號 {from_name} 與投影片不同，以投影片為準")
        return number
    return from_name


# ---------------------------------------------------------------- 找會員頁

def _iter_text(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # GROUP
            yield from _iter_text(sh.shapes)
        elif sh.has_text_frame:
            yield sh.text_frame.text
        elif getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    yield cell.text


def slide_text(slide) -> str:
    return "\n".join(_iter_text(slide.shapes))


def slide_tag(slide) -> str | None:
    return slide._element.cSld.get("name")


def find_member_slide(prs, number: str) -> int:
    """回傳主檔中屬於該會員那一頁的 index（0 起算）。"""
    tag = SLIDE_TAG_PREFIX + number
    tagged = [i for i, s in enumerate(prs.slides) if slide_tag(s) == tag]
    if len(tagged) == 1:
        return tagged[0]
    if len(tagged) > 1:
        raise MemberProblem(f"{number} 主檔有 {len(tagged)} 頁帶相同標記")

    line_re = re.compile(_LEADING_NUMBER.format(re.escape(number)), re.MULTILINE)
    candidates = [
        i for i, s in enumerate(prs.slides)
        if MEMBER_LAYOUT_KEYWORD in s.slide_layout.name and line_re.search(slide_text(s))
    ]
    if not candidates:
        raise MemberProblem(f"{number} 主檔找不到對應頁（可能是整頁圖片，請人工替換一次）")
    if len(candidates) > 1:
        raise MemberProblem(f"{number} 主檔有 {len(candidates)} 頁符合，無法判斷是哪一頁")
    return candidates[0]


def check_member_deck(member_prs, number: str) -> None:
    # 不檢查投影片尺寸：會員檔常見 12179300 vs 主檔 12192000 這種微小差異，使用者決定不管
    n = len(member_prs.slides)
    if n != 1:
        raise MemberProblem(f"{number} 上傳的簡報有 {n} 頁，只能 1 頁")
    risky = find_risky_content(member_prs)
    if risky:
        raise MemberProblem(f"{number} 檔案含 {'、'.join(sorted(risky))}，基於安全拒收")


# 主檔太大（> 100 MB）Google 不掃毒，所以把關放在入口：會員檔裡會執行程式碼或夾帶
# 其他檔案的東西一律拒收。巨集本來就不會被帶進主檔（掛在整份簡報上，不在單頁），
# 但為了讓會員知道、也一併拒收。
_RISKY_CONTENT_TYPES = {
    "vbaProject": "巨集",
    "macroEnabled": "巨集",
    "oleObject": "嵌入物件",
    "activeX": "ActiveX",
}
_RISKY_RELTYPE_SUFFIXES = {
    "/vbaProject": "巨集",
    "/oleObject": "嵌入物件",
    "/control": "ActiveX",
    "/activeXControlBinary": "ActiveX",
    "/package": "嵌入檔案",
}


_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _is_chart_workbook(part, rel) -> bool:
    """圖表的資料表也是用 /package 關聯內嵌一份 xlsx（xlsx 不能含巨集），放行。"""
    return (part.content_type.endswith("drawingml.chart+xml") and not rel.is_external
            and rel.target_part.content_type == _XLSX_MIME)


def find_risky_content(prs) -> set[str]:
    found = set()
    for part in prs.part.package.iter_parts():
        for key, label in _RISKY_CONTENT_TYPES.items():
            if key in part.content_type:
                found.add(label)
        for rel in part.rels.values():
            for suffix, label in _RISKY_RELTYPE_SUFFIXES.items():
                if rel.reltype.endswith(suffix) and not _is_chart_workbook(part, rel):
                    found.add(label)
            if rel.is_external and rel.target_ref.lower().startswith("file:"):
                found.add("連到本機檔案的連結")
    return found


# ---------------------------------------------------------------- 匯入（保留來源格式）

class DeckImporter:
    """把會員簡報的單頁連同版型、母片、佈景主題、媒體匯入主檔。"""

    def __init__(self, prs):
        self.prs = prs
        self.pkg = prs.part.package
        # OPC 規格：partname 不分大小寫。image1.gif 與 image1.GIF 會撞名，PowerPoint 會要求修復
        self._partnames = {str(p.partname).lower() for p in self.pkg.iter_parts()}
        self._layout_by_fp = {}
        for master in prs.slide_masters:
            for layout in master.slide_layouts:
                m = LAYOUT_FP_RE.search(layout.name)
                if m:
                    self._layout_by_fp[m.group(1)] = layout.part

    # -- partname：自己追蹤已配發的名稱，因為遞迴複製時新 part 尚未掛進關聯圖，
    #    package.next_partname() 看不到它們，會配出重複名稱。
    #    投影片例外：python-pptx 每次讀 prs.slides 都會把投影片依順序改名為 slideN，
    #    事先配好的 slideN 會過期撞名，所以新投影片用不會撞到 slideN 的暫用名稱。
    def _alloc_partname(self, src_partname: str, content_type: str) -> PackURI:
        if content_type.endswith("slide+xml"):
            tmpl = "/ppt/slides/bni25s%d.xml"
        else:
            tmpl = re.sub(r"\d*(\.[^./]+)$", lambda m: "%d" + m.group(1).lower(), src_partname.replace("%", "%%"))
        n = 1
        while (tmpl % n).lower() in self._partnames:
            n += 1
        name = tmpl % n
        self._partnames.add(name.lower())
        return PackURI(name)

    def _skip_rel(self, src_part, rel, memo) -> bool:
        ct = src_part.content_type
        if ct.endswith("slide+xml"):
            if rel.reltype in (RT.NOTES_SLIDE, RT.COMMENTS):
                return True
            if rel.reltype == RT.SLIDE:
                raise MemberProblem("含有連到其他投影片的連結，不支援")
        if ct.endswith("slideMaster+xml") and rel.reltype == RT.SLIDE_LAYOUT:
            # 母片只帶入正在匯入的那一個版型，其他版型不帶
            return rel.target_part not in memo
        return False

    def _copy_part(self, src, memo):
        if src in memo:
            return memo[src]
        new = PartFactory(self._alloc_partname(str(src.partname), src.content_type), src.content_type, self.pkg, src.blob)
        memo[src] = new
        if not isinstance(new, XmlPart):
            if len(src.rels):
                raise MemberProblem(f"內嵌物件 {src.content_type} 不支援")
            return new

        mapping = {}
        for rId, rel in list(src.rels.items()):
            if self._skip_rel(src, rel, memo):
                continue
            if rel.is_external:
                mapping[rId] = new.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                mapping[rId] = new.relate_to(self._copy_part(rel.target_part, memo), rel.reltype)

        el = new._element
        if src.content_type.endswith("slideMaster+xml"):
            lst = el.find(qn("p:sldLayoutIdLst"))
            for entry in list(lst):
                if entry.get(R_NS + "id") not in mapping:
                    lst.remove(entry)
        unresolved = set()
        for node in el.iter():
            for key, val in node.attrib.items():
                if key.startswith(R_NS):
                    if val == "":  # PowerPoint 常留空的 r:id（如動作按鈕），不是關聯
                        continue
                    if val in mapping:
                        node.set(key, mapping[val])
                    else:
                        unresolved.add(val)
        if unresolved:
            raise MemberProblem(f"{src.content_type} 有無法對應的關聯")
        return new

    def _next_master_or_layout_id(self) -> int:
        pres = self.prs.part._element
        ids = [int(e.get("id")) for e in pres.find(qn("p:sldMasterIdLst"))]
        for master in self.prs.slide_masters:
            lst = master._element.find(qn("p:sldLayoutIdLst"))
            ids += [int(e.get("id")) for e in lst]
        return max(ids + [MIN_MASTER_OR_LAYOUT_ID - 1]) + 1

    @staticmethod
    def fingerprint(layout_part) -> str:
        master_part = layout_part.part_related_by(RT.SLIDE_MASTER)
        h = hashlib.sha256()
        for part in (layout_part, master_part):
            h.update(part.blob)
            for rel in part.rels.values():
                if not rel.is_external and rel.reltype != RT.SLIDE_LAYOUT and rel.target_part is not master_part:
                    h.update(rel.target_part.blob)
        return h.hexdigest()[:12]

    def _import_layout(self, src_layout, fp: str):
        memo = {}
        new_layout = self._copy_part(src_layout, memo)
        src_master = src_layout.part_related_by(RT.SLIDE_MASTER)
        new_master = memo[src_master]

        # 先取號再掛上去：新母片尚未加入清單，取號時不會讀到還沒有 id 的項目
        master_id = self._next_master_or_layout_id()
        layout_entry = new_master._element.find(qn("p:sldLayoutIdLst"))[0]
        layout_entry.set("id", str(master_id + 1))
        rId = self.prs.part.relate_to(new_master, RT.SLIDE_MASTER)
        lst = self.prs.part._element.find(qn("p:sldMasterIdLst"))
        entry = etree.SubElement(lst, qn("p:sldMasterId"))
        entry.set("id", str(master_id))
        entry.set(R_NS + "id", rId)

        csld = new_layout._element.find(qn("p:cSld"))
        csld.set("name", f"{csld.get('name', '')} [bni25s:{fp}]".strip())
        self._layout_by_fp[fp] = new_layout
        return new_layout

    def import_slide(self, src_slide, number: str):
        src_layout = src_slide.part.part_related_by(RT.SLIDE_LAYOUT)
        fp = self.fingerprint(src_layout)
        layout = self._layout_by_fp.get(fp) or self._import_layout(src_layout, fp)
        new_slide = self._copy_part(src_slide.part, {src_layout: layout})
        new_slide._element.find(qn("p:cSld")).set("name", SLIDE_TAG_PREFIX + number)
        return new_slide


def replace_member_slide(prs, importer: DeckImporter, index: int, member_prs, number: str) -> None:
    new_part = importer.import_slide(member_prs.slides[0], number)
    rId = prs.part.relate_to(new_part, RT.SLIDE)
    sld_id_lst = prs.slides._sldIdLst
    sld_id_lst.add_sldId(rId)
    new_entry = sld_id_lst.sldId_lst[-1]
    old_entry = sld_id_lst.sldId_lst[index]
    old_entry.addprevious(new_entry)
    sld_id_lst.remove(old_entry)
    prs.part.drop_rel(old_entry.rId)
    # PowerPoint 的章節清單（p14:sectionLst）用 slide id 記錄每頁屬於哪個章節，
    # 不跟著換的話舊 id 變成孤兒、新頁不屬於任何章節
    old_id, new_id = old_entry.get("id"), new_entry.get("id")
    for el in prs.part._element.iter(P14_SLD_ID):
        if el.get("id") == old_id:
            el.set("id", new_id)


def prune_unused_imports(prs) -> int:
    """移除本程式匯入、且已沒有任何頁面使用的母片。原本就在主檔的母片一律不動。"""
    used = {s.slide_layout.part for s in prs.slides}
    lst = prs.part._element.find(qn("p:sldMasterIdLst"))
    removed = 0
    for entry in list(lst):
        rId = entry.get(R_NS + "id")
        master = prs.part.related_part(rId)
        layouts = [rel.target_part for rel in master.rels.values() if rel.reltype == RT.SLIDE_LAYOUT]
        imported = all(LAYOUT_FP_RE.search(l._element.find(qn("p:cSld")).get("name", "")) for l in layouts)
        if layouts and imported and not any(l in used for l in layouts):
            lst.remove(entry)
            prs.part.drop_rel(rId)
            removed += 1
    return removed


# ---------------------------------------------------------------- 套用（離線核心）

@dataclass
class UploadFile:
    label: str          # log 用：編號，或「第 N 個檔案」
    filename: str       # 只用來解析，不得印出
    path: Path


def apply_uploads(master_path: Path, uploads: list[UploadFile], out_path: Path) -> tuple[list[str], list[str]]:
    """回傳 (成功的編號, 問題清單)。有任何成功才寫出 out_path。"""
    prs = Presentation(str(master_path))
    importer = DeckImporter(prs)
    done, problems, seen = [], [], {}

    parsed = []
    for up in uploads:
        try:
            member = Presentation(str(up.path))
        except (PackageNotFoundError, zipfile.BadZipFile, KeyError):
            problems.append(f"{up.label} ❌ 檔案無法開啟（不是有效的 pptx）")
            continue
        try:
            number = resolve_upload_number(member, up.filename, prs)
        except MemberProblem as exc:
            problems.append(f"{up.label} ❌ {exc}")
            continue
        seen.setdefault(number, []).append(up)
        parsed.append((up, number, member))

    for up, number, member in parsed:
        if len(seen[number]) > 1:
            problems.append(f"{number} ❌ 同一編號有 {len(seen[number])} 個檔案，請只留一個")
            continue
        try:
            check_member_deck(member, number)
            index = find_member_slide(prs, number)
            replace_member_slide(prs, importer, index, member, number)
        except MemberProblem as exc:
            msg = str(exc).removeprefix(f"{number} ")
            problems.append(f"{number} ❌ {msg}")
            continue
        done.append(number)
        print(f"{number} ✅ 已替換（第 {index + 1} 頁）")

    if done:
        pruned = prune_unused_imports(prs)
        if pruned:
            print(f"清掉 {pruned} 個不再使用的匯入母片")
        len(prs.slides)  # 觸發 python-pptx 依順序重新命名投影片，再檢查
        names = [str(p.partname).lower() for p in prs.part.package.iter_parts()]
        if len(names) != len(set(names)):
            raise RuntimeError("產出檔內部有重複的 partname（不分大小寫），停止（不會上傳）")
        prs.save(str(out_path))
    return done, problems


# ---------------------------------------------------------------- Drive

def get_drive_service():
    """雲端：GDRIVE_CLIENT_ID / GDRIVE_CLIENT_SECRET / GDRIVE_REFRESH_TOKEN 三個環境變數。
    本機：沿用 gdrive_update.py 的 token 快取。"""
    env_names = ("GDRIVE_CLIENT_ID", "GDRIVE_CLIENT_SECRET", "GDRIVE_REFRESH_TOKEN")
    present = [n for n in env_names if os.environ.get(n)]
    if not present:
        import gdrive_update
        return gdrive_update.get_service()
    if len(present) != len(env_names):
        missing = [n for n in env_names if n not in present]
        sys.exit(f"雲端授權環境變數不完整，缺少：{', '.join(missing)}")

    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GDRIVE_REFRESH_TOKEN"],
        client_id=os.environ["GDRIVE_CLIENT_ID"],
        client_secret=os.environ["GDRIVE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_folder(svc, folder_id: str, extra_q: str = "") -> list[dict]:
    q = f"'{folder_id}' in parents and trashed = false and mimeType != '{FOLDER_MIME}'"
    if extra_q:
        q += f" and ({extra_q})"
    files, token = [], None
    while True:
        resp = svc.files().list(
            q=q, pageToken=token, pageSize=200,
            fields="nextPageToken, files(id, name, mimeType, modifiedTime, createdTime)",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        files += resp.get("files", [])
        token = resp.get("nextPageToken")
        if not token:
            return files


def find_decks(svc, folder_id: str, name: str) -> list[dict]:
    return list_folder(svc, folder_id, f"name = '{name}' or name = '{name}.pptx'")


def download(svc, f: dict, out: Path) -> None:
    from googleapiclient.http import MediaIoBaseDownload
    if f["mimeType"] == GSLIDES_MIME:
        req = svc.files().export_media(fileId=f["id"], mimeType=PPTX_MIME)
    else:
        req = svc.files().get_media(fileId=f["id"], supportsAllDrives=True)
    with out.open("wb") as fh:
        dl = MediaIoBaseDownload(fh, req, chunksize=64 * 1024 * 1024)
        done = False
        while not done:
            _, done = dl.next_chunk()


def upload(svc, path: Path, *, folder_id: str, name: str, existing_id: str | None) -> None:
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(str(path), mimetype=PPTX_MIME, resumable=True, chunksize=64 * 1024 * 1024)
    if existing_id:
        svc.files().update(fileId=existing_id, media_body=media, supportsAllDrives=True).execute()
    else:
        body = {"name": name, "parents": [folder_id]}
        svc.files().create(body=body, media_body=media, fields="id", supportsAllDrives=True).execute()


def is_abusive_file_error(exc) -> bool:
    """Drive 掃毒判定有毒的檔案，API 下載會回 403 cannotDownloadAbusiveFile。
    其他 403（例如沒權限）不算，照常往上拋。"""
    return exc.resp.status == 403 and b"cannotDownloadAbusiveFile" in (exc.content or b"")


def parse_drive_time(s: str) -> datetime:
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


# ---------------------------------------------------------------- 指令

def _today(args, tz: ZoneInfo) -> date:
    return date.fromisoformat(args.today) if args.today else datetime.now(tz).date()


def _label(idx: int, filename: str) -> str:
    try:
        return parse_upload_number(filename)
    except MemberProblem:
        return f"第 {idx} 個檔案"


def cmd_update(args) -> int:
    from googleapiclient.errors import HttpError

    cfg = bni_config.load("bni25s")
    tz = ZoneInfo(cfg["timezone"])
    w = meeting_wednesday(_today(args, tz))
    src_name, dst_name = deck_name(w - timedelta(days=7)), deck_name(w) + args.name_suffix
    print(f"本次例會 {w}：{src_name} → {dst_name}" + ("（dry-run，不寫入 Drive）" if args.dry_run else ""))

    svc = get_drive_service()
    ppt_folder, upload_folder = cfg["slides25s_ppt_folder_id"], cfg["slides25s_upload_folder_id"]

    sources = find_decks(svc, ppt_folder, src_name)
    if len(sources) != 1:
        sys.exit(f"❌ 找到 {len(sources)} 個 {src_name}，必須剛好 1 個")
    targets = find_decks(svc, ppt_folder, dst_name)
    if len(targets) > 1:
        sys.exit(f"❌ 已有 {len(targets)} 個 {dst_name}，請人工整理成最多 1 個")
    existing_id = targets[0]["id"] if targets else None
    if existing_id and not args.overwrite:
        sys.exit(f"❌ {dst_name} 已存在。確定要重跑請加 --overwrite（會就地覆蓋該檔）")

    stale_before = datetime.combine(w - timedelta(days=7), time(0), tz)
    files = sorted(list_folder(svc, upload_folder), key=lambda f: f["name"])
    print(f"上傳區共 {len(files)} 個檔案")

    problems: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        uploads = []
        for idx, f in enumerate(files, 1):
            label = _label(idx, f["name"])
            if parse_drive_time(f["modifiedTime"]) < stale_before:
                problems.append(f"{label} ❌ 上傳時間早於上週三，可能是沒清掉的舊檔，略過")
                continue
            if f["mimeType"] not in (PPTX_MIME, GSLIDES_MIME):
                problems.append(f"{label} ❌ 不是簡報檔，略過")
                continue
            path = tmpdir / f"upload_{idx}.pptx"
            try:
                download(svc, f, path)
            except HttpError as exc:
                if not is_abusive_file_error(exc):
                    raise
                problems.append(f"{label} ❌ Google 判定疑似有毒，略過（請人工確認）")
                continue
            uploads.append(UploadFile(label=label, filename=f["name"], path=path))

        master_path, out_path = tmpdir / "master.pptx", tmpdir / "out.pptx"
        if uploads:
            print("下載主檔…")
            download(svc, sources[0], master_path)
            done, member_problems = apply_uploads(master_path, uploads, out_path)
            problems += member_problems
        else:
            done = []

        if not done:
            print("沒有任何會員頁被替換，新檔內容與上週相同")
            if not uploads:
                print("下載主檔…")
                download(svc, sources[0], master_path)
            out_path = master_path

        if args.out:
            Path(args.out).write_bytes(out_path.read_bytes())
            print(f"本機副本：{args.out}")
        if args.dry_run:
            print("dry-run：略過上傳")
        else:
            print("上傳新檔…")
            upload(svc, out_path, folder_id=ppt_folder, name=dst_name, existing_id=existing_id)
            print(f"✅ 已{'覆蓋' if existing_id else '建立'} {dst_name}")

    print(f"\n替換 {len(done)} 位、問題 {len(problems)} 項")
    for p in problems:
        print(p)
    return 1 if problems else 0


def cmd_cleanup(args) -> int:
    from googleapiclient.errors import HttpError

    cfg = bni_config.load("bni25s")
    tz = ZoneInfo(cfg["timezone"])
    late_after = datetime.combine(last_wednesday(_today(args, tz)) - timedelta(days=1), time(15, 0), tz)

    svc = get_drive_service()
    files = sorted(list_folder(svc, cfg["slides25s_upload_folder_id"]), key=lambda f: f["name"])
    print(f"上傳區共 {len(files)} 個檔案" + ("（dry-run，不移動）" if args.dry_run else ""))

    problems = []
    for idx, f in enumerate(files, 1):
        label = _label(idx, f["name"])
        note = "（週二 15:00 後才上傳，可能沒套用到）" if parse_drive_time(f["modifiedTime"]) > late_after else ""
        if not args.dry_run:
            try:
                svc.files().update(fileId=f["id"], body={"trashed": True}, supportsAllDrives=True).execute()
            except HttpError as exc:
                if exc.resp.status != 403:
                    raise
                problems.append(f"{label} ❌ 沒有權限移到垃圾桶（HTTP 403），請人工刪除")
                continue
        print(f"{label} 🗑 {'將移到' if args.dry_run else '已移到'}垃圾桶{note}")

    for p in problems:
        print(p)
    return 1 if problems else 0


def cmd_apply_local(args) -> int:
    uploads = [
        UploadFile(label=_label(i, Path(p).name), filename=Path(p).name, path=Path(p))
        for i, p in enumerate(args.upload, 1)
    ]
    done, problems = apply_uploads(Path(args.master), uploads, Path(args.out))
    print(f"\n替換 {len(done)} 位、問題 {len(problems)} 項" + (f"，已存 {args.out}" if done else ""))
    for p in problems:
        print(p)
    return 1 if problems else 0


def main() -> None:
    p = argparse.ArgumentParser(description="BNI 25 秒會員簡報自動替換")
    sub = p.add_subparsers(dest="cmd", required=True)

    up = sub.add_parser("update", help="替換會員頁並另存本週三的新檔")
    up.add_argument("--dry-run", action="store_true", help="全部跑完但不寫入 Drive")
    up.add_argument("--overwrite", action="store_true", help="本週檔已存在時就地覆蓋")
    up.add_argument("--name-suffix", default="", help="新檔名加後綴（測試用，例 -TEST）")
    up.add_argument("--out", help="另存一份產出到本機路徑")
    up.add_argument("--today", help="指定今天日期 YYYY-MM-DD（測試用）")
    up.set_defaults(func=cmd_update)

    cl = sub.add_parser("cleanup", help="上傳區檔案移到垃圾桶")
    cl.add_argument("--dry-run", action="store_true", help="只列出不移動")
    cl.add_argument("--today", help="指定今天日期 YYYY-MM-DD（測試用）")
    cl.set_defaults(func=cmd_cleanup)

    al = sub.add_parser("apply-local", help="離線：本機主檔＋本機上傳檔 → 本機產出")
    al.add_argument("--master", required=True)
    al.add_argument("--upload", required=True, action="append", help="可重複指定")
    al.add_argument("--out", required=True)
    al.set_defaults(func=cmd_apply_local)

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
