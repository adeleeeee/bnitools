#!/usr/bin/env python3
"""BNI 週會簡報組裝：上週主檔＋各職務／各組上傳檔＋會員個人檔＋來賓名單 → 本週主檔。

主簡報「分屍」成各自負責的檔，大家編輯後上傳到 Drive「Dismember-PPT」，程式依章節組回去。

上傳區（Dismember-PPT）：
  根目錄              職務檔 <職務>YYYYMMDD.pptx、組別檔 25秒-<組名>YYYYMMDD.pptx
  YYYYMMDD/ 子資料夾  會員個人 25 秒單頁檔（取名稱日期最新的那個資料夾）

組裝規則（內容一律以「章節」為準，Mac PowerPoint 會弄丟備忘稿，章節比較可靠）：
  職務檔   上傳檔的章節名對應主檔章節，整段替換。章節名允許「主名-附註」。
           上傳檔有章節但沒有投影片 → 本週沒內容，保留空章節
           上傳檔缺少負責的章節     → 主檔原內容不動並報告
           上傳檔完全沒有章節       → 整份略過並報告
           副主席的 BNI簡介 只換備忘稿含 #BNI台灣發展現況 的那一頁
  組別檔   只換組別頁＝請上台頁之後、第一張會員個人頁之前的所有頁。請上台頁沿用主檔，
           上傳檔第一頁若是請上台頁就丟掉。主檔有標 #固定 的組會交叉比對，不一致就不換。
  會員檔   投影片上的編號優先，其次檔名；換掉該會員在組別章節裡的個人頁。
  換組     組長上傳的組別頁上列有組員編號：會員個人頁目前所在的組沒有列他、
           而剛好另一組有列 → 搬到那組（排在第一位之後，其餘依編號）。判斷不了就不搬並報告。
  來賓     來賓章節保留 #固定 頁，其餘換成來賓名單投影片；有來賓 → 前兩頁顯示，沒有 → 隱藏
  輪替     上週第一組移到最後
  報告     各步驟結果、問題清單、會員成品頁碼、換組清單、含影片／音訊的頁（最後人工檢查）

會員個人頁＝頁面上剛好一個 3 位數會員編號，或帶有本程式的會員標記（cSld name BNI25S-<編號>）。

🔴 公開 repo：報告與 log 裡會員只寫編號，不寫姓名與會員檔名。職務檔與組別檔的檔名不含個資，可以印。

用法：
    R="uv run --with python-pptx --with google-api-python-client --with google-auth-oauthlib --with google-auth python tools/assemble.py"
    $R dry-run --work scratch/assemble/run1        # 從 Drive 抓最新檔案組裝，成品只存本機
    $R local --master m.pptx --uploads 上傳區/ --members 會員/ --guest g.pptx --out o.pptx --report r.txt
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.exc import PackageNotFoundError
from pptx.oxml.ns import qn

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
GUEST_SECTION = "來賓20秒自我介紹"
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
    stem = re.sub(r"\d{8}.*$", "", Path(name).stem).strip(" _-")
    if stem.startswith(GROUP_FILE_PREFIX):
        return "group", stem.removeprefix(GROUP_FILE_PREFIX)
    for role in ROLES:
        if stem.startswith(role) or (role == "三分鐘" and stem.startswith("3分鐘")):
            return "role", role
    return None, None


def match_section(upload_name: str, owned: list[str]) -> str | None:
    if upload_name in owned:
        return upload_name
    base = upload_name.split("-")[0]
    return base if base in owned else None


# ---------------------------------------------------------------- 各步驟

def apply_role(deck: Deck, role: str, path: Path) -> None:
    rep = deck.report
    up = Presentation(str(path))
    secs = sections_of(up)
    if not secs:
        raise MemberProblem("上傳檔沒有任何章節，整份略過（請用 PowerPoint 確認章節還在）")
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
        try:
            new = deck.import_many([(s, (path.name, pages[id(s)])) for s in slides])
        except MemberProblem as exc:
            rep.problems.append(f"{path.name}：章節「{target}」匯入失敗：{exc}（這段沒換）")
            continue
        for i in sec[2]:
            deck.drop(i)
        sec[2] = new
        rep.lines.append(f"  {target}：整段換成上傳檔 {len(slides)} 頁"
                         + ("（空章節，本週沒有內容）" if not slides else ""))
    for o in owned:
        if o not in used:
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
    up = Presentation(str(path))
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
    for i in fixed[:2]:
        el = deck.slide[i]._element
        if names:
            el.attrib.pop("show", None)
        else:
            el.set("show", "0")
    deck.report.lines.append(f"  來賓 {len(names)} 位 → 前兩頁{'顯示' if names else '隱藏'}；來賓名單 {len(new)} 頁")


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

def assemble(master: Path, uploads: list[Path], members: list[MemberFile], guest: Path,
             out: Path) -> tuple[Deck, Report]:
    rep = Report()
    deck = Deck(master, rep)
    for f in sorted(uploads):
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
    rep.lines.append(f"[會員個人檔 {len(members)} 份]")
    apply_members(deck, members)
    auto_move(deck)
    rep.lines.append("[來賓名單投影片]")
    apply_guests(deck, guest)
    rotate_groups(deck)
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


def run_local(master: Path, uploads_dir: Path, members_dir: Path | None, guest: Path,
              out: Path, report_path: Path) -> int:
    uploads = sorted(p for p in uploads_dir.iterdir() if p.suffix.lower() == ".pptx")
    deck, rep = assemble(master, uploads, member_files(members_dir), guest, out)
    text = render_report(deck, rep, out)
    report_path.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 1 if rep.problems else 0


# ---------------------------------------------------------------- Drive

def list_subfolders(svc, folder_id: str) -> list[dict]:
    resp = svc.files().list(
        q=f"'{folder_id}' in parents and trashed = false and mimeType = '{bni25s.FOLDER_MIME}'",
        fields="files(id, name)", pageSize=200,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return resp.get("files", [])


def is_deck(f: dict) -> bool:
    return f["mimeType"] in (bni25s.PPTX_MIME, bni25s.GSLIDES_MIME)


def fetch(svc, cfg: dict, work: Path) -> tuple[Path, Path, Path | None, Path]:
    """抓最新主檔、上傳區根目錄的簡報、最新日期子資料夾的會員檔、來賓名單。"""
    decks = [f for f in bni25s.list_folder(svc, cfg["slides25s_ppt_folder_id"])
             if DATE_RE.match(f["name"].removesuffix(".pptx")) and is_deck(f)]
    if not decks:
        sys.exit("❌ PPT 資料夾找不到 YYYYMMDD.pptx 主檔")
    master_f = max(decks, key=lambda f: f["name"])
    master = work / "master.pptx"
    print(f"主檔：{master_f['name']}")
    bni25s.download(svc, master_f, master)

    uploads_dir = work / "uploads"
    uploads_dir.mkdir()
    ups = [f for f in bni25s.list_folder(svc, cfg["dismember_folder_id"]) if is_deck(f)]
    for f in ups:
        bni25s.download(svc, f, uploads_dir / (f["name"].removesuffix(".pptx") + ".pptx"))
    print(f"上傳區：{len(ups)} 份職務／組別檔")

    dated = [d for d in list_subfolders(svc, cfg["dismember_folder_id"]) if DATE_RE.match(d["name"])]
    members_dir = None
    if dated:
        latest = max(dated, key=lambda d: d["name"])
        members_dir = work / "members"
        members_dir.mkdir()
        mfiles = [f for f in bni25s.list_folder(svc, latest["id"]) if is_deck(f)]
        for f in mfiles:
            bni25s.download(svc, f, members_dir / (f["name"].removesuffix(".pptx") + ".pptx"))
        print(f"會員資料夾：{latest['name']}（{len(mfiles)} 份）")
    else:
        print("會員資料夾：沒有 YYYYMMDD 子資料夾，本次不換會員頁")

    guest = work / "guest.pptx"
    gf = svc.files().get(fileId=cfg["guest_slides_file_id"], fields="id, mimeType",
                         supportsAllDrives=True).execute()
    bni25s.download(svc, gf, guest)
    return master, uploads_dir, members_dir, guest


def cmd_dry_run(args) -> int:
    cfg = bni_config.load("assemble")
    work = Path(args.work)
    work.mkdir(parents=True, exist_ok=False)
    svc = bni25s.get_drive_service()
    master, uploads_dir, members_dir, guest = fetch(svc, cfg, work)
    return run_local(master, uploads_dir, members_dir, guest, work / "assembled.pptx", work / "report.txt")


def cmd_local(args) -> int:
    return run_local(Path(args.master), Path(args.uploads), Path(args.members) if args.members else None,
                     Path(args.guest), Path(args.out), Path(args.report))


def main() -> None:
    p = argparse.ArgumentParser(description="BNI 週會簡報組裝")
    sub = p.add_subparsers(dest="cmd", required=True)
    dr = sub.add_parser("dry-run", help="從 Drive 抓最新檔案組裝，成品只存本機")
    dr.add_argument("--work", required=True, help="工作資料夾（必須是新的）")
    dr.set_defaults(func=cmd_dry_run)
    lo = sub.add_parser("local", help="完全離線：本機檔案組裝")
    lo.add_argument("--master", required=True)
    lo.add_argument("--uploads", required=True, help="職務檔／組別檔所在資料夾")
    lo.add_argument("--members", help="會員個人檔所在資料夾")
    lo.add_argument("--guest", required=True)
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
