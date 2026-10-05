---
name: bnippt
description: 更新 BNI 分會「來賓邀請名單」投影片(pptx)＋「來賓簽到表」(docx)。當使用者要填/換 BNI 來賓名單、做來賓簡報、給一批「姓名/專業別/引薦人」名單時啟用。自動從 Drive 下載範本、投影片每頁最多5位分頁、填表、保留背景與字型；簽到表同步更新日期＋姓名/介紹人欄；兩份各自覆蓋回原 Drive 檔（連結不變）。
---

# BNI 來賓名單投影片＋簽到表更新

分會每週來賓換人就重跑的**週期性重複任務**。一次產兩份、各自覆蓋回原 Drive 檔（連結不變）：

1. **來賓名單投影片**（pptx）
2. **來賓簽到表**（docx）

## 焦點區塊（每次開始先輸出給使用者）
```
焦點：把本次來賓名單填進 BNI 來賓簡報＋簽到表
Now：解析名單 → 確認
Next：從Drive抓兩份範本 → 產檔 → 覆蓋回Drive
Blockers：名單有無缺欄位、簽到表姓名欄要不要含專業別
```

## 🔴 設定一律從 config.json 讀，不寫死

投影片與簽到表的 Google 檔案 ID 是私有資料，只存在**未進版控**的 `config.json`
（範本見 `config.example.json`）。本檔與所有公開檔案都不得出現真實值。

```bash
python3 tools/bni_config.py --flow bnippt
```

需要 `guest_slides_file_id`（投影片）與 `signin_sheet_file_id`（簽到表）。
`guest_slides_parent_folder_id` 僅供人工在 Drive 定位，非必要。
人工開檔網址＝`https://docs.google.com/presentation/d/<guest_slides_file_id>/edit`、
`https://docs.google.com/document/d/<signin_sheet_file_id>/edit`。

⚠️ **本 skill 只能在本機跑**：它依賴 `~/.config/gdrive_api/` 的 OAuth token，
雲端排程 routine 碰不到，不要把這個流程排進 routine。

## 共同事實
- **範本一律從 Drive 下載**（不要用 Downloads／本機副本，那些會被清）。下載＝零 context 成本，且拿到最新版：
  ```bash
  uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
    python tools/gdrive_update.py download --file-id "$FILE_ID" --out <scratch>/xxx
  ```
  （`$FILE_ID` 由你從 config.json 讀出後代入；CLI 不會自己讀設定。）
- 覆蓋回 Drive（保留連結，需先確認過名單）：
  ```bash
  uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
    python tools/gdrive_update.py update --file-id "$FILE_ID" --local <OUT>
  ```
  （憑證放 `~/.config/gdrive_api/`；若首次或 token 失效，先跑 `whoami` 走一次 OAuth。）
- 工作檔一律放 `scratch/`（已被 .gitignore 排除）。
- 名單解析：每列多為 `姓名 / 專業別 / 引薦人`，分隔符混用 `/｜|／`、常有多餘空白 → 逐列拆三欄並 strip。用 markdown 表格請使用者核對（中英夾雜、特殊字如 `＆`、`－`、括號綽號等）。

## 流程
1. **解析名單** → markdown 表格請使用者核對。
2. **確認簽到表姓名欄格式**：預設「只放姓名」（簽到表沒有專業別欄，只是核對簽名用）。若使用者要含專業別再改。
3. **抓兩份範本**到 `scratch/`。
4. **產兩份檔**（腳本見下）。
5. **驗證**（見各節）。
6. **各自覆蓋回 Drive** — 名單在步驟 1 已核對過，這裡**不要再問一次**「要不要覆蓋」。驗證（步驟5）通過就直接執行 `gdrive_update.py update`，一路做到完。

---

## 一、來賓名單投影片（pptx）

- **分頁規則：一頁最多 5 位，貪婪填**（前面各頁填滿 5，最後一頁放餘數）。N=3→[3]、N=4→[4]、N=5→[5]、N=8→[5,3]、N=10→[5,5]、N=11→[5,5,1]。少於 5 位那頁列數＝人數。
- **範本頁數＋每頁列數都不固定**（等於上次成品）。產檔務必：**只留 slide 0 → 砍其餘所有頁 → 複製到 ceil(N/5) 頁**；每頁 table 再用 `set_row_count` 調成該頁人數（1..5）。不能假設固定列數（曾遇到範本 slide0 只有 2 列）。
- ⚠️ **砍頁一定要 `drop_rel`，不能只 `sldIdLst.remove`**：範本非空（已含多頁，即上次成品）時，只從 `sldIdLst` 移除會留下孤兒 slide part，`add_slide` 會重用 slide2/3/4 檔名 → **產出 zip 重複條目的壞檔**。正解：移除前先 `prs.part.drop_rel(sldId.get(qn('r:id')))`。驗證用 `Counter(zip.namelist())` 確認**無重複條目**。
- ⚠️ 塞到 5 列時那頁表格會比 4 列版面擠，產完提醒使用者開檔目視一下行高。
- 表格：每頁一個 **6 欄** 的 table、列數＝該頁人數：
  - `col0` = 來賓N 標籤（連號 來賓1..來賓N）
  - `col1` = 姓名　`col3` = 專業別內容　`col5` = 引薦人
  - `col2`（「專業別」標籤）、`col4`（「介紹人」標籤）**不要動**
- 字型：微軟雅黑（Microsoft YaHei）18pt。背景圖在 slideLayout1／master（非每頁 shape），用同一 layout 複製頁面會自動繼承。每頁有小圖 image2.png（r:embed），複製時要一併搬並重映 rId。
- 環境：`uv run --with python-pptx`；python-pptx 1.0.2 實測 OK。
- **驗證**：頁數正確、每頁 embed 都 resolve、name-cell 字型 = Microsoft YaHei 18pt、**zip 無重複條目**。

### 腳本範本
```python
import copy, math
from pptx import Presentation
from pptx.oxml.ns import qn

SRC = "<scratch>/template.pptx"          # 從 Drive 下載的範本
OUT = "<scratch>/來賓名單投影片_<日期場次>.pptx"

guests = [
    # ("姓名", "專業別", "引薦人"),  ← 填入本次解析後的名單
]

prs = Presentation(SRC)

def duplicate_slide(prs, index):
    src = prs.slides[index]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(shp._element))
    old_to_new = {}
    for rId, rel in src.part.rels.items():
        rt = rel.reltype
        if rt.endswith('slideLayout') or rt.endswith('notesSlide'):
            continue
        old_to_new[rId] = new.part.relate_to(rel._target, rt, rel.is_external)
    for elem in new.shapes._spTree.iter():
        for attr in (qn('r:embed'), qn('r:link')):
            if elem.get(attr) in old_to_new:
                elem.set(attr, old_to_new[elem.get(attr)])
    return new

def set_row_count(table, n):
    """調整 table 的 <a:tr> 數量：多則刪尾列，少則複製最後一列。"""
    tbl = table._tbl
    trs = tbl.findall(qn('a:tr'))
    while len(trs) > n:
        tbl.remove(trs[-1]); trs = tbl.findall(qn('a:tr'))
    while len(trs) < n:
        tbl_tr = copy.deepcopy(trs[-1]); trs[-1].addnext(tbl_tr); trs = tbl.findall(qn('a:tr'))

# 1) 只留 slide 0 當範本，砍掉其餘所有頁（範本頁數不固定）
#    ⚠️ 一定要 drop_rel，否則孤兒 part 會導致 zip 重複條目壞檔
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:]:
    prs.part.drop_rel(sldId.get(qn('r:id')))
    sldIdLst.remove(sldId)
# 2) 複製到本次需要的頁數（一頁最多 5 位，貪婪填）
PER_PAGE = 5
pages = math.ceil(len(guests) / PER_PAGE)
for _ in range(pages - 1):
    duplicate_slide(prs, 0)

def set_cell_text(cell, text):
    tf = cell.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.add_run().text = text
    for extrap in tf.paragraphs[1:]:
        extrap._p.getparent().remove(extrap._p)

for si, slide in enumerate(prs.slides):
    table = next(sh.table for sh in slide.shapes if sh.has_table)
    rows_here = min(PER_PAGE, len(guests) - si * PER_PAGE)
    set_row_count(table, rows_here)
    for r in range(rows_here):
        idx = si * PER_PAGE + r
        name, prof, intro = guests[idx]
        set_cell_text(table.cell(r, 0), f"來賓{idx+1}")
        set_cell_text(table.cell(r, 1), name)
        set_cell_text(table.cell(r, 3), prof)
        set_cell_text(table.cell(r, 5), intro)

prs.save(OUT)

# 驗證
import zipfile
from collections import Counter
dups = {n: c for n, c in Counter(zipfile.ZipFile(OUT).namelist()).items() if c > 1}
assert not dups, f"zip 重複條目（壞檔）: {dups}"
v = Presentation(OUT)
print("slides:", len(v.slides))
for i, s in enumerate(v.slides):
    embeds = [e.get(qn('r:embed')) for e in s.shapes._spTree.iter() if e.get(qn('r:embed'))]
    assert all(r in s.part.rels for r in embeds), f"slide{i} embed 未 resolve"
    t = next(sh.table for sh in s.shapes if sh.has_table)
    for r in range(len(t.rows)):
        print(i, [t.cell(r, c).text for c in (0, 1, 3, 5)])
print("SAVED:", OUT)
```

---

## 二、來賓簽到表（docx）

- 結構：
  - 標題段落（`doc.paragraphs[0]`）內含一個 `日期：YYYY.MM.DD` 的 run → 每次要**改成本次場次日期**（格式維持 `2026.09.09` 這種點分）。
  - 一個 table：**17 列 × 5 欄**。row 0 是表頭 `['', '姓名', '介紹人', '簽名\n中文名稱', '繳費...']`；row 1~16 是編號 1~16。
  - `col0`=編號（不動）　`col1`=姓名　`col2`=介紹人　`col3`/`col4`=現場手簽（留空不動）。
- 每次要做：**改日期** ＋ **把本週來賓填進 col1/col2 的前 N 列**（清掉上週舊資料），其餘列留空。
- 姓名欄預設**只放姓名**（不含專業別）；介紹人欄放引薦人（與 PPT 的 col5 同值，可能含括號綽號備註，原樣保留）。
- 字型：填入的 run 用 Microsoft YaHei、`w:sz` 36（比照現有已填列）。
- 環境：`uv run --with python-docx`。
- **驗證**：重讀 OUT，P0 日期正確、table 前 N 列 col1/col2 正確、其餘列空、font = Microsoft YaHei。

### 腳本範本
```python
from docx import Document
from docx.oxml.ns import qn

SRC = "<scratch>/signin.docx"
OUT = "<scratch>/來賓簽到表_<日期場次>.docx"

NEW_DATE = "日期：2026.09.09"          # ← 改成本次場次日期
rows = [
    # ("姓名", "介紹人"),  ← 填入本次名單（前 N 列）
]

doc = Document(SRC)

hit = False
for p in doc.paragraphs:
    for run in p.runs:
        if run.text.strip().startswith("日期："):
            run.text = NEW_DATE; hit = True
assert hit, "找不到日期 run"

t = doc.tables[0]

def set_cell(cell, text):
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    run = p.add_run(text)
    rpr = run._r.get_or_add_rPr()
    rf = rpr.makeelement(qn('w:rFonts'), {})
    for a in ('w:ascii', 'w:cs', 'w:eastAsia', 'w:hAnsi'):
        rf.set(qn(a), 'Microsoft YaHei')
    rpr.insert(0, rf)
    for tag in ('w:sz', 'w:szCs'):
        e = rpr.makeelement(qn(tag), {}); e.set(qn('w:val'), '36'); rpr.append(e)

for i, (name, intro) in enumerate(rows):
    tr = t.rows[i + 1]
    set_cell(tr.cells[1], name)
    set_cell(tr.cells[2], intro)

doc.save(OUT)

# 驗證
v = Document(OUT)
print("P0:", repr(v.paragraphs[0].text))
for ri, row in enumerate(v.tables[0].rows):
    print(ri, [c.text for c in row.cells])
```
