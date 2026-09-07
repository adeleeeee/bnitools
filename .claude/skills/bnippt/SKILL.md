---
name: bnippt
description: 更新 BNI 分會「來賓邀請名單」投影片(pptx)。當使用者要填/換 BNI 來賓名單、做來賓簡報、給一批「姓名/專業別/引薦人」名單要排進投影片時啟用。自動從 Drive 下載範本、每頁4位分頁、填表、保留背景與字型，再覆蓋回原 Drive 檔（連結不變）。
---

# BNI 來賓名單投影片更新

分會「來賓邀請名單」投影片，是**週期性重複任務**：每次來賓換人就重跑。

## 焦點區塊（每次開始先輸出給使用者）
```
焦點：把本次來賓名單填進 BNI 來賓簡報
Now：解析名單 → 確認
Next：從Drive抓範本 → 產檔 → 覆蓋回Drive
Blockers：名單有無缺欄位
```

## 🔴 設定一律從 config.json 讀，不寫死

投影片的 Google 檔案 ID 是私有資料，只存在**未進版控**的 `config.json`
（範本見 `config.example.json`）。本檔與所有公開檔案都不得出現真實值。

```bash
python3 tools/bni_config.py --flow bnippt
```

需要 `guest_slides_file_id`（`guest_slides_parent_folder_id` 僅供人工在 Drive 定位，非必要）。
人工開檔網址＝`https://docs.google.com/presentation/d/<guest_slides_file_id>/edit`。

⚠️ **本 skill 只能在本機跑**：它依賴 `~/.config/gdrive_api/` 的 OAuth token，
雲端排程 routine 碰不到，不要把投影片流程排進 routine。

## 關鍵事實
- **範本一律從 Drive 下載**（不要用 Downloads／本機副本，那些會被清）。下載＝零 context 成本，且拿到最新版：
  ```bash
  uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
    python tools/gdrive_update.py download --file-id "$GUEST_SLIDES_FILE_ID" --out <scratch>/template.pptx
  ```
  （`$GUEST_SLIDES_FILE_ID` 由你從 config.json 讀出後代入；CLI 不會自己讀設定。）
- **範本頁數不固定**（上次填幾位就幾頁）。所以產檔時務必：**只留 slide 0 當範本 → 砍掉其餘所有頁 → 再複製到本次需要的頁數**。以 slide 0 為準（結構乾淨，表格寬度正常）。
- ⚠️ **砍頁一定要 `drop_rel`，不能只 `sldIdLst.remove`**：範本非空（已含多頁，即上次成品）時，只從 `sldIdLst` 移除會留下孤兒 slide part，`add_slide` 會重用 slide2/3/4 檔名 → **產出 zip 重複條目的壞檔**（PowerPoint 可能開得起來但不乾淨）。正解：移除前先 `prs.part.drop_rel(sldId.get(qn('r:id')))`。驗證時務必用 `Counter(zip.namelist())` 確認**無重複條目**。
- 表格：每頁一個 **4 列 × 6 欄** 的 table，欄位：
  - `col0` = 來賓N 標籤（連號 來賓1..來賓N）
  - `col1` = 姓名　`col3` = 專業別內容　`col5` = 引薦人
  - `col2`（「專業別」標籤）、`col4`（「介紹人」標籤）**不要動**
- 字型：微軟雅黑（Microsoft YaHei）18pt。背景圖在 slideLayout1／master（非每頁 shape），用同一 layout 複製頁面會自動繼承。每頁有小圖 image2.png（r:embed），複製時要一併搬並重映 rId。
- 環境：`uv run --with python-pptx`；python-pptx 1.0.2 實測 OK。工作檔一律放 `scratch/`（已被 .gitignore 排除）。

## 流程
1. **解析名單**：每列格式多為 `姓名 / 專業別 / 引薦人`，但分隔符混用 `/｜|／`、常有多餘空白 → 逐列拆成三欄並 strip。用 markdown 表格列出結果請使用者核對（尤其中英夾雜、特殊字如 `＆`、`－`）。
2. **抓範本**：跑上面的 `download` 指令把 Drive 原檔下載到 `scratch/`。
3. **產檔**：跑下方腳本（改 `SRC`=下載的範本、`guests`、`OUT`=含日期場次的檔名）。邏輯＝留 slide0→砍其餘→複製到 ceil(N/4) 頁→填 col0/1/3/5。
4. **驗證**：頁數正確、每頁 embed 都 resolve、name-cell 字型 = Microsoft YaHei 18pt、**zip 無重複條目**（`Counter(ZipFile(OUT).namelist())`）。
5. **覆蓋回 Drive**（保留連結，需先確認過名單）：
   ```bash
   uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
     python tools/gdrive_update.py update --file-id "$GUEST_SLIDES_FILE_ID" --local <OUT>
   ```
   （憑證放 `~/.config/gdrive_api/`；若首次或 token 失效，先跑 `whoami` 走一次 OAuth。）

## 腳本範本
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

# 1) 只留 slide 0 當範本，砍掉其餘所有頁（範本頁數不固定）
#    ⚠️ 一定要 drop_rel，否則孤兒 part 會導致 zip 重複條目壞檔
sldIdLst = prs.slides._sldIdLst
for sldId in list(sldIdLst)[1:]:
    prs.part.drop_rel(sldId.get(qn('r:id')))
    sldIdLst.remove(sldId)
# 2) 複製到本次需要的頁數
pages = math.ceil(len(guests) / 4)
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
    for r in range(4):
        idx = si * 4 + r
        if idx < len(guests):
            name, prof, intro = guests[idx]
            set_cell_text(table.cell(r, 0), f"來賓{idx+1}")
            set_cell_text(table.cell(r, 1), name)
            set_cell_text(table.cell(r, 3), prof)
            set_cell_text(table.cell(r, 5), intro)
        else:
            # 最後一頁不滿 4 位：清空多餘列
            for c in (0, 1, 3, 5):
                set_cell_text(table.cell(r, c), "")

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
