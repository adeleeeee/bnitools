---
name: bnimembers
description: 處理 BNI 分會會員離會：從「列印用會員簽到表」(docx) 與「會員名單」(xlsx) 兩份檔案同步刪除該會員。當使用者說「XX 號離會」「XX 號會員離開了，更新一下」時啟用。docx 直接覆蓋；xlsx 用複製分頁的方式保留舊版本，不開新檔案。兩份文件內的日期都改成下一個最近的星期三。
---

# BNI 會員離會 → 同步更新會員簽到表＋會員名單

**不定期任務**（會員離會時才跑，不是每週）。用「會編」比對，兩份檔案連動：

1. **會員簽到表**（docx，列印用）：刪除該會員那一列，**直接覆蓋**回 Drive。
2. **會員名單**（xlsx）：**不直接覆蓋**——先複製「目前現用的分頁」成一份新分頁（保留舊資料當備份），在新分頁刪除該會員，設為現用分頁，整份檔案覆蓋回 Drive（舊分頁原樣保留在同一個檔案裡）。

兩份檔案裡都各有一個「日期」欄位，同時改成**下一個最近的星期三**（分會例會日）。

## 焦點區塊（每次開始先輸出給使用者）
```
焦點：把 XX 號離會同步反映到會員簽到表＋會員名單
Now：抓兩份檔案現況 → 找到該會員 → 給使用者看要刪的資料
Next：改日期＋刪列 → 覆蓋回 Drive
Blockers：兩份檔案裡查到的會員資料是否一致（姓名/專業別對不上要停下來問）
```

## 🔴 設定一律從 config.json 讀，不寫死

```bash
python3 tools/bni_config.py --flow bnimembers
```

需要 `member_roster_doc_id`（會員簽到表 docx）與 `member_roster_sheet_id`（會員名單 xlsx）。

⚠️ **只能在本機跑**：依賴 `~/.config/gdrive_api/` 的 OAuth token，雲端排程碰不到。

## 下一個最近的星期三

分會例會固定星期三。今天如果剛好是星期三，算「下一個」（今天以後最近的星期三，不是今天）：

```python
import datetime
today = datetime.date.today()
days_ahead = (2 - today.weekday()) % 7 or 7   # Wed=2；今天是週三則跳到下週三
next_wed = today + datetime.timedelta(days=days_ahead)
```

## 開始前：先找到會員、給使用者看過再動手

先各自抓下兩份檔案（見下），各自找出「編號＝該會員編號」的那一列，**兩邊姓名/專業別要一致**才繼續；不一致就停下來問使用者，不要自己猜哪邊才是對的。

抓檔範本（`$FILE_ID` 從 config.json 讀出後代入）：
```bash
uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
  python tools/gdrive_update.py download --file-id "$FILE_ID" --out scratch/xxx
```
覆蓋回 Drive：
```bash
uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
  python tools/gdrive_update.py update --file-id "$FILE_ID" --local scratch/xxx
```

---

## 一、會員簽到表（docx）

- 結構：`doc.paragraphs[0]`＝標題；`doc.paragraphs[1]`＝`日期：YYYY.MM.DD`（點分格式）。
  一個 table：row 0＝表頭（編號／姓名／專業別／簽名／簽到時間／代理人簽名），row 0 的
  `tblHeader=1`（Word 原生「標題列重複」，每頁自動重出現，2026-09-17 已修好，之後**不要**
  再用手動複製表頭列的方式）；row 1..N＝會員，`col0`＝編號（字串）。
- 找列：`row.cells[0].text.strip() == "編號字串"`。
- 改日期：`paragraphs[1]` 只留第一個 run 改文字，其餘 run 刪掉（避免殘留舊字型片段）。
- 環境：`uv run --with python-docx`。

### 腳本範本
```python
from docx import Document

SRC = "scratch/xxx.docx"
OUT = "scratch/xxx_out.docx"
NEW_DATE = "日期：2026.MM.DD"     # ← 下一個最近的星期三
REMOVE_ID = "131"                  # ← 該會員編號（字串）

doc = Document(SRC)

date_p = doc.paragraphs[1]
assert date_p.text.startswith("日期："), date_p.text
for run in date_p.runs[1:]:
    run._r.getparent().remove(run._r)
date_p.runs[0].text = NEW_DATE

table = doc.tables[0]
target = next(r for r in table.rows if r.cells[0].text.strip() == REMOVE_ID)
removed_text = [c.text for c in target.cells]   # 存起來給使用者核對
target._tr.getparent().remove(target._tr)

doc.save(OUT)

# 驗證：重讀 OUT，日期正確、編號已消失、列數少 1
```

---

## 二、會員名單（xlsx）

- 一個檔案裡有多個分頁，**現用分頁＝ `wb.active`**（不要假設固定分頁名，分頁名每次
  换成新的會議日期）。舊分頁（例如上次的 `1150804`）原樣保留，不要動。
- 現用分頁結構：row1＝標題橫幅，`F1`＝日期備註（格式 `YYYY.MM.DD.update`，跟簽到表的
  點分格式一致，但多了 `.update` 字尾）；row2＝欄名（會編／姓名／英文名字／專業別／
  三大服務項目／筆記，實際資料從 col B 開始，col A 是空白版面欄）；row3 起是會員，
  但**中間夾雜「分類標題」的合併儲存格列**（`B:F` merge，整列只有一個分類名稱文字，
  沒有會編）。
- 找列：逐列比對 `ws.cell(r, 2).value`；會編是**數字**存的（`131.0` 是 float，不是字串
  `"131"`），比對時用 `float(輸入的編號)`，不要用字串比對。
- 🔴 **`ws.delete_rows()` 不會自動位移合併儲存格範圍**——刪列前記錄所有 `merged_cells.ranges`、
  先 `unmerge_cells` 全部拆掉，刪完列後把 **min_row 大於被刪列**的範圍整體 `-1`
  再重新 `merge_cells`。忘記做這步，分類標題會合併到錯的列。
- 新分頁命名＝下一個最近的星期三，格式 `YYYYMMDD`（例：`20260923`）。用
  `wb.copy_worksheet()` 複製現用分頁 → 改標題 → `wb.move_sheet()` 移到最前面 →
  `wb.active =` 該分頁 index。
- 環境：`uv run --with openpyxl`。

### 腳本範本
```python
import openpyxl

SRC = "scratch/xxx.xlsx"
OUT = "scratch/xxx_out.xlsx"
NEW_SHEET = "20260923"          # ← 下一個最近的星期三，YYYYMMDD
NEW_DATE_NOTE = "2026.09.23.update"
REMOVE_ID = 131.0                # ← 該會員編號（float）

wb = openpyxl.load_workbook(SRC)
old_ws = wb.active
old_name = old_ws.title

new_ws = wb.copy_worksheet(old_ws)
new_ws.title = NEW_SHEET
wb.move_sheet(NEW_SHEET, offset=-(wb.sheetnames.index(NEW_SHEET) - wb.sheetnames.index(old_name)))
wb.active = wb.sheetnames.index(NEW_SHEET)

assert "update" in str(new_ws.cell(1, 6).value)
new_ws.cell(1, 6).value = NEW_DATE_NOTE

target_row = next(r for r in range(1, new_ws.max_row + 1) if new_ws.cell(r, 2).value == REMOVE_ID)
removed = [new_ws.cell(target_row, c).value for c in range(1, new_ws.max_column + 1)]  # 給使用者核對

merged_ranges = [(r.min_row, r.min_col, r.max_row, r.max_col) for r in new_ws.merged_cells.ranges]
for min_row, min_col, max_row, max_col in merged_ranges:
    new_ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

new_ws.delete_rows(target_row, 1)

for min_row, min_col, max_row, max_col in merged_ranges:
    if min_row > target_row:
        min_row -= 1
        max_row -= 1
    new_ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

wb.save(OUT)

# 驗證：新分頁列數少1、編號已消失、舊分頁 old_name 列數/內容完全不變、
#      合併儲存格範圍逐一核對已正確位移、wb.active 是新分頁
```

---

## 跟「分組報告順序輪替」是兩件獨立的事

會員名單 xlsx 現用分頁裡的分類分組順序，每週要整體輪替一次（跟這個 skill 的
「刪離會會員」無關，不是每次跑這個 skill 都要輪替）。那件事改用 Google Apps Script
排程做，不歸這個 skill 管，見 `apps_script/rotate_category_order.gs` 與 README
「分組報告順序輪替」一節。這個 skill **只處理刪會員＋改日期**，不要動分類順序。

## 收尾

- 兩邊都改完、驗證過再覆蓋回 Drive（各自 update，`--file-id` 用對應的
  `member_roster_doc_id` / `member_roster_sheet_id`）。
- 跟使用者報告：刪了哪個會員（姓名／編號）、日期改成哪天、xlsx 新分頁叫什麼名字、
  舊分頁還在不在。
- `scratch/` 裡的中間檔含真實會員姓名，處理完不用主動刪，但別把這些檔案的內容貼進
  commit message 或任何要進版控的地方。
