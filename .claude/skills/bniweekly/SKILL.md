---
name: bniweekly
description: 組 BNI 分會每週的週會主簡報。當使用者說「組本週週會簡報」「更新這週的簡報」「週會簡報」並給本週會員資料夾的 Drive 連結時啟用。自動抓上週主檔、領導團隊上傳的職務檔／組別檔、本週會員個人檔、來賓名單，組成本週主檔；使用者檢查過後上傳成 YYYYMMDD.pptx 新檔到 PPT 資料夾。
---

# BNI 週會簡報組裝

每週一次。組裝規則全部寫在 `tools/assemble.py` 開頭，這份 skill 只管「怎麼跑、怎麼回報、怎麼上傳」。

## 焦點區塊（每次開始先輸出給使用者）
```
焦點：組出本週週會主簡報並上傳
Now：抓檔 → 組裝
Next：回報給使用者檢查 → 上傳成 YYYYMMDD.pptx
Blockers：會員資料夾連結、有沒有職務／組別沒交新檔
```

## 🔴 設定一律從 config.json 讀，不寫死

```bash
python3 tools/bni_config.py --flow assemble
```

需要 `slides25s_ppt_folder_id`（主檔所在的 PPT 資料夾）、`dismember_folder_id`（領導團隊上傳區）、
`guest_slides_file_id`（來賓名單投影片）。本檔與所有公開檔案都不得出現真實 ID、會員姓名。

⚠️ 只能在本機跑（依賴 `~/.config/gdrive_api/` 的 OAuth token）。
⚠️ 每週會員資料夾到截止時間會轉為私人，所以跑組裝的 Google 帳號要和建立那個資料夾的帳號相同。

## 使用者要給什麼

只有一樣：**本週會員資料夾的 Drive 連結**（每週不同）。沒給就問，不要猜、不要沿用上週的。

其他都自動抓：

| 來源 | 怎麼找 |
|---|---|
| 上週主檔（當底） | PPT 資料夾裡檔名 `YYYYMMDD` 開頭、日期最新的簡報 |
| 職務檔／組別檔 | 領導團隊上傳區，每個職務／組別只取檔名日期最新、且晚於主檔的一份 |
| 會員個人檔 | 使用者給的連結 |
| 來賓名單 | 來賓名單投影片；修改時間沒有晚於主檔＝本週沒來賓 |

**有來賓的週次，要先跑完 `bnippt` 再組裝**，否則來賓章節會被當成沒來賓而隱藏。
開始前問一句「本週來賓名單更新了嗎？」即可。

## 流程

1. **組裝**（工作資料夾必須是新的，用 `scratch/assemble/run_<MMDD><字母>`）：
   ```bash
   uv run --with python-pptx --with pymupdf --with google-api-python-client --with google-auth-oauthlib --with google-auth \
     python tools/assemble.py dry-run --work scratch/assemble/run_<MMDD>a --members "<本週會員資料夾連結>"
   ```
   成品 `assembled.pptx` 與 `report.txt` 存在工作資料夾，這一步不會碰 Drive 上的任何檔。
2. **讀 `report.txt`**，照下面格式回報。
3. **等使用者檢查成品**。使用者會開 `assembled.pptx` 逐頁看，尤其是有影片的頁。給完整檔案路徑。
4. **使用者說可以才上傳**（見「上傳」）。
5. 上傳後附一則 LINE 訊息（見「LINE 訊息」）。

## 回報格式（極簡）

使用者拿成品逐頁對，只要最少的字。

- **會員頁**：每行 `編號 名字 → 第 N 頁`，有影片才在後面加「（有影片）」。
  名字用使用者在對話裡貼的名單，順序也照那份名單；使用者沒貼名單就只寫編號。
  （`report.txt` 只有編號，這是刻意的：公開 repo 不留姓名。）
- **一定列出本週沒交新檔、沿用上週的職務與組別**。這是使用者要去催的，要顯眼。
- **問題清單**：`report.txt` 有列才講，原文照列。
- 不要貼整份報告、不用表格、不跟上週比。輪替、母片、主檔原本就有的影片都不提。

## 上傳

- 檔名＝`<本週例會日期 YYYYMMDD>.pptx`，例會在星期三。**上傳前把檔名講出來讓使用者確認**；
  使用者指定別的檔名就照字面用。
- 一律上傳成**新檔**，不覆蓋上週主檔：
  ```bash
  uv run --with google-api-python-client --with google-auth-oauthlib --with google-auth \
    python tools/gdrive_update.py create --folder-id "$PPT_FOLDER_ID" \
      --local scratch/assemble/run_<MMDD>a/assembled.pptx --name "<YYYYMMDD>.pptx"
  ```
  （`$PPT_FOLDER_ID`＝config.json 的 `slides25s_ppt_folder_id`，由你讀出後代入。）
- 資料夾裡已有同名檔時指令會停下來。這時問使用者：要覆蓋那一份（`update --file-id`），還是換檔名。不要自己決定。
- 檔名一定要 `YYYYMMDD` 開頭，下週組裝才會拿它當底。

## 🔴 主檔上傳後還要改

**不要再跑一次完整組裝**——會多輪替一次組別順序，而且上傳區的檔已經不比主檔新，會全部被略過。

改法：從 Drive 下載那份主檔，只動要改的章節（例如只換來賓、只換某個職務的章節），
再照使用者說的存成新檔或覆蓋。存新檔還是覆蓋、檔名是什麼，動手前用一句話確認。

## LINE 訊息

上傳後主動附一則可直接貼 LINE 的純文字（不用 markdown）：

```
【MM/DD 週會簡報】25秒會員頁已更新
編號 名字 第N頁
編號 名字 第N頁（有影片）
組別頁已更新：<組名>、<組名>
請大家確認自己的頁面，有問題再跟我說
```

## 常見狀況

| 狀況 | 處理 |
|---|---|
| 上傳檔沒有日期、或同一職務同一天交了兩份 | 程式會略過並寫進報告。請使用者決定用哪一份，本機改好檔名後用 `assemble.py local` 重組 |
| 上傳檔檔名看不出職務（例如沒寫「三分鐘」） | 同上，本機改名後重組 |
| 上傳檔是 PDF | 正常，每頁會轉成圖片投影片。不能編輯、沒有影片是預期結果，不用當問題回報 |
| 上傳檔尺寸不同（Canva） | 比例相同會自動縮放；比例不同整份略過並寫進報告 |
| 成品很大（數百 MB） | 正常，上傳要等幾分鐘 |
