# bnitools

BNI 分會幾件例行雜務的自動化：

| 工作 | 頻率 | 做什麼 |
|---|---|---|
| **小房間預約狀況** | 每週 | 從預約表單抓「還能約哪幾天」、從彙整試算表抓「誰約了」，組成可直接貼 LINE 群組的訊息 |
| **來賓名單投影片** | 每週 | 把本週來賓（姓名／專業別／引薦人）填進投影片範本，覆蓋回 Google Drive 原檔（連結不變） |
| **來賓簽到表** | 每週 | 同一份來賓名單：更新簽到表 docx 的場次日期＋姓名/介紹人欄，覆蓋回原檔（連結不變） |
| **會員離會同步更新** | 不定期 | 從「會員簽到表」(docx) 與「會員名單」(xlsx) 同步刪除離會會員；docx 直接覆蓋，xlsx 用新分頁保留舊版本；兩邊日期都改成下一個最近的星期三 |

這些都是給 [Claude Code](https://claude.com/claude-code) 用的**專案 skill**：
在這個資料夾裡開 Claude Code，說「本週小房間狀態」「本週的來賓」或「XX 號離會」即可。

---

## 🔴 這是公開 repo，真實資料一律不進版控

分會名稱、Google 表單／試算表／投影片的檔案 ID，全部放在 `config.json`，
而 `config.json` 已被 `.gitignore` 排除。公開檔案裡只有 `config.example.json` 的佔位符。

會這樣做是因為那幾份 Google 檔案的分享設定可能是「知道連結的人可以編輯」——
這種情況下**檔案 ID 就是存取入口**，公開等於把編輯權公告出去。
（分享設定該不該收緊是另一件事，請檔案擁有者處理；repo 這側先做去識別。）

---

## 設定

```bash
cp config.example.json config.json
# 編輯 config.json，填入自己的分會名稱與檔案 ID
```

| 欄位 | 說明 |
|---|---|
| `schema_version` | 目前為 `1`。與工具不符會直接報錯 |
| `chapter_name` / `chapter_short` | 分會全名／訊息標題用的短名 |
| `timezone` | 用來算「今天」與「本週三」，例 `Asia/Taipei` |
| `room_form_id` | 小房間預約表單的 Google 檔案 ID |
| `room_form_responder_url` | 表單填寫網址（貼給會員的那個連結） |
| `room_sheet_id` | 預約彙整試算表的 Google 檔案 ID |
| `guest_slides_file_id` | 來賓名單投影片的 Google 檔案 ID |
| `guest_slides_parent_folder_id` | 投影片所在資料夾 ID（僅供人工定位） |
| `signin_sheet_file_id` | 來賓簽到表 docx 的 Google 檔案 ID |
| `member_roster_doc_id` | 列印用會員表 docx 的 Google 檔案 ID |
| `member_roster_sheet_id` | 會員名單 xlsx 的 Google 檔案 ID |
| `slides25s_ppt_folder_id` | 週會主簡報所在的 Drive 資料夾 ID |
| `slides25s_upload_folder_id` | 25 秒會員檔上傳資料夾 ID（`bni25s` 流程使用） |
| `dismember_folder_id` | 「Dismember-PPT」上傳資料夾 ID（週會組裝使用） |

檢查設定是否齊全（只印欄位名，不回顯值）：

```bash
python3 tools/bni_config.py --flow bni-local
python3 tools/bni_config.py --flow bnippt
python3 tools/bni_config.py --flow bnimembers
```

缺欄位、空白、或還是佔位符，它會直接停下來告訴你缺哪一個 —— 不會猜、不會用空值硬跑。

## Google 授權（一次性，你自己做）

1. Google Cloud Console 建專案
2. 啟用 **Google Drive API** 與 **Google Forms API**
3. OAuth 同意畫面：User Type = External，把自己加為 Test user
4. 建 OAuth 用戶端 ID（類型：桌面應用程式）→ 下載 JSON
5. 改名 `credentials.json` 放到 `~/.config/gdrive_api/`

第一次執行會開瀏覽器要你登入，token 快取在 `~/.config/gdrive_api/token.json`。
兩支工具共用同一份 token（Forms API 的讀取操作接受 `auth/drive` scope）。

⚠️ `credentials.json` 與 `token.json` 是機密，放在家目錄、不在本 repo 內。

## 安裝 commit 前的敏感內容閘門

**第一步：先檢查**，確認你沒有已在使用的 hooks 設定（包含從全域繼承來的）：

```bash
git config --show-origin --get core.hooksPath   # 看有效設定與它來自哪個檔案
ls .git/hooks                                   # 也看一下有沒有自己放的 hook
```

**第二步：確認過沒問題再安裝**（只對這個 repo 生效，不動你的全域設定）：

```bash
git config --local core.hooksPath .githooks
```

（刻意分成兩步，避免整段貼上直接覆蓋掉既有設定。）

之後每次 commit 會自動掃 staged 內容，抓到真實 Google 網址、檔案 ID、email
或憑證結構就擋下來。

- clone 不會自動啟用，接手的人要自己跑上面這行。
- `git commit --no-verify` 可以繞過，所以**首次公開發布前**另外跑：
  ```bash
  python3 tools/check_no_secrets.py --all-tracked
  python3 tools/check_no_secrets.py --history <每一個要發布的 branch 或 tag>
  ```
  `--history` **一次只掃你指定的那一個 ref**，不會順手帶出其他 refs。
  要發布多個 branch／tag 就每個各跑一次，並用明確的 refspec 推送。
  另須**逐檔人工確認**，以及確認 commit 作者 email 是不是你有意公開的身分。
  掃描通過不等於證明乾淨。
- 提示（advisory）是**啟發式**的：只提示同時含數字與大寫的長字串，
  全小寫或純數字的 ID 不會出現在提示裡。它不負責補阻擋規則的漏洞。
- 誤判時在 `.secretscan-allow.json` 加入**精確字串＋理由**。禁止整檔豁免，
  也不得把真實 ID 寫進這個公開的白名單。

## 工具

| 檔案 | 用途 |
|---|---|
| `tools/gforms.py` | 唯讀讀 Google 表單的題目、選項、回覆 |
| `tools/gdrive_update.py` | 下載／就地覆蓋既有 Drive 檔內容／另存副本（fileId／分享連結不變） |
| `tools/bni_config.py` | 私有設定的載入與驗證 |
| `tools/room_dates.py` | 小房間預約的日期計算（skill 與測試共用同一份邏輯） |
| `tools/check_no_secrets.py` | 敏感內容掃描 |
| `tools/bni25s.py` | 25 秒會員頁替換的共用邏輯（`assemble.py` 使用） |
| `tools/assemble.py` | 週會簡報組裝：各職務／各組上傳檔＋會員個人檔＋來賓名單 → 本週主檔（規則見檔案開頭） |

⚠️ `gforms.py get/responses` 與 `gdrive_update.py whoami` 的**輸出**含會員姓名、
email、檔案連結。不要貼進 issue、log 或截圖，也不要把輸出存成 fixture 進版控。

## 測試

```bash
python3 tools/tests/test_room_dates.py
python3 tools/tests/test_check_no_secrets.py
uv run --with google-auth python tools/tests/test_gdrive_update.py
uv run --with python-pptx --with google-api-python-client --with XlsxWriter python tools/tests/test_bni25s.py
uv run --with python-pptx --with google-api-python-client python tools/tests/test_assemble.py
```

前兩個只用標準庫；第三個需要 `google-auth`（測試本身以假物件替換外部 API，
不碰真網路、不彈瀏覽器）。

## 支援邊界

這套工具是照一個特定分會的既有流程做的，不是萬用框架。目前假設：

- 例會在**星期三**，小房間以「哪一個星期三」為預約單位
- 預約表單有「編號與姓名／預約日期／組別或PT／是否需要使用投影機？」這幾題，
  且「預約日期」是選擇題
- 表單掛有 Apps Script，會自動移除已被預約與已過期的日期選項
- 投影片範本每頁一個 6 欄的表格；分頁規則為一頁最多 5 位、貪婪填（列數隨人數調整）
- 簽到表 docx 是「標題段落含 `日期：YYYY.MM.DD` ＋ 一個 17 列 × 5 欄的表格（姓名/介紹人在 col1/col2）」

分會流程不同的話，改 SKILL.md 裡的規則即可，但表格形狀與題目名稱可能要一起調整。

## 雲端排程

> ⚠️ **狀態：尚未驗證。** 以下是設計，還沒有在真實的雲端 routine 上跑過。
> 目前所有測試與流程驗收都是在本機完成的，本機成功不代表雲端可行。

小房間訊息可以掛成 Claude 的雲端排程 routine，每週自動跑一次，本機不用開著。
routine 在雲端沙箱 clone 這個 repo，**碰不到你的本機檔案與 OAuth token**，所以：

- 可預約日期改用 WebFetch 讀表單填寫網址，不走 `gforms.py`
- 設定放在 routine 的 prompt 裡（同一份 schema），不從 repo 讀
- routine 要掛 Google-Drive 連接器才讀得到彙整試算表
- **來賓投影片與簽到表不能在雲端跑**（需要本機 OAuth）

## 分組報告順序輪替（Google Apps Script，全自動）

> ✅ **狀態：2026-09-18 上線，每週二自動執行，已確認自動排程正常運作。**
> 重新部署時仍建議先用程式碼裡的 `DRY_RUN=true` 測過再正式上線。

會員名單 xlsx 現用分頁裡，會員依分類分成好幾組，每次會議後報告順序要整體輪替一次
（第一組排到最後，其餘往前一位）。這件事跟「刪離會會員」（`bnimembers` skill）是
獨立的、每週都要做的固定規則，不需要人判斷，所以改用 **Google Apps Script 排程**
做，不靠 Claude Code 也不靠本機——程式碼在 `apps_script/rotate_category_order.gs`，
部署步驟寫在該檔案開頭的註解。檔案本身維持 `.xlsx` 格式不變（腳本執行時暫時轉成
Google Sheet 編輯，編輯完轉回 xlsx 覆蓋回原檔，暫存副本用完即刪）。

## 週會簡報組裝（本機執行）

> ⏳ **狀態：本機執行，不排程；成品由人工檢查後上傳 Drive（2026-10-27 首次正式使用）。**

主簡報拆成各職務／各組自己的檔，大家編輯後上傳到 Drive「Dismember-PPT」，
`tools/assemble.py` 依章節把它們組回本週主檔。會員個人 25 秒單頁的資料夾每週不同，執行時用 `--members` 給連結；
組長上傳到這個資料夾的組別檔（請上台頁開頭、可夾帶個人頁）也會自動當組別檔替換。
主檔取 PPT 資料夾裡 `YYYYMMDD` 開頭、日期最新的檔（同一天多份取修改時間最新的）；來賓名單投影片沒有比主檔新就不換。
上傳 PDF 會每頁轉成圖片投影片放進去；Canva 等尺寸不同但比例相同的上傳檔會自動等比例縮放；三分鐘、活動組沒有章節時也能判斷。
完整組裝規則寫在 `tools/assemble.py` 開頭。

```bash
R="uv run --with python-pptx --with pymupdf --with google-api-python-client --with google-auth-oauthlib --with google-auth python tools/assemble.py"
$R dry-run --work <新的工作資料夾> --members <本週會員資料夾連結>   # 從 Drive 抓檔組裝，成品與報告只存本機
$R local --master m.pptx --uploads 上傳區/ --members 會員/ [--guest g.pptx] --out o.pptx --report r.txt
```

`config.json` 需要 `slides25s_ppt_folder_id`、`dismember_folder_id`、`guest_slides_file_id`。
報告會列出：各步驟結果、問題清單、會員在成品的頁碼、自動換組清單、含影片／音訊的頁（最後人工檢查）。

**會員頁怎麼換**（共用 `tools/bni25s.py`）：先看投影片上「行首的 3 位數編號」，判斷不出來才用檔名的
3 位數，兩邊不一致時以投影片為準。替換方式等同 PowerPoint「保留來源格式」：會員頁連同版型、
母片、背景圖一起匯入，換過的頁面帶隱藏標記，下週即使是整頁圖片也找得到。

**報告只印會員編號**（公開 repo），不印姓名與會員檔名。

**安全把關**：主檔超過 100 MB，Google 不掃毒，所以在入口擋——會員檔含巨集、嵌入物件（OLE）、
ActiveX、嵌入檔案、連到本機檔案的連結一律拒收（圖表內嵌的 xlsx 資料表放行）。

## Google Apps Script（用 clasp 管理）

| 檔案 | 用途 |
|---|---|
| `apps_script/weekly_upload_folder.gs` | 每週一建立本週上傳資料夾並發通知，週二 17:30 關閉分享 |
| `apps_script/form_date_options.gs` | 舊版預約表單 script，**不用部署**（表單已有別人維護的 script） |
| `apps_script/rotate_category_order.gs` | 會員名單分組順序輪替 |

GAS 跑在 Google 帳號底下，**clone 這個 repo 不會讓它們跑起來**，接手的人要用自己的帳號部署。

```bash
npm install -g @google/clasp
clasp login
```

1. 到 [Apps Script 設定](https://script.google.com/home/usersettings) 開啟 Google Apps Script API。
2. 新建專案後用 `clasp clone <scriptId>` 抓到工作資料夾，把 repo 裡的 `.gs` 內容放進去，再 `clasp push`。
3. 在「專案設定 → 指令碼屬性」填入設定值，並依各檔案開頭說明建立觸發器。

clasp **不會**帶走觸發器和指令碼屬性，一定要重設。`.clasp.json` 不進版控，`~/.clasprc.json` 是登入憑證，不要給別人。

### 通知管道

接手的人可用的工具可能不同，選一個即可。`weekly_upload_folder.gs` 內建 Teams 和 Email，要換其他管道就改 `notify_()`。

| 管道 | 需要什麼 |
|---|---|
| Email | 最簡單，用 `MailApp`。**Google 帳號可能更換，收件人（`NOTIFY_EMAIL`）要手動設定** |
| Teams | 用 Workflows 建立 webhook |
| Slack | Incoming Webhook |
| Google Chat | 空間的 webhook（需要 Google Workspace 帳號） |
| Discord | 頻道 webhook |
| Telegram | Bot API |
| LINE | Messaging API（LINE Notify 已於 2025 年停止服務） |

⚠️ webhook 網址等同發訊息的憑證，只放指令碼屬性，不要寫進程式或貼到公開地方。

## 交接給下一位

```bash
git clone https://github.com/adeleeeee/bnitools.git
```

repo 不會帶走以下項目，要另外處理：

1. **`config.json`**：由前一位私下提供，不要貼在群組或 issue。分會 Drive 的檔案靠裡面的 ID 存取。
2. **Google 授權**：照「Google 授權」段落用自己的帳號設定。
3. **工具環境**：安裝 Claude Code、`uv`、Python；部署 GAS 另需 clasp。
4. **Apps Script**：每週上傳資料夾的 script 在個人帳號下，照「Google Apps Script」段落用自己的帳號部署、設定觸發器，並選擇通知管道。預約表單的 script 不用部署。
   `PARENT_ID` 設成自己 Drive 裡的資料夾，每週資料夾就會開在那裡。
5. **commit 閘門**：照「安裝 commit 前的敏感內容閘門」段落啟用。

**兩個 Drive 要分清楚：** 會員上傳的每週資料夾在部署者的個人 Drive；25 秒簡報、預約表單、來賓名單和週會組裝的檔案在分會 Drive。

⚠️ 每週資料夾在週二 17:30 轉為私人，所以組裝用的 Google 帳號要和部署該 script 的帳號相同。

⚠️ 新的 script 上線後，請前一位刪掉自己那支的觸發器，否則每週會多開一個資料夾。

這個 repo 留在前一位的帳號下，你的修改推不回來。要自己維護的話，請建自己的 repo，再把 `origin` 改過去。
