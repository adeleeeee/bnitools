# bnitools

BNI 分會兩件每週雜務的自動化：

| 工作 | 做什麼 |
|---|---|
| **小房間預約狀況** | 從預約表單抓「還能約哪幾天」、從彙整試算表抓「誰約了」，組成可直接貼 LINE 群組的訊息 |
| **來賓名單投影片** | 把本週來賓（姓名／專業別／引薦人）填進投影片範本，覆蓋回 Google Drive 原檔（連結不變） |

兩件事都是給 [Claude Code](https://claude.com/claude-code) 用的**專案 skill**：
在這個資料夾裡開 Claude Code，說「本週小房間狀態」或「本週的來賓」即可。

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

檢查設定是否齊全（只印欄位名，不回顯值）：

```bash
python3 tools/bni_config.py --flow bni-local
python3 tools/bni_config.py --flow bnippt
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

**先確認你沒有已在使用的 hooks 設定**，這行會覆蓋掉它：

```bash
git config --local --get core.hooksPath   # 有輸出代表已經設過，先確認要不要保留
ls .git/hooks                             # 也看一下有沒有自己放的 hook
git config --local core.hooksPath .githooks
```

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
| `tools/gdrive_update.py` | 就地覆蓋既有 Drive 檔內容（fileId／分享連結不變） |
| `tools/bni_config.py` | 私有設定的載入與驗證 |
| `tools/room_dates.py` | 小房間預約的日期計算（skill 與測試共用同一份邏輯） |
| `tools/check_no_secrets.py` | 敏感內容掃描 |

⚠️ `gforms.py get/responses` 與 `gdrive_update.py whoami` 的**輸出**含會員姓名、
email、檔案連結。不要貼進 issue、log 或截圖，也不要把輸出存成 fixture 進版控。

## 測試

```bash
python3 tools/tests/test_room_dates.py
python3 tools/tests/test_check_no_secrets.py
uv run --with google-auth python tools/tests/test_gdrive_update.py
```

前兩個只用標準庫；第三個需要 `google-auth`（測試本身以假物件替換外部 API，
不碰真網路、不彈瀏覽器）。

## 支援邊界

這套工具是照一個特定分會的既有流程做的，不是萬用框架。目前假設：

- 例會在**星期三**，小房間以「哪一個星期三」為預約單位
- 預約表單有「編號與姓名／預約日期／組別或PT／是否需要使用投影機？」這幾題，
  且「預約日期」是選擇題
- 表單掛有 Apps Script，會自動移除已被預約與已過期的日期選項
- 投影片範本每頁一個 4 列 × 6 欄的表格，每頁放 4 位來賓

分會流程不同的話，改 SKILL.md 裡的規則即可，但表格形狀與題目名稱可能要一起調整。

## 雲端排程

> ⚠️ **狀態：尚未驗證。** 以下是設計，還沒有在真實的雲端 routine 上跑過。
> 目前所有測試與流程驗收都是在本機完成的，本機成功不代表雲端可行。

小房間訊息可以掛成 Claude 的雲端排程 routine，每週自動跑一次，本機不用開著。
routine 在雲端沙箱 clone 這個 repo，**碰不到你的本機檔案與 OAuth token**，所以：

- 可預約日期改用 WebFetch 讀表單填寫網址，不走 `gforms.py`
- 設定放在 routine 的 prompt 裡（同一份 schema），不從 repo 讀
- routine 要掛 Google-Drive 連接器才讀得到彙整試算表
- **來賓投影片不能在雲端跑**（需要本機 OAuth）

## 交接給下一位

clone 或下載這個 repo → 填自己的 `config.json` → 用自己的 Google 授權。
前一位使用者的 `config.json` 未進版控，不會跟著過來，也不會自動備份。

若你以前曾把同名的 `bni` / `bnippt` skill 安裝到全域（`~/.claude/skills/`），
請確認那份是否為不同版本，避免兩份規則同時存在。
