---
name: bni
description: 產出 BNI 分會「小房間預約狀況」的 LINE 群組訊息。當使用者說「bni 本週小房間狀態」「小房間預約狀況」「這週誰借小房間」，或要在 LINE 群組公告小房間可預約日期時啟用。自動抓表單剩餘可預約日期＋試算表已預約名單，組成可直接複製貼上的 LINE 純文字訊息。
---

# 小房間預約狀況 → LINE 訊息

分會每週三例會可借小房間。這是**每週重複任務**：每週在 LINE 群組貼一次狀態＋預約連結。

搭配 skill [[bnippt]]（本週來賓投影片）使用，但**兩者各自獨立**：
- 「bni 本週小房間狀態」→ 本 skill，產出 LINE 訊息
- 「bnippt 本週的來賓」→ bnippt skill，產出投影片

## 焦點區塊（每次開始先輸出給使用者）
```
焦點：產出本週小房間狀況的 LINE 訊息
Now：抓表單剩餘日期 → 抓試算表預約名單
Next：交叉驗證 → 組訊息給使用者複製
Blockers：兩邊資料對不起來時要標出來
```

## 🔴 設定一律從 config.json 讀，不寫死

分會名稱、表單網址、檔案 ID 全部是私有資料，只存在**未進版控**的 `config.json`
（範本見 `config.example.json`）。本檔與所有公開檔案都不得出現真實值。

開工前先驗設定，缺什麼它會直接告訴你：

```bash
python3 tools/bni_config.py --flow bni-local     # 雲端用 --flow bni-cloud
```

需要的欄位：

| 欄位 | 用途 | 本機 | 雲端 |
|---|---|---|---|
| `chapter_short` | 訊息標題 | ✅ | ✅ |
| `timezone` | 算「今天／本週三」 | ✅ | ✅ |
| `room_form_id` | 用 API 讀表單選項 | ✅ | — |
| `room_form_responder_url` | 貼 LINE 的連結；雲端用來 WebFetch | ✅ | ✅ |
| `room_sheet_id` | 讀彙整試算表 | ✅ | ✅ |

雲端流程**不需要也不得要求** `room_form_id` 或任何投影片欄位。

## 關鍵事實

- 表單後面掛的 Apps Script 會把「已被預約」與「已過期」的選項自動移除，所以
  **表單上還看得到的日期＝還能預約的日期**，不用自己算，也不必驗證是不是星期三。
- ⚠️ **不要碰表單本身**：題目／選項由那支 Apps Script 維護，用 API 去寫會打架。本 skill 只讀。
  Apps Script 也讀不到（綁定式腳本不在 Drive 檔案清單裡，且需要 Apps Script API）。
  要改表單邏輯請找表單擁有者。

### 資料來源分工（不要弄反）
| 要什麼 | 去哪拿 | 為什麼 |
|---|---|---|
| **還能預約哪幾天** | 表單的「預約日期」選項（取法見流程 1） | Apps Script 即時維護，最權威 |
| **誰約了、哪一組、要不要投影機** | 彙整試算表（Drive MCP `read_file_content`） | 表單看不到已預約的人；試算表一次讀完最省 |

備援：`tools/gforms.py responses` 也讀得到全部回覆（含組別、投影機、完整時間戳記），
試算表壞掉或要交叉對帳時用（僅本機）。

### 🖥️ 執行環境：本機 vs 雲端排程
唯一差別是**流程 1 怎麼拿可預約日期**：

| | 本機 | 雲端排程 routine |
|---|---|---|
| 可預約日期 | `tools/gforms.py options`（OAuth） | `WebFetch` 表單填寫網址 |
| 為什麼不同 | — | 雲端沒有 `~/.config/gdrive_api/` 的 token，gforms.py 一定失敗 |
| 試算表 | Drive MCP | Drive MCP（routine 要掛 Google-Drive 連接器） |

兩種取法**在本機比對過結果一致**（同一天分別用 API 與 WebFetch 取得相同選項）。
⚠️ 雲端沙箱內尚未實測，本機一致不等於雲端已驗證。

## 流程

1. **抓可預約日期**

   本機：
   ```bash
   uv run --with google-auth --with google-auth-oauthlib python tools/gforms.py \
     options --form-id "$ROOM_FORM_ID" --question 預約日期
   ```
   （`$ROOM_FORM_ID` 由你從 config.json 讀出後代入；CLI 不會自己讀設定。）
   若報「Forms API 尚未啟用」，照訊息裡的連結請使用者去啟用（不能代做）。

   雲端排程：`WebFetch` `room_form_responder_url`，問「預約日期這題有哪些選項，只列選項」。
   雲端**不要**跑 gforms.py。

2. **抓預約名單**
   Google Drive MCP `read_file_content`，fileId ＝ `room_sheet_id`。
   欄位：`時間戳記｜電子郵件地址｜編號與姓名｜預約日期｜組別或PT｜預約使用人數｜是否需要使用投影機？`

3. **整理資料** —— 寫成 payload JSON 後交給計算模組，**不要自己心算日期**：

   ```bash
   python3 tools/room_dates.py --flow bni-local --input <scratch>/payload.json
   ```

   payload 形狀：
   ```json
   {
     "bookings": [
       {"timestamp": "2030/7/6 上午 11:08:53", "date": "10/14",
        "who": "001 王小明", "group": "某某分享", "projector": true}
     ],
     "available": ["9/30", "11/18"]
   }
   ```

   它會處理掉三個最容易翻車的地方：依設定時區取「今天」、同日多筆選最新（比到秒）、
   日期一律正規化成 date 再比對（`9/3` 與 `09/03` 視為同一天）。

   🔴 **抓取失敗一律中止，不要傳空陣列**。空陣列代表「確實查到沒有」，不是「查不到」。

4. **交叉驗證**：模組回傳的 `warnings.clashes` 應為空。
   有值＝有人約了但表單沒關掉（或反之）→ **在訊息外面標給使用者看，不要吞掉**，
   訊息照發但註明哪一天有疑慮。`warnings.duplicated` 同理。

   ### 🔴 遇到 `AmbiguousDate` 時的優先順序（兩種環境都適用）

   表單選項只有 `M/D`、沒有時間戳，今年已過的日期無法判斷是「過期殘留」還是「明年場次」，
   模組會拋 `AmbiguousDate`。此時：

   - **一律停止，不產出訊息。** 這條優先於「雲端要直接輸出」。
   - 本機：把有問題的日期列給使用者，請他確認年份。
     - 確認是**未來**的場次 → 在 payload 的 `available` 改成完整日期 `YYYY-MM-DD` 再跑一次。
     - 確認是**已過期**的殘留選項 → 從 `available` 移除。
       ⚠️ 不要把過期日期改寫成完整日期送回去——模組會再次中止，
       因為知道年份不代表那天還能預約。
   - 雲端：**只輸出失敗原因與需要確認的日期**，不要猜年份、不要略過該筆、
     也不要產出「正常」的預約公告。等使用者下次手動處理。

5. **輸出 LINE 訊息**：純文字 code block（使用者要直接複製貼上，
   **不要套引言 `>`、不要加 markdown 粗體**）。

6. 使用者若說內容 OK，這件事就結束——**不需要也不能代發 LINE**。
   雲端排程沒有人可以問，直接把訊息當最終輸出印出來就結束，同樣**絕不代發**。

## 訊息呈現規則

- **姓名格式不統一**（大小寫、編號在前在後）→ 顯示時統一成 `編號 姓名`（數字在前、空一格）。
- **投影機**只在「本週」那一行標註 `（需投影機）`；接下來的預約不標，訊息才不會太長。
- **不要逐一列出「4 場之後」還有哪些天被預約**。訊息外面一句帶過就好：
  **「沒列在『尚可預約』裡的日期就是已經被預約了。」** 列一串沒人要看。

## LINE 訊息範本

```
【<chapter_short> 小房間預約狀況】<今天 M/D> 更新

■ 本週 <M/D>（三）
　<編號 姓名>｜<組別或PT><（需投影機）>

■ 接下來的預約
　<M/D>　<編號 姓名>｜<組別或PT>
　（最多 4 筆）

■ 尚可預約的日期
　<以「、」串接>

要借小房間請填表單（先填先贏）：
<room_form_responder_url>
```

變化：
- 本週沒人預約 → 該區塊改成一行 `　目前無人預約，歡迎登記`
- 接下來都沒預約 → 整個「接下來的預約」區塊拿掉
- 沒有可預約日期 → 改成 `　目前已額滿，請洽表單負責人開放新日期`

## 雲端排程（Claude 沒開也會跑）

routine 跑在 Anthropic 雲端沙箱，clone 本 repo，**碰不到本機檔案與 OAuth token**。

- 流程 1 走 WebFetch，不要碰 gforms.py。
- routine 必須掛 **Google-Drive 連接器**才讀得到彙整試算表。
- `config.json` 不在 repo 裡（被 .gitignore 擋掉），所以 routine 的設定放在 prompt 的設定區塊。
  完整做法：

  ```bash
  # 1) 把 prompt 裡的設定區塊原樣存成暫存檔（不要存進 repo 根目錄）
  mkdir -p /tmp/bni && cat > /tmp/bni/config.json <<'JSON'
  { …照 config.example.json 的 schema，填 prompt 給的值… }
  JSON

  # 2) 驗證設定（缺欄位會直接告訴你缺哪個）
  python3 tools/bni_config.py --flow bni-cloud --config /tmp/bni/config.json

  # 3) 算日期
  python3 tools/room_dates.py --flow bni-cloud --config /tmp/bni/config.json \
      --input /tmp/bni/payload.json
  ```

- 兩個設定來源互斥：雲端環境不會有 repo config.json，若兩者同時存在，工具會直接報錯。
  **暫存檔放 repo 外面**（如 `/tmp/bni/`），避免不小心被當成 repo 設定或被 commit。
- 雲端沒有互動對象：不要問問題、不要等確認，直接輸出訊息＋警示。
  但這只適用於**流程正常完成**的情況：**任何前置驗證失敗、資料取得失敗或日期處理失敗
  （設定缺漏、表單／試算表抓不到、`AmbiguousDate`、無效輸入…），一律只回報原因並停止，
  不產出正常公告。**
- ⚠️ bnippt **不能**在雲端跑（需要本機 OAuth），不要把投影片流程排進 routine。
