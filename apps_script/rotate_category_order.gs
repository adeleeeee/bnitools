/**
 * BNI 分組報告順序輪替（每週全自動，跑在 Google 雲端，不需要本機、不需要 Claude）。
 *
 * 背景：會員名單 xlsx 現用分頁裡的會員，用「合併儲存格的分類標題列」分成 12 組。
 * 每次會議後，順序要整體輪替一次：第一組搬到最後，其餘全部往前一位。這個檔案
 * 一直維持 .xlsx 格式（不轉成原生 Google Sheet），做法是：執行時先暫時把它轉一份
 * 原生 Google Sheet 副本、用 SpreadsheetApp 編輯、編輯完再轉出成 xlsx 覆蓋回原檔、
 * 刪掉暫存副本。
 *
 * ✅ 2026-09-18 已實際部署上線，每週二由時間觸發器自動執行，已確認運作正常。
 * 以下是部署時踩過的坑，重新部署或修改時注意：
 *   1. 先用 DRY_RUN=true 跑過幾次，人工核對輸出的測試檔沒問題，才切成正式模式。
 *   2. 已知一個命名差異（已在下面程式碼修正）：Apps Script 的 Drive 進階服務裡，
 *      刪除方法叫 `Drive.Files.remove`，不是 REST 文件寫的 `delete`（create/update/get
 *      都跟 REST 同名，只有這個特例）。如果之後又遇到「XXX is not a function」，
 *      先去 Apps Script 編輯器左側「Services」點進 Drive API 看它實際列出的方法名稱，
 *      不要照 REST 文件的名字硬猜。
 *   3. xlsx → Sheets → xlsx 這個轉換來回，Google 不保證 100% 還原所有 Excel
 *      專屬格式（少數字型/框線細節可能跑掉），正式上線前開結果檔目視確認一次。
 *
 * 部署步驟（使用者自己動手，一次性）：
 *   1. https://script.google.com/ → 新增專案，把這整份檔案貼進去
 *   2. 把下面 FILE_ID 改成 config.json 裡的 member_roster_sheet_id 真實值
 *      （只改在 Apps Script 編輯器裡，不要把真實 ID 存回這個 repo 檔案）
 *   3. 左側「服務」旁的 + → 加入「Drive API」（Advanced Google Services，選 v3）
 *   4. 手動執行一次 rotateCategoryOrder()（會跳出授權畫面，同意即可；
 *      這時 DRY_RUN 還是 true，只會產生一份測試檔，不會動到正式檔）
 *   5. 收信確認測試結果、開測試檔目視核對輪替順序與格式都正常
 *   6. 把 DRY_RUN 改成 false，手動執行一次確認正式檔真的被正確覆蓋
 *   7. 左側「觸發器」→ 新增觸發器 → 選 rotateCategoryOrder →
 *      時間驅動 → 選你要的星期幾/時段 → 儲存
 */

const FILE_ID = '<member_roster_sheet_id 貼這裡，不要存進 repo>';
const DRY_RUN = true; // 先測試，確認沒問題再改 false
const NOTIFY_EMAIL = Session.getActiveUser().getEmail();
const HEADER_COL_START = 2; // B 欄
const HEADER_COL_SPAN = 5; // B:F

function rotateCategoryOrder() {
  try {
    const result = doRotate_();
    if (DRY_RUN) {
      MailApp.sendEmail(
        NOTIFY_EMAIL,
        '[BNI][測試] 分組順序輪替（DRY RUN，未動正式檔）',
        `測試檔：https://drive.google.com/file/d/${result.testFileId}/view\n` +
          `日期備註改成：${result.dateNote}\n` +
          `原順序：${result.oldOrder.join('、')}\n` +
          `新順序：${result.newOrder.join('、')}`
      );
    } else {
      MailApp.sendEmail(
        NOTIFY_EMAIL,
        '[BNI] 分組順序輪替完成',
        `現用分頁「${result.sheetName}」已完成輪替。\n` +
          `日期備註改成：${result.dateNote}\n` +
          `原順序：${result.oldOrder.join('、')}\n` +
          `新順序：${result.newOrder.join('、')}`
      );
    }
  } catch (err) {
    MailApp.sendEmail(
      NOTIFY_EMAIL,
      '[BNI] 分組順序輪替失敗（正式檔未被修改）',
      `${err.message}\n${err.stack || ''}`
    );
    throw err;
  }
}

function doRotate_() {
  const originalBlob = DriveApp.getFileById(FILE_ID).getBlob();

  // 1) 暫時轉成 Google Sheet 才能用 SpreadsheetApp 編輯
  const tempResource = {
    name: '__bni_rotate_temp__' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS,
  };
  const tempFile = Drive.Files.create(tempResource, originalBlob);

  try {
    const ss = SpreadsheetApp.openById(tempFile.id);
    const sheet = ss.getSheets()[0];

    if (!/^\d{8}$/.test(sheet.getName())) {
      throw new Error(
        `現用分頁名稱「${sheet.getName()}」不是預期的 YYYYMMDD 格式，停止執行（怕抓錯分頁）`
      );
    }
    const sheetName = sheet.getName();

    // 2) 找出所有「分類標題」列（B:F 合併、單列）
    const merges = sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .getMergedRanges();
    const headerRows = merges
      .filter(
        (r) =>
          r.getNumRows() === 1 &&
          r.getColumn() === HEADER_COL_START &&
          r.getNumColumns() === HEADER_COL_SPAN
      )
      .map((r) => r.getRow())
      .sort((a, b) => a - b);
    if (headerRows.length === 0) {
      throw new Error('找不到任何分類標題列（B:F 合併），停止執行');
    }

    const lastDataRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // 3) 切出區塊 [start, end]，最後一塊到 lastDataRow 為止
    const blocks = headerRows.map((start, i) => {
      const end = i + 1 < headerRows.length ? headerRows[i + 1] - 1 : lastDataRow;
      return { start, end };
    });
    const totalRows = blocks.reduce((sum, b) => sum + (b.end - b.start + 1), 0);
    if (totalRows !== lastDataRow - headerRows[0] + 1) {
      throw new Error('分類區塊列數加總跟資料範圍不符，可能有漏算的列，停止執行以免搬錯');
    }

    const oldOrder = blocks.map((b) =>
      sheet.getRange(b.start, HEADER_COL_START).getValue()
    );

    // 4) 讀出每個區塊的值與基本格式
    // 列高是黏在「實體列號」上的，不是跟著內容走——區塊搬動位置後，
    // 沒把列高一起搬會變成「換兩行的文字被塞進單行高度的列」，內容擠在一起。
    const snapshot = blocks.map((b) => {
      const numRows = b.end - b.start + 1;
      const range = sheet.getRange(b.start, 1, numRows, lastCol);
      const rowHeights = [];
      for (let r = b.start; r <= b.end; r++) {
        rowHeights.push(sheet.getRowHeight(r));
      }
      return {
        values: range.getValues(),
        weights: range.getFontWeights(),
        colors: range.getFontColors(),
        backgrounds: range.getBackgrounds(),
        numberFormats: range.getNumberFormats(),
        wraps: range.getWrapStrategies(),
        rowHeights,
        numRows,
      };
    });

    // 5) 輪替：第一塊搬到最後，其餘往前一位
    const rotated = snapshot.slice(1).concat([snapshot[0]]);
    const newOrder = oldOrder.slice(1).concat([oldOrder[0]]);

    // 6) 先拆掉原本的分類標題合併，才能整段改值
    headerRows.forEach((r) => {
      sheet.getRange(r, HEADER_COL_START, 1, HEADER_COL_SPAN).breakApart();
    });

    // 7) 依新順序整段寫回（值＋格式＋列高），並在新的分類標題列重新合併
    let cursor = headerRows[0];
    rotated.forEach((block) => {
      const range = sheet.getRange(cursor, 1, block.numRows, lastCol);
      range.setValues(block.values);
      range.setFontWeights(block.weights);
      range.setFontColors(block.colors);
      range.setBackgrounds(block.backgrounds);
      range.setNumberFormats(block.numberFormats);
      range.setWrapStrategies(block.wraps);
      block.rowHeights.forEach((h, i) => sheet.setRowHeight(cursor + i, h));
      sheet.getRange(cursor, HEADER_COL_START, 1, HEADER_COL_SPAN).merge();
      // 前四欄（編號／姓名／英文名字／專業別）統一置中，跟框線一樣不逐列搬，
      // 直接對這個區塊底下的會員列（不含分類標題本身那列）套用
      if (block.numRows > 1) {
        sheet
          .getRange(cursor + 1, HEADER_COL_START, block.numRows - 1, 4)
          .setHorizontalAlignment('center');
      }
      cursor += block.numRows;
    });

    // 7b) 框線是整份表格統一風格（使用者確認過），不逐格搬，
    // 輪替完直接對整個資料範圍重畫一次格線比較穩，不用啃 Apps Script 讀框線的麻煩 API
    sheet
      .getRange(headerRows[0], HEADER_COL_START, lastDataRow - headerRows[0] + 1, lastCol - HEADER_COL_START + 1)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    // 7c) F1 的日期備註改成「下一個最近的星期三」（跟簽到表 docx 的日期同一套規則，
    // 但那份是 bnimembers 才會改，這份 xlsx 只有這支排程會動）
    const dateNote = nextWednesdayDateNote_();
    sheet.getRange(1, 6).setValue(dateNote);

    // 8) 轉出成 xlsx
    // File.getAs(MimeType.MICROSOFT_EXCEL) 對 Google Sheet 不支援這個轉換
    // （會丟 "Converting from ... is not supported"），改用匯出網址 + OAuth token。
    SpreadsheetApp.flush(); // 確保前面對 sheet 的操作已經套用，匯出前先強制寫入
    const exportUrl =
      'https://docs.google.com/spreadsheets/d/' + tempFile.id + '/export?format=xlsx';
    const exportResponse = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    });
    const xlsxBlob = exportResponse.getBlob().setName(sheetName + '.xlsx');

    if (DRY_RUN) {
      const testFile = DriveApp.createFile(xlsxBlob).setName(
        'bni_rotate_dryrun_' + sheetName + '.xlsx'
      );
      return { sheetName, oldOrder, newOrder, dateNote, testFileId: testFile.getId() };
    }

    // 9) 正式模式：覆蓋回原檔內容（fileId／分享連結不變）
    Drive.Files.update({}, FILE_ID, xlsxBlob);
    return { sheetName, oldOrder, newOrder, dateNote };
  } finally {
    // 10) 清掉暫存的 Google Sheet 副本
    // 注意：Apps Script 的 Drive 進階服務裡，這個方法叫 remove，不是 delete
    // （create/update/get 都跟 REST v3 同名，只有這個特例）
    Drive.Files.remove(tempFile.id);
  }
}

// 算「下一個最近的星期三」，回傳 F1 日期備註要用的字串，例如 "2026.09.23.update"。
// 今天如果剛好是星期三，算下週三（不是今天）——跟 bnimembers skill 的規則一致。
// 明確指定 Asia/Taipei 時區，不依賴 Apps Script 專案設定的預設時區（怕跟腳本擁有者
// 的專案設定不一致，算錯成隔一天）。
function nextWednesdayDateNote_() {
  const tz = 'Asia/Taipei';
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const today = new Date(todayStr + 'T00:00:00');
  const day = today.getDay(); // 0=Sun ... 3=Wed ... 6=Sat
  let daysAhead = (3 - day + 7) % 7;
  if (daysAhead === 0) daysAhead = 7;
  const nextWed = new Date(today.getTime() + daysAhead * 24 * 60 * 60 * 1000);
  return Utilities.formatDate(nextWed, tz, 'yyyy.MM.dd') + '.update';
}
