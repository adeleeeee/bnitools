/**
 * 每週上傳資料夾：週一 08:00 建立以本週三日期命名的資料夾、開放連結編輯並發通知；
 * 週二 17:30（台灣時間）自動關閉連結分享。
 *
 * 指令碼屬性（專案設定 → 指令碼屬性，不要寫進程式）：
 *   PARENT_ID       必填，上層資料夾 ID
 *   TEAMS_WEBHOOK   通知管道，與 NOTIFY_EMAIL 至少填一個
 *   NOTIFY_EMAIL    通知管道，與 TEAMS_WEBHOOK 至少填一個（帳號可能更換，由接手的人手動填）
 *
 * 部署後手動執行一次 setup() 建立排程，再用 testNotify()、testNow() 測試。
 */
const TZ = 'Asia/Taipei';

// 只需手動執行一次：建立「每週一早上 8 點」的排程
function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('openWeeklyFolder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .inTimezone(TZ)
    .create();
}

// 每週一執行：建立本週三日期的資料夾並開放上傳
function openWeeklyFolder() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1=週一 ... 7=週日
  const DAY = 86400000;

  const wed = new Date(now.getTime() + (3 - dow) * DAY);
  const name = Utilities.formatDate(wed, TZ, 'yyyyMMdd');

  const parent = DriveApp.getFolderById(requireProp_('PARENT_ID'));
  const existing = parent.getFoldersByName(name);
  const folder = existing.hasNext() ? existing.next() : parent.createFolder(name);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  // 本週二 17:30（台灣時間）收權限
  const tue = new Date(now.getTime() + (2 - dow) * DAY);
  const closeAt = new Date(Utilities.formatDate(tue, TZ, 'yyyy-MM-dd') + 'T17:30:00+08:00');
  PropertiesService.getScriptProperties().setProperty('OPEN_FOLDER_ID', folder.getId());
  ScriptApp.newTrigger('closeWeeklyFolder').timeBased().at(closeAt).create();

  notify_('本週上傳資料夾 ' + name + '：' + folder.getUrl());
}

// 時間到執行：關閉連結分享
function closeWeeklyFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('OPEN_FOLDER_ID');
  if (id) {
    DriveApp.getFolderById(id)
      .setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    props.deleteProperty('OPEN_FOLDER_ID');
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'closeWeeklyFolder')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// 測試用：馬上建立測試資料夾，5 分鐘後收權限
function testNow() {
  const parent = DriveApp.getFolderById(requireProp_('PARENT_ID'));
  const name = 'TEST_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  const folder = parent.createFolder(name);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  PropertiesService.getScriptProperties().setProperty('OPEN_FOLDER_ID', folder.getId());
  ScriptApp.newTrigger('closeWeeklyFolder').timeBased().after(5 * 60 * 1000).create();
  notify_('測試上傳資料夾 ' + name + '：' + folder.getUrl());
}

// 讀必填的指令碼屬性，沒設定就直接報錯
function requireProp_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('缺少指令碼屬性：' + key);
  return value;
}

// 所有通知都走這裡；要換通知管道，只改這個函式
function notify_(text) {
  const props = PropertiesService.getScriptProperties();
  const webhook = props.getProperty('TEAMS_WEBHOOK');
  const email = props.getProperty('NOTIFY_EMAIL');
  if (!webhook && !email) throw new Error('請設定 TEAMS_WEBHOOK 或 NOTIFY_EMAIL');
  if (webhook) postToTeams_(webhook, text);
  if (email) MailApp.sendEmail(email, '[BNI] 每週上傳資料夾', text);
}

// 送訊息到 Teams（Flow bot 聊天）
function postToTeams_(webhook, text) {
  const payload = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [{ type: 'TextBlock', text: text, wrap: true }]
      }
    }]
  };
  UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}

function testNotify() {
  notify_('通知測試訊息，現在時間 ' + Utilities.formatDate(new Date(), TZ, 'HH:mm'));
}