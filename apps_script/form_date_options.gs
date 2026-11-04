/**
 * 小房間預約表單的日期選項維護（綁定在表單上的 script）。
 *   updateFormDates：刪除第一個日期，並在最後補下一週
 *   removeSubmittedChoice：有人預約後移除該日期，並在最後補下一週
 *
 * ⚠️ 程式找的題目標題是 "appointment"，目前表單的題目是「預約日期」，兩者不一致。
 *    標題不符時程式只記 log 就結束，修改前請先確認實際行為。
 */
/**
 * 這個函式會自動更新 Google 表單中的日期選項。
 * 它會刪除第一個選項，並新增一個新的選項，該選項為最後一個日期的下一個星期三。
 */
function updateFormDates() {
  // 根據你的表單，將問題標題設定為 "appointment"
  const questionTitle = "appointment"; 
  
  try {
    const form = FormApp.getActiveForm();
    const items = form.getItems();
    let dateItem;

    // 找到標題為 "appointment" 的問題
    for (const item of items) {
      if (item.getTitle() === questionTitle) {
        dateItem = item;
        break;
      }
    }

    if (!dateItem) {
      Logger.log("找不到標題為 '" + questionTitle + "' 的問題。請檢查問題標題是否正確。");
      return;
    }

    const itemType = dateItem.getType();
    if (itemType !== FormApp.ItemType.MULTIPLE_CHOICE && itemType !== FormApp.ItemType.CHECKBOX) {
      Logger.log("此問題類型不是多選題或核取方塊。此指令碼僅支援這兩種問題類型。");
      return;
    }

    const choiceItem = dateItem.asMultipleChoiceItem();
    const choices = choiceItem.getChoices().map(choice => choice.getValue());
    
    if (choices.length < 1) {
      Logger.log("選項列表為空，無法執行更新。");
      return;
    }

    // 刪除第一個選項
    const updatedChoices = choices.slice(1);
    
    // 找出目前最後一個日期的月份和日期
    const lastDateString = choices[choices.length - 1];
    const [month, day] = lastDateString.split('/').map(Number);
    
    // 建立一個代表最後一個日期的 JavaScript Date 物件（假設年份為當年）
    const now = new Date();
    let lastDate = new Date(now.getFullYear(), month - 1, day);
    
    // 新增 7 天，以得到下一個星期的日期
    const nextDate = new Date(lastDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    
    // 格式化新日期為 MM/DD
    const newMonth = nextDate.getMonth() + 1;
    const newDay = nextDate.getDate();
    const newDateString = `${newMonth}/${newDay}`;

    // 將新選項添加到列表中
    updatedChoices.push(newDateString);

    // 將新的選項列表設定回表單
    const newChoices = updatedChoices.map(value => choiceItem.createChoice(value));
    choiceItem.setChoices(newChoices);
    
    Logger.log("表單已成功更新。刪除第一個選項並新增了 '" + newDateString + "'。");

  } catch (e) {
    Logger.log("執行時發生錯誤：" + e.toString());
  }
}

/**
 * 當表單被提交時觸發，移除被選擇的日期選項，並新增一個新的日期以維持選項數量。
 * @param {Object} e - 表單提交事件物件。
 */
function removeSubmittedChoice(e) {
  // 請將 "appointment" 替換成你的問題標題
  const questionTitle = "appointment";
  
  try {
    const form = FormApp.getActiveForm();
    const items = form.getItems();
    let dateItem;

    for (const item of items) {
      if (item.getTitle() === questionTitle) {
        dateItem = item;
        break;
      }
    }

    if (!dateItem) {
      Logger.log("找不到問題標題：'" + questionTitle + "'");
      return;
    }

    const itemType = dateItem.getType();
    if (itemType !== FormApp.ItemType.MULTIPLE_CHOICE) {
      Logger.log("此問題類型不支援。");
      return;
    }

    const choiceItem = dateItem.asMultipleChoiceItem();
    let choices = choiceItem.getChoices().map(choice => choice.getValue());
    
    // 從提交的回答中，找到使用者選擇的日期
    const submittedResponses = e.response.getItemResponses();
    let chosenDate;

    for (const response of submittedResponses) {
      if (response.getItem().getTitle() === questionTitle) {
        chosenDate = response.getResponse();
        break;
      }
    }

    if (!chosenDate) {
      Logger.log("找不到已選擇的日期選項。");
      return;
    }

    // 1. 移除被選擇的日期
    const newChoices = choices.filter(choice => choice !== chosenDate);

    // 2. 找出目前列表的最後一個日期，並計算出下一個日期的字串
    const lastDateString = choices[choices.length - 1];
    const [month, day] = lastDateString.split('/').map(Number);
    
    // 建立一個代表最後一個日期的 JavaScript Date 物件（假設年份為當年）
    const now = new Date();
    let lastDate = new Date(now.getFullYear(), month - 1, day);
    
    // 新增 7 天，以得到下一個日期的字串
    const nextDate = new Date(lastDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    const newMonth = nextDate.getMonth() + 1;
    const newDay = nextDate.getDate();
    const newDateString = `${newMonth}/${newDay}`;

    // 3. 將新的日期加入列表，以維持選項數量
    newChoices.push(newDateString);
    
    // 4. 更新表單的選項
    const newFormChoices = newChoices.map(value => choiceItem.createChoice(value));
    choiceItem.setChoices(newFormChoices);
    
    Logger.log("已成功移除被選擇的日期：" + chosenDate + "，並新增新日期：" + newDateString);

  } catch (err) {
    Logger.log("執行時發生錯誤：" + err.toString());
  }
}