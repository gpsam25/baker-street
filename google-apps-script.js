// ========================================
// 貝克街謎題計數器 - Google Apps Script
// ========================================
// Google Sheet 欄位說明：
//   A1 = 挑戰人次
//   A2 = 成功解謎數
//   A3 = 標準答案（預設 7）
//   A4 = 管理員密碼（預設 baker2026，請自行修改）

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const action = e.parameter.action;

    let attempt = Number(sheet.getRange('A1').getValue()) || 0;
    let correct = Number(sheet.getRange('A2').getValue()) || 0;
    const correctAnswer = String(sheet.getRange('A3').getValue() || '7');
    const adminPassword = String(sheet.getRange('A4').getValue() || 'baker2026');

    // 一般計數
    if (action === 'increment_attempt') {
      attempt++;
      sheet.getRange('A1').setValue(attempt);

    } else if (action === 'increment_correct') {
      correct++;
      sheet.getRange('A2').setValue(correct);

    // 讀取計數
    } else if (action === 'get_counts') {
      return ContentService
        .createTextOutput(JSON.stringify({ attempt, correct }))
        .setMimeType(ContentService.MimeType.JSON);

    // 讀取答案
    } else if (action === 'get_answer') {
      return ContentService
        .createTextOutput(JSON.stringify({ answer: correctAnswer }))
        .setMimeType(ContentService.MimeType.JSON);

    // 管理員操作
    } else if (action === 'admin') {
      const password = e.parameter.password;
      if (password !== adminPassword) {
        return ContentService
          .createTextOutput(JSON.stringify({ error: '密碼錯誤' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const adminAction = e.parameter.admin_action;

      if (adminAction === 'get_info') {
        return ContentService
          .createTextOutput(JSON.stringify({ attempt, correct, answer: correctAnswer }))
          .setMimeType(ContentService.MimeType.JSON);

      } else if (adminAction === 'set_answer') {
        const newAnswer = String(e.parameter.value || '7');
        sheet.getRange('A3').setValue(newAnswer);
        return ContentService
          .createTextOutput(JSON.stringify({ success: true, answer: newAnswer }))
          .setMimeType(ContentService.MimeType.JSON);

      } else if (adminAction === 'reset_attempt') {
        sheet.getRange('A1').setValue(0);
        attempt = 0;

      } else if (adminAction === 'reset_correct') {
        sheet.getRange('A2').setValue(0);
        correct = 0;

      } else if (adminAction === 'reset_all') {
        sheet.getRange('A1').setValue(0);
        sheet.getRange('A2').setValue(0);
        attempt = 0;
        correct = 0;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ attempt, correct }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
