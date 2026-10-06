/**
 * UNITECH 官網表單 → Google 試算表
 * ---------------------------------------------------------------
 * 部署步驟（約 5 分鐘，只需做一次）：
 * 1. 新增一個 Google 試算表（命名例如「UNITECH 官網名單」）
 * 2. 試算表上方選單：擴充功能 → Apps Script
 * 3. 把本檔案內容全部貼進編輯器，覆蓋原本的 myFunction
 * 4. 修改下方兩行設定：SHEET_NAME（工作表名稱）、NOTIFY_EMAIL（收件通知信箱）
 * 5. 按「部署」→「新增部署作業」→ 類型選「網路應用程式」
 *    - 執行身分：我
 *    - 存取權：所有人（重要，否則官網送不進來）
 * 6. 首次會要求授權 → 允許
 * 7. 複製產生的「網路應用程式 URL」，貼回 index-v2.html 的 LEAD_ENDPOINT
 * 8. 日後若修改程式碼，要「部署 → 管理部署 → 建立新版本」，URL 才會更新
 * ---------------------------------------------------------------
 * 注意：每次修改程式後都要重新建立版本，否則線上仍是舊程式。
 */

var SHEET_NAME = '名單';              // 試算表裡的工作表名稱，不存在會自動建立
var NOTIFY_EMAIL = '';                // 留白＝不寄通知信；填入例如 'you@company.com'

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);               // 避免同時多筆寫入打架
  try {
    var d = {};
    try { d = JSON.parse(e.postData.contents); } catch (err) { d = e.parameter; }

    // 蜜罐欄位：有填代表是機器人，靜默略過
    if (d.hp) return json({ ok: true, spam: true });

    if (!d.tel && !d.name) return json({ ok: false, error: 'empty' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['時間', '稱呼', '公司名稱', '聯絡電話', '聯絡方式', '來源表單', 'User Agent']);
      sh.getRange('A1:G1').setFontWeight('bold').setBackground('#f1f3f4');
      sh.setFrozenRows(1);
    }

    var tz = Session.getScriptTimeZone();
    sh.appendRow([
      Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss'),
      d.name || '',
      d.co || '',
      d.tel || '',
      d.chan || '',
      d.src || '',
      d.ua || ''
    ]);
    sh.autoResizeColumns(1, 7);

    if (NOTIFY_EMAIL) {
      try {
        MailApp.sendEmail({
          to: NOTIFY_EMAIL,
          subject: '🔔 官網新名單：' + (d.co || d.name || '—') + '（' + (d.chan || '—') + '）',
          body: [
            '稱呼：' + (d.name || ''),
            '公司：' + (d.co || ''),
            '電話：' + (d.tel || ''),
            '希望聯絡方式：' + (d.chan || ''),
            '來源表單：' + (d.src || ''),
            '時間：' + Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss'),
            '',
            '試算表：' + ss.getUrl()
          ].join('\n')
        });
      } catch (mailErr) { /* 寄信失敗不影響寫入 */ }
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 瀏覽器用 GET 開啟時，顯示簡單狀態，方便確認部署成功
function doGet() {
  return ContentService
    .createTextOutput('UNITECH lead endpoint is running. 請用 POST 送出表單。')
    .setMimeType(ContentService.MimeType.TEXT);
}

function json(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
