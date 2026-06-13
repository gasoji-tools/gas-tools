// © 2026 ゆゆくま / GASおじ All Rights Reserved.
/**
 * ============================================================
 *  GASSheetReport 無料版 v1.0
 *  スプレッドシートデータの自動レポート生成＆メール送信
 * ============================================================
 * 
 * シート構成:
 * - 「レポート設定」: レポートの定義
 * - 「送信ログ」: 送信履歴
 * - 「設定」: 設定項目
 * - 「ログ」: 操作ログ
 * 
 * 無料版制限: レポート上限3件
 */

// ========== 定数 ==========
var SH_REPORTS = 'レポート設定';
var SH_SEND_LOG = '送信ログ';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';
var MAX_REPORTS_FREE = 3;

// ========== 共通ユーティリティ ==========

function getConfig_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_SETTINGS);
  if (!sh || sh.getLastRow() < 2) return '';
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) { if (data[i][0] === key) return data[i][1]; }
  return '';
}

function log_(level, msg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SH_LOG);
    if (!sh) { sh = ss.insertSheet(SH_LOG); sh.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[new Date(), level, msg]]);
    if (sh.getLastRow() > 501) sh.deleteRows(2, sh.getLastRow() - 501);
  } catch(e) { Logger.log(level + ': ' + msg); }
}

function escHtml_(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fetchWithRetry_(url, options, retries) {
  retries = retries || 3;
  for (var i = 0; i < retries; i++) {
    try { Utilities.sleep(350); var res = UrlFetchApp.fetch(url, options); if (res.getResponseCode() < 300) return res; }
    catch(e) { if (i === retries - 1) throw e; Utilities.sleep(1000 * Math.pow(2, i)); }
  }
  return null;
}

// ========== 初期セットアップ ==========

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sh = ss.getSheetByName(SH_REPORTS);
  if (!sh) sh = ss.insertSheet(SH_REPORTS);
  sh.getRange('A1:H1').setValues([['ID', 'レポート名', '対象シート名', '送信先メール', '件名テンプレート', 'スケジュール', '有効', '最終送信']]);
  sh.setFrozenRows(1);

  var enableRule = SpreadsheetApp.newDataValidation().requireValueInList(['ON', 'OFF']).setAllowInvalid(false).build();
  sh.getRange('G2:G20').setDataValidation(enableRule);

  var schedRule = SpreadsheetApp.newDataValidation().requireValueInList(['毎日', '毎週月曜', '毎月1日']).setAllowInvalid(false).build();
  sh.getRange('F2:F20').setDataValidation(schedRule);

  var sl = ss.getSheetByName(SH_SEND_LOG);
  if (!sl) { sl = ss.insertSheet(SH_SEND_LOG); sl.getRange('A1:D1').setValues([['日時', 'レポート名', '送信先', 'ステータス']]); sl.setFrozenRows(1); }

  var set = ss.getSheetByName(SH_SETTINGS);
  if (!set) {
    set = ss.insertSheet(SH_SETTINGS);
    set.getRange('A1:B1').setValues([['項目', '値']]);
    set.getRange('A2:B4').setValues([
      ['送信者名', 'GASSheetReport'],
      ['送信時刻', '9'],
      ['フッターテキスト', 'このレポートはGASSheetReportにより自動生成されました。']
    ]);
  }

  var lg = ss.getSheetByName(SH_LOG);
  if (!lg) { lg = ss.insertSheet(SH_LOG); lg.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }

  log_('INFO', 'セットアップ完了');
  SpreadsheetApp.getUi().alert('GASSheetReport 無料版のセットアップが完了しました！');
}

// ========== カスタムメニュー ==========

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 SheetReport')
    .addItem('➕ レポート追加', 'showAddReportDialog')
    .addItem('▶️ 手動送信（全レポート）', 'sendAllReports')
    .addItem('📋 送信ログ', 'showSendLog')
    .addItem('⏰ 自動送信ON', 'setupTrigger')
    .addItem('⏹️ 自動送信OFF', 'removeTrigger')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'setupSheets')
    .addToUi();
}

// ========== レポート追加 ==========

function showAddReportDialog() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_REPORTS);
  var count = sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
  if (count >= MAX_REPORTS_FREE) {
    SpreadsheetApp.getUi().alert('無料版のレポート上限（' + MAX_REPORTS_FREE + '件）に達しています。');
    return;
  }
  var html = HtmlService.createHtmlOutput(getAddReportHtml_()).setWidth(400).setHeight(450).setTitle('レポート追加');
  SpreadsheetApp.getUi().showSidebar(html);
}

function addReport(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_REPORTS);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_REPORTS); }
  var count = Math.max(sh.getLastRow() - 1, 0);
  if (count >= MAX_REPORTS_FREE) return { success: false, message: 'レポート上限に達しています。' };

  var id = count + 1;
  sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[
    id, data.name || '', data.sheetName || '', data.email || '',
    data.subject || '【レポート】' + data.name, data.schedule || '毎日', 'ON', ''
  ]]);
  log_('INFO', 'レポート追加: ' + data.name);
  return { success: true, message: 'レポート「' + data.name + '」を追加しました。' };
}

// ========== レポート送信 ==========

function sendAllReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_REPORTS);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('レポートが設定されていません。'); return; }

  var reports = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var footer = getConfig_('フッターテキスト') || '';
  var sent = 0;

  for (var i = 0; i < reports.length; i++) {
    if (reports[i][6] !== 'ON') continue;

    var sheetName = reports[i][2];
    var email = reports[i][3];
    var subject = reports[i][4];
    if (!sheetName || !email) continue;

    var targetSh = ss.getSheetByName(sheetName);
    if (!targetSh) { log_('WARN', 'シート未発見: ' + sheetName); continue; }

    // データ集計
    var lastRow = targetSh.getLastRow();
    var lastCol = targetSh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;

    var headers = targetSh.getRange(1, 1, 1, lastCol).getValues()[0];
    var dataRows = lastRow > 1 ? targetSh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

    // HTML表を生成
    var html = '<h2>' + escHtml_(reports[i][1]) + '</h2>';
    html += '<p>シート: ' + escHtml_(sheetName) + ' / 件数: ' + dataRows.length + '行</p>';
    html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-family:sans-serif;font-size:13px">';
    html += '<tr style="background:#4285f4;color:white">';
    for (var h = 0; h < headers.length; h++) html += '<th>' + escHtml_(String(headers[h])) + '</th>';
    html += '</tr>';

    var maxRows = Math.min(dataRows.length, 50);
    for (var r = 0; r < maxRows; r++) {
      html += '<tr style="background:' + (r % 2 === 0 ? '#fff' : '#f8f9fa') + '">';
      for (var c = 0; c < headers.length; c++) {
        var val = dataRows[r][c];
        if (val instanceof Date) val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
        html += '<td>' + escHtml_(String(val || '')) + '</td>';
      }
      html += '</tr>';
    }
    html += '</table>';

    if (dataRows.length > 50) html += '<p>※50行まで表示。全' + dataRows.length + '行</p>';
    if (footer) html += '<hr><p style="color:#888;font-size:11px">' + escHtml_(footer) + '</p>';

    // 日付変数展開
    var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    subject = subject.replace(/\{\{date\}\}/g, today);

    try {
      GmailApp.sendEmail(email, subject, '', { htmlBody: html });
      sent++;
      sh.getRange(i + 2, 8).setValue(new Date());

      var logSh = ss.getSheetByName(SH_SEND_LOG);
      if (logSh) logSh.getRange(logSh.getLastRow() + 1, 1, 1, 4).setValues([[new Date(), reports[i][1], email, '成功']]);
      log_('INFO', 'レポート送信: ' + reports[i][1] + ' → ' + email);
    } catch(e) {
      var logSh2 = ss.getSheetByName(SH_SEND_LOG);
      if (logSh2) logSh2.getRange(logSh2.getLastRow() + 1, 1, 1, 4).setValues([[new Date(), reports[i][1], email, '失敗: ' + e.message]]);
      log_('ERROR', 'レポート送信失敗: ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert(sent + '件のレポートを送信しました。');
}

// ========== 送信ログ ==========

function showSendLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSh = ss.getSheetByName(SH_SEND_LOG);
  var total = logSh && logSh.getLastRow() > 1 ? logSh.getLastRow() - 1 : 0;
  SpreadsheetApp.getUi().alert('📋 送信ログ\n\n合計送信: ' + total + '件');
}

// ========== トリガー ==========

function setupTrigger() {
  removeTrigger();
  var hour = parseInt(getConfig_('送信時刻')) || 9;
  ScriptApp.newTrigger('scheduledSend').timeBased().everyDays(1).atHour(hour).create();
  SpreadsheetApp.getUi().alert('毎日' + hour + '時に自動送信を設定しました。');
  log_('INFO', 'トリガー設定: 毎日' + hour + '時');
}

function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scheduledSend') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function scheduledSend() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_REPORTS);
  if (!sh || sh.getLastRow() < 2) return;

  var reports = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var now = new Date();
  var dayOfWeek = now.getDay();
  var dayOfMonth = now.getDate();

  for (var i = 0; i < reports.length; i++) {
    if (reports[i][6] !== 'ON') continue;
    var schedule = reports[i][5];
    if (schedule === '毎日' ||
        (schedule === '毎週月曜' && dayOfWeek === 1) ||
        (schedule === '毎月1日' && dayOfMonth === 1)) {
      // 個別送信
      sendSingleReport_(i, reports[i]);
    }
  }
}

function sendSingleReport_(index, report) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = report[2];
  var email = report[3];
  var subject = report[4];
  if (!sheetName || !email) return;

  var targetSh = ss.getSheetByName(sheetName);
  if (!targetSh) return;

  var lastRow = targetSh.getLastRow();
  var lastCol = targetSh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  var headers = targetSh.getRange(1, 1, 1, lastCol).getValues()[0];
  var dataRows = lastRow > 1 ? targetSh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var footer = getConfig_('フッターテキスト') || '';

  var html = '<h2>' + escHtml_(report[1]) + '</h2><p>件数: ' + dataRows.length + '行</p>';
  html += '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;font-size:13px">';
  html += '<tr style="background:#4285f4;color:white">';
  for (var h = 0; h < headers.length; h++) html += '<th>' + escHtml_(String(headers[h])) + '</th>';
  html += '</tr>';
  var max = Math.min(dataRows.length, 50);
  for (var r = 0; r < max; r++) {
    html += '<tr style="background:' + (r % 2 === 0 ? '#fff' : '#f8f9fa') + '">';
    for (var c = 0; c < headers.length; c++) {
      var v = dataRows[r][c];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      html += '<td>' + escHtml_(String(v || '')) + '</td>';
    }
    html += '</tr>';
  }
  html += '</table>';
  if (footer) html += '<hr><p style="color:#888;font-size:11px">' + escHtml_(footer) + '</p>';

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  subject = subject.replace(/\{\{date\}\}/g, today);

  try {
    GmailApp.sendEmail(email, subject, '', { htmlBody: html });
    var sh = ss.getSheetByName(SH_REPORTS);
    if (sh) sh.getRange(index + 2, 8).setValue(new Date());
    log_('INFO', '自動レポート送信: ' + report[1]);
  } catch(e) { log_('ERROR', 'レポート送信失敗: ' + e.message); }
}

// ========== フォームHTML ==========

function getAddReportHtml_() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return [SH_REPORTS, SH_SEND_LOG, SH_SETTINGS, SH_LOG].indexOf(n) < 0; });
  var opts = sheets.map(function(s) { return '<option value="' + s + '">' + s + '</option>'; }).join('');

  return '<style>body{font-family:sans-serif;padding:16px}label{display:block;margin:12px 0 4px;font-weight:bold;font-size:13px}' +
    'input,select{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'button{margin-top:16px;width:100%;padding:12px;background:#4285f4;color:white;border:none;border-radius:4px;font-size:15px;cursor:pointer}' +
    'button:hover{background:#3367d6}#result{margin-top:12px;padding:8px;border-radius:4px;display:none}' +
    '.ok{background:#e6f4ea;color:#137333}.ng{background:#fce8e6;color:#c5221f}</style>' +
    '<label>レポート名 *</label><input id="name" placeholder="日次売上レポート">' +
    '<label>対象シート *</label><select id="sheetName">' + opts + '</select>' +
    '<label>送信先メール *</label><input id="email" type="email" placeholder="report@example.com">' +
    '<label>件名</label><input id="subject" placeholder="【レポート】{{date}} 売上">' +
    '<label>スケジュール</label><select id="schedule"><option value="毎日">毎日</option><option value="毎週月曜">毎週月曜</option><option value="毎月1日">毎月1日</option></select>' +
    '<button onclick="submit()">追加</button><div id="result"></div>' +
    '<script>function submit(){var d={name:document.getElementById("name").value,' +
    'sheetName:document.getElementById("sheetName").value,email:document.getElementById("email").value,' +
    'subject:document.getElementById("subject").value,schedule:document.getElementById("schedule").value};' +
    'if(!d.name||!d.email){alert("レポート名と送信先は必須");return;}' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";' +
    'el.className=r.success?"ok":"ng";el.textContent=r.message;if(r.success)document.getElementById("name").value="";}).addReport(d);}</script>';
}
