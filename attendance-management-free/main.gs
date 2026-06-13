// © 2026 ゆゆくま / GASおじ All Rights Reserved.
/**
 * ============================================================
 *  GASAttendance 無料版 v1.0
 *  スプレッドシートベースの勤怠管理ツール（KING OF TIME/ジョブカン代替）
 * ============================================================
 * 
 * シート構成:
 * - 「社員マスタ」: 社員情報
 * - 「勤怠記録」: 出退勤ログ
 * - 「月次集計」: 月間の勤務時間集計
 * - 「設定」: 設定項目
 * - 「ログ」: 操作ログ
 * 
 * 無料版制限: 社員上限10名
 */

var SH_EMPLOYEES = '社員マスタ';
var SH_ATTENDANCE = '勤怠記録';
var SH_MONTHLY = '月次集計';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';
var MAX_EMPLOYEES_FREE = 10;

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

function escHtml_(str) { if (!str) return ''; return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function fetchWithRetry_(url, options, retries) {
  retries = retries || 3;
  for (var i = 0; i < retries; i++) {
    try { Utilities.sleep(350); var res = UrlFetchApp.fetch(url, options); if (res.getResponseCode() < 300) return res; }
    catch(e) { if (i === retries - 1) throw e; Utilities.sleep(1000 * Math.pow(2, i)); }
  }
  return null;
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var emp = ss.getSheetByName(SH_EMPLOYEES);
  if (!emp) emp = ss.insertSheet(SH_EMPLOYEES);
  emp.getRange('A1:F1').setValues([['社員ID', '氏名', '部署', '雇用形態', '所定労働時間', 'メール']]);
  emp.setFrozenRows(1);
  var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['正社員', 'パート', 'アルバイト', '契約社員', '業務委託']).setAllowInvalid(false).build();
  emp.getRange('D2:D50').setDataValidation(typeRule);

  var att = ss.getSheetByName(SH_ATTENDANCE);
  if (!att) att = ss.insertSheet(SH_ATTENDANCE);
  att.getRange('A1:H1').setValues([['日付', '社員ID', '氏名', '出勤時刻', '退勤時刻', '休憩（分）', '実労働時間', 'メモ']]);
  att.setFrozenRows(1);

  var mon = ss.getSheetByName(SH_MONTHLY);
  if (!mon) { mon = ss.insertSheet(SH_MONTHLY); mon.getRange('A1:G1').setValues([['年月', '社員ID', '氏名', '出勤日数', '総労働時間', '残業時間', '有給消化']]); mon.setFrozenRows(1); }

  var set = ss.getSheetByName(SH_SETTINGS);
  if (!set) {
    set = ss.insertSheet(SH_SETTINGS);
    set.getRange('A1:B1').setValues([['項目', '値']]);
    set.getRange('A2:B5').setValues([
      ['所定労働時間/日', '8'], ['通知メール', ''], ['残業アラート閾値（時間/月）', '45'], ['締め日', '末日']
    ]);
  }

  var lg = ss.getSheetByName(SH_LOG);
  if (!lg) { lg = ss.insertSheet(SH_LOG); lg.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }

  log_('INFO', 'セットアップ完了');
  SpreadsheetApp.getUi().alert('GASAttendance 無料版のセットアップが完了しました！');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⏰ 勤怠管理')
    .addItem('🟢 出勤打刻', 'showClockInDialog')
    .addItem('🔴 退勤打刻', 'showClockOutDialog')
    .addItem('📊 今月の集計', 'calculateMonthly')
    .addItem('⚠️ 残業アラート', 'checkOvertimeAlerts')
    .addItem('➕ 社員追加', 'showAddEmployeeDialog')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'setupSheets')
    .addToUi();
}

function showAddEmployeeDialog() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_EMPLOYEES);
  var count = sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
  if (count >= MAX_EMPLOYEES_FREE) { SpreadsheetApp.getUi().alert('無料版の社員上限（' + MAX_EMPLOYEES_FREE + '名）に達しています。'); return; }
  var html = HtmlService.createHtmlOutput(getAddEmployeeHtml_()).setWidth(400).setHeight(450).setTitle('社員追加');
  SpreadsheetApp.getUi().showSidebar(html);
}

function addEmployee(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_EMPLOYEES);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_EMPLOYEES); }
  var count = Math.max(sh.getLastRow() - 1, 0);
  if (count >= MAX_EMPLOYEES_FREE) return { success: false, message: '社員上限に達しています。' };

  var id = 'E' + String(count + 1).padStart(3, '0');
  sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([[
    id, data.name || '', data.dept || '', data.type || '正社員', Number(data.hours) || 8, data.email || ''
  ]]);
  log_('INFO', '社員追加: ' + data.name);
  return { success: true, message: '社員「' + data.name + '」を追加しました。ID: ' + id };
}

function showClockInDialog() {
  var html = HtmlService.createHtmlOutput(getClockDialog_('出勤')).setWidth(350).setHeight(300).setTitle('出勤打刻');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showClockOutDialog() {
  var html = HtmlService.createHtmlOutput(getClockDialog_('退勤')).setWidth(350).setHeight(350).setTitle('退勤打刻');
  SpreadsheetApp.getUi().showSidebar(html);
}

function clockIn(empId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSh = ss.getSheetByName(SH_EMPLOYEES);
  if (!empSh || empSh.getLastRow() < 2) return { success: false, message: '社員マスタが空です。' };

  var employees = empSh.getRange(2, 1, empSh.getLastRow() - 1, 6).getValues();
  var emp = null;
  for (var i = 0; i < employees.length; i++) { if (employees[i][0] === empId) { emp = employees[i]; break; } }
  if (!emp) return { success: false, message: '社員ID未発見。' };

  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');

  var att = ss.getSheetByName(SH_ATTENDANCE);
  if (!att) { setupSheets(); att = ss.getSheetByName(SH_ATTENDANCE); }

  // 重複チェック
  if (att.getLastRow() > 1) {
    var records = att.getRange(2, 1, att.getLastRow() - 1, 4).getValues();
    for (var r = 0; r < records.length; r++) {
      var recDate = records[r][0] instanceof Date ? Utilities.formatDate(records[r][0], 'Asia/Tokyo', 'yyyy-MM-dd') : String(records[r][0]);
      if (recDate === today && records[r][1] === empId && records[r][3]) return { success: false, message: '本日は既に出勤打刻済みです。' };
    }
  }

  att.getRange(att.getLastRow() + 1, 1, 1, 8).setValues([[today, empId, emp[1], time, '', 60, '', '']]);
  log_('INFO', '出勤: ' + emp[1] + ' ' + time);
  return { success: true, message: emp[1] + 'さん 出勤打刻完了（' + time + '）' };
}

function clockOut(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var att = ss.getSheetByName(SH_ATTENDANCE);
  if (!att || att.getLastRow() < 2) return { success: false, message: '勤怠記録がありません。' };

  var now = new Date();
  var today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var time = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');

  var records = att.getRange(2, 1, att.getLastRow() - 1, 8).getValues();
  for (var i = records.length - 1; i >= 0; i--) {
    var recDate = records[i][0] instanceof Date ? Utilities.formatDate(records[i][0], 'Asia/Tokyo', 'yyyy-MM-dd') : String(records[i][0]);
    if (recDate === today && records[i][1] === data.empId && !records[i][4]) {
      att.getRange(i + 2, 5).setValue(time);
      var breakMin = Number(data.breakMin) || 60;
      att.getRange(i + 2, 6).setValue(breakMin);

      // 実労働時間計算
      var inParts = String(records[i][3]).split(':');
      var outParts = time.split(':');
      var inMin = parseInt(inParts[0]) * 60 + parseInt(inParts[1]);
      var outMin = parseInt(outParts[0]) * 60 + parseInt(outParts[1]);
      var workMin = outMin - inMin - breakMin;
      var workHours = Math.round(workMin / 60 * 100) / 100;
      att.getRange(i + 2, 7).setValue(workHours);
      if (data.memo) att.getRange(i + 2, 8).setValue(data.memo);

      log_('INFO', '退勤: ' + records[i][2] + ' ' + time + ' 実働' + workHours + 'h');
      return { success: true, message: records[i][2] + 'さん 退勤打刻完了（' + time + '）実働: ' + workHours + '時間' };
    }
  }
  return { success: false, message: '本日の出勤記録が見つかりません。' };
}

function calculateMonthly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var att = ss.getSheetByName(SH_ATTENDANCE);
  var emp = ss.getSheetByName(SH_EMPLOYEES);
  var mon = ss.getSheetByName(SH_MONTHLY);
  if (!att || att.getLastRow() < 2) { SpreadsheetApp.getUi().alert('勤怠記録なし。'); return; }

  var now = new Date();
  var yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  var stdHours = Number(getConfig_('所定労働時間/日')) || 8;

  var records = att.getRange(2, 1, att.getLastRow() - 1, 8).getValues();
  var byEmployee = {};

  for (var i = 0; i < records.length; i++) {
    var recDate = records[i][0] instanceof Date ? Utilities.formatDate(records[i][0], 'Asia/Tokyo', 'yyyy-MM') : String(records[i][0]).substring(0, 7);
    if (recDate !== yearMonth) continue;
    var eid = records[i][1];
    if (!byEmployee[eid]) byEmployee[eid] = { name: records[i][2], days: 0, totalHours: 0 };
    byEmployee[eid].days++;
    byEmployee[eid].totalHours += Number(records[i][6]) || 0;
  }

  // 月次集計シートに書き出し
  if (!mon) { setupSheets(); mon = ss.getSheetByName(SH_MONTHLY); }
  // 既存の同月データを削除
  if (mon.getLastRow() > 1) {
    var existing = mon.getRange(2, 1, mon.getLastRow() - 1, 1).getValues();
    for (var e = existing.length - 1; e >= 0; e--) {
      if (String(existing[e][0]) === yearMonth) mon.deleteRow(e + 2);
    }
  }

  var msg = '📊 月次集計（' + yearMonth + '）\n\n';
  for (var id in byEmployee) {
    var d = byEmployee[id];
    var overtime = Math.max(d.totalHours - d.days * stdHours, 0);
    mon.getRange(mon.getLastRow() + 1, 1, 1, 7).setValues([[yearMonth, id, d.name, d.days, Math.round(d.totalHours * 100) / 100, Math.round(overtime * 100) / 100, 0]]);
    msg += d.name + ': ' + d.days + '日 / ' + Math.round(d.totalHours * 10) / 10 + 'h / 残業' + Math.round(overtime * 10) / 10 + 'h\n';
  }

  SpreadsheetApp.getUi().alert(msg);
  log_('INFO', '月次集計完了: ' + yearMonth);
}

function checkOvertimeAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mon = ss.getSheetByName(SH_MONTHLY);
  if (!mon || mon.getLastRow() < 2) { calculateMonthly(); mon = ss.getSheetByName(SH_MONTHLY); }
  if (!mon || mon.getLastRow() < 2) { SpreadsheetApp.getUi().alert('集計データなし。'); return; }

  var threshold = Number(getConfig_('残業アラート閾値（時間/月）')) || 45;
  var data = mon.getRange(2, 1, mon.getLastRow() - 1, 7).getValues();
  var alerts = [];

  for (var i = 0; i < data.length; i++) {
    var overtime = Number(data[i][5]) || 0;
    if (overtime >= threshold) alerts.push('🚨 ' + data[i][2] + ': 残業' + overtime + '時間（上限' + threshold + '時間）');
    else if (overtime >= threshold * 0.8) alerts.push('⚠️ ' + data[i][2] + ': 残業' + overtime + '時間（上限の' + Math.round(overtime / threshold * 100) + '%）');
  }

  if (alerts.length === 0) { SpreadsheetApp.getUi().alert('残業アラートはありません。'); return; }
  SpreadsheetApp.getUi().alert('⏰ 残業アラート\n\n' + alerts.join('\n'));

  var email = getConfig_('通知メール');
  if (email) { try { GmailApp.sendEmail(email, '【GASAttendance】残業アラート', alerts.join('\n')); } catch(e) {} }
}

function getAddEmployeeHtml_() {
  return '<style>body{font-family:sans-serif;padding:16px}label{display:block;margin:12px 0 4px;font-weight:bold;font-size:13px}' +
    'input,select{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'button{margin-top:16px;width:100%;padding:12px;background:#4285f4;color:white;border:none;border-radius:4px;font-size:15px;cursor:pointer}' +
    '#result{margin-top:12px;padding:8px;border-radius:4px;display:none}.ok{background:#e6f4ea;color:#137333}.ng{background:#fce8e6;color:#c5221f}</style>' +
    '<label>氏名 *</label><input id="name">' +
    '<label>部署</label><input id="dept">' +
    '<label>雇用形態</label><select id="type"><option>正社員</option><option>パート</option><option>アルバイト</option><option>契約社員</option><option>業務委託</option></select>' +
    '<label>所定労働時間/日</label><input id="hours" type="number" value="8">' +
    '<label>メール</label><input id="email" type="email">' +
    '<button onclick="submit()">追加</button><div id="result"></div>' +
    '<script>function submit(){var d={name:document.getElementById("name").value,dept:document.getElementById("dept").value,' +
    'type:document.getElementById("type").value,hours:document.getElementById("hours").value,' +
    'email:document.getElementById("email").value};if(!d.name){alert("氏名は必須");return;}' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";' +
    'el.className=r.success?"ok":"ng";el.textContent=r.message;if(r.success)document.getElementById("name").value="";}).addEmployee(d);}</script>';
}

function getClockDialog_(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_EMPLOYEES);
  var opts = '';
  if (sh && sh.getLastRow() > 1) {
    var d = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < d.length; i++) opts += '<option value="' + d[i][0] + '">' + d[i][0] + ' - ' + d[i][1] + '</option>';
  }
  var extra = type === '退勤' ? '<label>休憩時間（分）</label><input id="breakMin" type="number" value="60"><label>メモ</label><input id="memo">' : '';

  return '<style>body{font-family:sans-serif;padding:16px}label{display:block;margin:12px 0 4px;font-weight:bold;font-size:13px}' +
    'select,input{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'button{margin-top:16px;width:100%;padding:14px;background:' + (type === '出勤' ? '#34a853' : '#ea4335') + ';color:white;border:none;border-radius:4px;font-size:16px;cursor:pointer}' +
    '#result{margin-top:12px;padding:8px;border-radius:4px;display:none}.ok{background:#e6f4ea;color:#137333}.ng{background:#fce8e6;color:#c5221f}</style>' +
    '<h2>' + (type === '出勤' ? '🟢' : '🔴') + ' ' + type + '打刻</h2>' +
    '<label>社員</label><select id="empId">' + opts + '</select>' + extra +
    '<button onclick="submit()">' + type + '打刻</button><div id="result"></div>' +
    '<script>function submit(){' + (type === '出勤' ?
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";el.className=r.success?"ok":"ng";el.textContent=r.message;}).clockIn(document.getElementById("empId").value);' :
    'var d={empId:document.getElementById("empId").value,breakMin:document.getElementById("breakMin").value,memo:document.getElementById("memo").value};' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";el.className=r.success?"ok":"ng";el.textContent=r.message;}).clockOut(d);') + '}</script>';
}
