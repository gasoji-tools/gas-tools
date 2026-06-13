// © 2026 ゆゆくま / GASおじ All Rights Reserved.
/**
 * ============================================================
 *  GASHelpDesk 無料版 v1.0
 *  スプレッドシートベースのお問い合わせ管理（Zendesk/Freshdesk代替）
 * ============================================================
 * 
 * シート構成:
 * - 「チケット」: お問い合わせチケット
 * - 「カテゴリ」: カテゴリ管理
 * - 「設定」: 設定項目
 * - 「ログ」: 操作ログ
 * 
 * 無料版制限: チケット100件/月、カテゴリ5件
 */

var SH_TICKETS = 'チケット';
var SH_CATEGORIES = 'カテゴリ';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';
var MAX_TICKETS_FREE = 100;
var MAX_CATEGORIES_FREE = 5;

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

  var tk = ss.getSheetByName(SH_TICKETS);
  if (!tk) tk = ss.insertSheet(SH_TICKETS);
  tk.getRange('A1:J1').setValues([['チケットID', '件名', 'お客様名', 'メール', 'カテゴリ', '優先度', 'ステータス', '担当者', '作成日', '内容']]);
  tk.setFrozenRows(1);
  var prRule = SpreadsheetApp.newDataValidation().requireValueInList(['低', '中', '高', '緊急']).setAllowInvalid(false).build();
  tk.getRange('F2:F500').setDataValidation(prRule);
  var stRule = SpreadsheetApp.newDataValidation().requireValueInList(['新規', '対応中', '保留', '解決済み', 'クローズ']).setAllowInvalid(false).build();
  tk.getRange('G2:G500').setDataValidation(stRule);

  var cat = ss.getSheetByName(SH_CATEGORIES);
  if (!cat) {
    cat = ss.insertSheet(SH_CATEGORIES);
    cat.getRange('A1:B1').setValues([['カテゴリ名', '説明']]);
    cat.getRange('A2:B4').setValues([['一般', '一般的なお問い合わせ'], ['技術', '技術的な質問'], ['請求', '請求に関する問い合わせ']]);
  }

  var set = ss.getSheetByName(SH_SETTINGS);
  if (!set) {
    set = ss.insertSheet(SH_SETTINGS);
    set.getRange('A1:B1').setValues([['項目', '値']]);
    set.getRange('A2:B4').setValues([['通知メール', ''], ['自動返信', 'ON'], ['自動返信文', 'お問い合わせありがとうございます。担当者が確認次第、ご連絡いたします。']]);
  }

  var lg = ss.getSheetByName(SH_LOG);
  if (!lg) { lg = ss.insertSheet(SH_LOG); lg.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }

  log_('INFO', 'セットアップ完了');
  SpreadsheetApp.getUi().alert('GASHelpDesk 無料版のセットアップが完了しました！');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎫 ヘルプデスク')
    .addItem('➕ チケット作成', 'showCreateTicketDialog')
    .addItem('📊 ダッシュボード', 'showDashboard')
    .addItem('📋 未対応チケット', 'showOpenTickets')
    .addItem('🔍 チケット検索', 'showSearchDialog')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'setupSheets')
    .addToUi();
}

function showCreateTicketDialog() {
  var html = HtmlService.createHtmlOutput(getCreateTicketHtml_()).setWidth(420).setHeight(520).setTitle('チケット作成');
  SpreadsheetApp.getUi().showSidebar(html);
}

function createTicket(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TICKETS);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_TICKETS); }

  var now = new Date();
  var id = 'TK-' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmss');

  sh.getRange(sh.getLastRow() + 1, 1, 1, 10).setValues([[
    id, data.subject || '', data.name || '', data.email || '', data.category || '一般',
    data.priority || '中', '新規', data.assignee || '', now, data.content || ''
  ]]);

  // 自動返信
  if (getConfig_('自動返信') === 'ON' && data.email) {
    var replyText = getConfig_('自動返信文') || '';
    try { GmailApp.sendEmail(data.email, '【受付完了】' + data.subject + ' (' + id + ')', data.name + '様\n\n' + replyText + '\n\nチケットID: ' + id); } catch(e) {}
  }

  log_('INFO', 'チケット作成: ' + id + ' ' + data.subject);
  return { success: true, message: 'チケット作成完了！ID: ' + id };
}

function showDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TICKETS);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('チケットなし。'); return; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  var total = data.length, byStatus = {}, byPriority = {}, byCategory = {};
  for (var i = 0; i < data.length; i++) {
    byStatus[data[i][6] || '不明'] = (byStatus[data[i][6] || '不明'] || 0) + 1;
    byPriority[data[i][5] || '不明'] = (byPriority[data[i][5] || '不明'] || 0) + 1;
    byCategory[data[i][4] || '不明'] = (byCategory[data[i][4] || '不明'] || 0) + 1;
  }

  var msg = '🎫 ヘルプデスク ダッシュボード\n\n合計: ' + total + '件\n\n【ステータス別】\n';
  for (var s in byStatus) msg += '  ' + s + ': ' + byStatus[s] + '件\n';
  msg += '\n【優先度別】\n';
  for (var p in byPriority) msg += '  ' + p + ': ' + byPriority[p] + '件\n';
  msg += '\n【カテゴリ別】\n';
  for (var c in byCategory) msg += '  ' + c + ': ' + byCategory[c] + '件\n';
  SpreadsheetApp.getUi().alert(msg);
}

function showOpenTickets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TICKETS);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('チケットなし。'); return; }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  var open = data.filter(function(r) { return r[6] === '新規' || r[6] === '対応中'; });
  if (open.length === 0) { SpreadsheetApp.getUi().alert('未対応チケットなし！'); return; }

  var msg = '📋 未対応チケット（' + open.length + '件）\n\n';
  for (var i = 0; i < Math.min(open.length, 20); i++) {
    msg += '[' + open[i][5] + '] ' + open[i][0] + ' ' + open[i][1] + ' — ' + open[i][2] + ' [' + open[i][6] + ']\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}

function showSearchDialog() {
  var query = SpreadsheetApp.getUi().prompt('🔍 チケット検索', 'キーワードを入力:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if (query.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;

  var keyword = query.getResponseText().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_TICKETS);
  if (!sh || sh.getLastRow() < 2) return;

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  var results = data.filter(function(r) {
    return String(r[1]).toLowerCase().indexOf(keyword) >= 0 || String(r[2]).toLowerCase().indexOf(keyword) >= 0 || String(r[9]).toLowerCase().indexOf(keyword) >= 0;
  });

  if (results.length === 0) { SpreadsheetApp.getUi().alert('該当なし。'); return; }
  var msg = '🔍 検索結果（' + results.length + '件）\n\n';
  for (var i = 0; i < Math.min(results.length, 10); i++) {
    msg += results[i][0] + ' ' + results[i][1] + ' — ' + results[i][2] + ' [' + results[i][6] + ']\n';
  }
  SpreadsheetApp.getUi().alert(msg);
}

function getCreateTicketHtml_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catSh = ss.getSheetByName(SH_CATEGORIES);
  var catOpts = '<option>一般</option>';
  if (catSh && catSh.getLastRow() > 1) {
    catOpts = catSh.getRange(2, 1, catSh.getLastRow() - 1, 1).getValues().map(function(r) { return '<option>' + r[0] + '</option>'; }).join('');
  }

  return '<style>body{font-family:sans-serif;padding:16px}label{display:block;margin:10px 0 4px;font-weight:bold;font-size:13px}' +
    'input,select,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'textarea{height:80px}button{margin-top:16px;width:100%;padding:12px;background:#4285f4;color:white;border:none;border-radius:4px;font-size:15px;cursor:pointer}' +
    '#r{margin-top:12px;display:none;padding:8px;border-radius:4px}.ok{background:#e6f4ea}.ng{background:#fce8e6}</style>' +
    '<label>件名 *</label><input id="subject"><label>お客様名 *</label><input id="name"><label>メール</label><input id="email" type="email">' +
    '<label>カテゴリ</label><select id="category">' + catOpts + '</select>' +
    '<label>優先度</label><select id="priority"><option>低</option><option selected>中</option><option>高</option><option>緊急</option></select>' +
    '<label>担当者</label><input id="assignee"><label>内容 *</label><textarea id="content"></textarea>' +
    '<button onclick="go()">チケット作成</button><div id="r"></div>' +
    '<script>function go(){var d={subject:document.getElementById("subject").value,name:document.getElementById("name").value,' +
    'email:document.getElementById("email").value,category:document.getElementById("category").value,' +
    'priority:document.getElementById("priority").value,assignee:document.getElementById("assignee").value,' +
    'content:document.getElementById("content").value};if(!d.subject||!d.name||!d.content){alert("必須項目");return;}' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("r");el.style.display="block";' +
    'el.className=r.success?"ok":"ng";el.textContent=r.message;}).createTicket(d);}</script>';
}
