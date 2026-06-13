/**
 * ============================================================
 *  GASCRM 無料版 v1.0
 *  スプレッドシートベースの顧客管理ツール（Salesforce代替）
 * ============================================================
 * 
 * シート構成:
 * - 「顧客一覧」: 顧客情報管理
 * - 「商談履歴」: 商談・アプローチ記録
 * - 「設定」: 設定項目
 * - 「ログ」: 操作ログ
 * 
 * 無料版制限: 顧客上限100件、商談上限200件
 */

// ========== 定数 ==========
var SH_CUSTOMERS = '顧客一覧';
var SH_DEALS = '商談履歴';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';

var MAX_CUSTOMERS_FREE = 100;
var MAX_DEALS_FREE = 200;

var DEAL_LEAD = 'リード';
var DEAL_CONTACT = 'コンタクト済';
var DEAL_PROPOSAL = '提案中';
var DEAL_NEGOTIATION = '交渉中';
var DEAL_WON = '受注';
var DEAL_LOST = '失注';

var RANK_A = 'A（重要）';
var RANK_B = 'B（標準）';
var RANK_C = 'C（低）';

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
    try {
      Utilities.sleep(350);
      var res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() < 300) return res;
    } catch(e) {
      if (i === retries - 1) throw e;
      Utilities.sleep(1000 * Math.pow(2, i));
    }
  }
  return null;
}

// ========== カスタムメニュー ==========

function onOpen() {
  SpreadsheetApp.getUi().createMenu('👥 CRM')
    .addItem('➕ 顧客追加', 'showAddCustomerDialog')
    .addItem('📝 商談追加', 'showAddDealDialog')
    .addItem('📊 ダッシュボード', 'showDashboard')
    .addItem('⏰ フォローアップチェック', 'checkFollowUps')
    .addItem('📈 売上レポート', 'showSalesReport')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'setupSheets')
    .addToUi();
}

// ========== 顧客追加 ==========

function showAddCustomerDialog() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CUSTOMERS);
  var count = sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
  if (count >= MAX_CUSTOMERS_FREE) {
    SpreadsheetApp.getUi().alert('無料版の顧客上限（' + MAX_CUSTOMERS_FREE + '件）に達しています。\nPro版にアップグレードしてください。');
    return;
  }
  var html = HtmlService.createHtmlOutput(getAddCustomerHtml_())
    .setWidth(400).setHeight(550).setTitle('顧客追加');
  SpreadsheetApp.getUi().showSidebar(html);
}

function addCustomer(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CUSTOMERS);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_CUSTOMERS); }

  var count = Math.max(sh.getLastRow() - 1, 0);
  if (count >= MAX_CUSTOMERS_FREE) {
    return { success: false, message: '顧客上限（' + MAX_CUSTOMERS_FREE + '件）に達しています。' };
  }

  var id = count + 1;
  var now = new Date();

  sh.getRange(sh.getLastRow() + 1, 1, 1, 12).setValues([[
    id, data.company || '', data.contact || '', data.email || '', data.phone || '',
    data.address || '', data.rank || RANK_B, DEAL_LEAD, '', '', now, data.memo || ''
  ]]);

  log_('INFO', '顧客追加: ' + data.company);
  return { success: true, message: '顧客「' + data.company + '」を追加しました。' };
}

// ========== 商談追加 ==========

function showAddDealDialog() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_DEALS);
  var count = sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
  if (count >= MAX_DEALS_FREE) {
    SpreadsheetApp.getUi().alert('無料版の商談上限（' + MAX_DEALS_FREE + '件）に達しています。');
    return;
  }
  var html = HtmlService.createHtmlOutput(getAddDealHtml_())
    .setWidth(400).setHeight(450).setTitle('商談追加');
  SpreadsheetApp.getUi().showSidebar(html);
}

function addDeal(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_DEALS);
  if (!sh) { setupSheets(); sh = ss.getSheetByName(SH_DEALS); }

  var count = Math.max(sh.getLastRow() - 1, 0);
  if (count >= MAX_DEALS_FREE) {
    return { success: false, message: '商談上限に達しています。' };
  }

  var id = count + 1;
  sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[
    id, data.customerId || '', data.company || '', data.content || '',
    data.amount || 0, data.status || DEAL_LEAD, new Date(), data.memo || ''
  ]]);

  // 顧客の最終コンタクト日を更新
  if (data.customerId) {
    var custSh = ss.getSheetByName(SH_CUSTOMERS);
    if (custSh && custSh.getLastRow() > 1) {
      var custData = custSh.getRange(2, 1, custSh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < custData.length; i++) {
        if (String(custData[i][0]) === String(data.customerId)) {
          custSh.getRange(i + 2, 9).setValue(new Date());
          if (data.status) custSh.getRange(i + 2, 8).setValue(data.status);
          break;
        }
      }
    }
  }

  log_('INFO', '商談追加: ' + data.company + ' - ' + data.content);
  return { success: true, message: '商談を追加しました。' };
}

// ========== ダッシュボード ==========

function showDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var custSh = ss.getSheetByName(SH_CUSTOMERS);
  var dealSh = ss.getSheetByName(SH_DEALS);

  var totalCust = custSh && custSh.getLastRow() > 1 ? custSh.getLastRow() - 1 : 0;
  var totalDeals = dealSh && dealSh.getLastRow() > 1 ? dealSh.getLastRow() - 1 : 0;

  var byStatus = {};
  var byRank = {};
  if (custSh && custSh.getLastRow() > 1) {
    var custData = custSh.getRange(2, 1, custSh.getLastRow() - 1, 12).getValues();
    for (var i = 0; i < custData.length; i++) {
      var st = custData[i][7] || 'リード';
      byStatus[st] = (byStatus[st] || 0) + 1;
      var rk = custData[i][6] || 'B（標準）';
      byRank[rk] = (byRank[rk] || 0) + 1;
    }
  }

  var totalAmount = 0;
  var wonAmount = 0;
  if (dealSh && dealSh.getLastRow() > 1) {
    var dealData = dealSh.getRange(2, 1, dealSh.getLastRow() - 1, 8).getValues();
    for (var j = 0; j < dealData.length; j++) {
      var amt = Number(dealData[j][4]) || 0;
      totalAmount += amt;
      if (dealData[j][5] === DEAL_WON) wonAmount += amt;
    }
  }

  var msg = '📊 CRM ダッシュボード\n\n' +
    '顧客数: ' + totalCust + ' / ' + MAX_CUSTOMERS_FREE + '件\n' +
    '商談数: ' + totalDeals + ' / ' + MAX_DEALS_FREE + '件\n\n' +
    '【ステータス別】\n';
  for (var s in byStatus) msg += '  ' + s + ': ' + byStatus[s] + '件\n';
  msg += '\n【ランク別】\n';
  for (var r in byRank) msg += '  ' + r + ': ' + byRank[r] + '件\n';
  msg += '\n💰 商談総額: ¥' + totalAmount.toLocaleString() +
    '\n🎉 受注金額: ¥' + wonAmount.toLocaleString() +
    '\n📈 受注率: ' + (totalDeals > 0 ? Math.round(wonAmount / totalAmount * 100) : 0) + '%';

  SpreadsheetApp.getUi().alert(msg);
  log_('INFO', 'ダッシュボード表示');
}

// ========== フォローアップチェック ==========

function checkFollowUps() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CUSTOMERS);
  if (!sh || sh.getLastRow() < 2) return;

  var followDays = parseInt(getConfig_('フォローアップ日数')) || 7;
  var email = getConfig_('通知メール');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  var now = new Date();
  var alerts = [];

  for (var i = 0; i < data.length; i++) {
    var status = data[i][7];
    if (status === DEAL_WON || status === DEAL_LOST) continue;

    var lastContact = data[i][8];
    if (!lastContact) {
      alerts.push('📌 未コンタクト: ' + data[i][1] + '（' + data[i][2] + '）');
      continue;
    }

    var diff = Math.ceil((now - new Date(lastContact)) / (1000 * 60 * 60 * 24));
    if (diff >= followDays) {
      alerts.push('⏰ ' + diff + '日未連絡: ' + data[i][1] + '（' + data[i][2] + '）');
    }
  }

  if (alerts.length === 0) {
    SpreadsheetApp.getUi().alert('フォローアップが必要な顧客はいません。');
    return;
  }

  var msg = '👥 フォローアップアラート\n\n' + alerts.join('\n');
  SpreadsheetApp.getUi().alert(msg);

  if (email) {
    try { GmailApp.sendEmail(email, '【GASCRM】フォローアップアラート', msg); log_('INFO', 'フォローアップメール送信'); }
    catch(e) { log_('ERROR', 'メール送信失敗: ' + e.message); }
  }

  log_('INFO', 'フォローアップチェック: ' + alerts.length + '件');
}

// ========== 売上レポート ==========

function showSalesReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_DEALS);
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('商談データがありません。');
    return;
  }

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var monthly = {};
  var byCompany = {};

  for (var i = 0; i < data.length; i++) {
    if (data[i][5] !== DEAL_WON) continue;
    var date = new Date(data[i][6]);
    var month = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
    var amt = Number(data[i][4]) || 0;

    monthly[month] = (monthly[month] || 0) + amt;

    var company = data[i][2] || '不明';
    byCompany[company] = (byCompany[company] || 0) + amt;
  }

  var msg = '📈 売上レポート（受注分のみ）\n\n【月別】\n';
  var months = Object.keys(monthly).sort();
  for (var m = 0; m < months.length; m++) {
    msg += '  ' + months[m] + ': ¥' + monthly[months[m]].toLocaleString() + '\n';
  }

  msg += '\n【会社別】\n';
  var companies = Object.keys(byCompany).sort(function(a, b) { return byCompany[b] - byCompany[a]; });
  for (var c = 0; c < Math.min(companies.length, 10); c++) {
    msg += '  ' + companies[c] + ': ¥' + byCompany[companies[c]].toLocaleString() + '\n';
  }

  SpreadsheetApp.getUi().alert(msg);
  log_('INFO', '売上レポート表示');
}

// ========== フォームHTML ==========

function getAddCustomerHtml_() {
  return '<style>' +
    'body{font-family:sans-serif;padding:16px}' +
    'label{display:block;margin:12px 0 4px;font-weight:bold;font-size:13px}' +
    'input,select,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'textarea{height:60px;resize:vertical}' +
    'button{margin-top:16px;width:100%;padding:12px;background:#4285f4;color:white;border:none;border-radius:4px;font-size:15px;cursor:pointer}' +
    'button:hover{background:#3367d6}' +
    '#result{margin-top:12px;padding:8px;border-radius:4px;display:none}' +
    '.ok{background:#e6f4ea;color:#137333}.ng{background:#fce8e6;color:#c5221f}' +
    '</style>' +
    '<label>会社名 *</label><input id="company" placeholder="例: 株式会社ABC">' +
    '<label>担当者名</label><input id="contact" placeholder="例: 山田太郎">' +
    '<label>メール</label><input id="email" type="email" placeholder="例: info@example.com">' +
    '<label>電話</label><input id="phone" placeholder="例: 03-1234-5678">' +
    '<label>住所</label><input id="address" placeholder="例: 東京都渋谷区...">' +
    '<label>ランク</label><select id="rank"><option value="A（重要）">A（重要）</option><option value="B（標準）" selected>B（標準）</option><option value="C（低）">C（低）</option></select>' +
    '<label>メモ</label><textarea id="memo" placeholder="補足情報"></textarea>' +
    '<button onclick="submit()">追加</button><div id="result"></div>' +
    '<script>function submit(){var d={company:document.getElementById("company").value,' +
    'contact:document.getElementById("contact").value,email:document.getElementById("email").value,' +
    'phone:document.getElementById("phone").value,address:document.getElementById("address").value,' +
    'rank:document.getElementById("rank").value,memo:document.getElementById("memo").value};' +
    'if(!d.company){alert("会社名を入力してください");return;}' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";' +
    'el.className=r.success?"ok":"ng";el.textContent=r.message;' +
    'if(r.success)document.getElementById("company").value="";}).addCustomer(d);}</script>';
}

function getAddDealHtml_() {
  // 顧客リスト取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var custSh = ss.getSheetByName(SH_CUSTOMERS);
  var custOpts = '<option value="">-- 選択 --</option>';
  if (custSh && custSh.getLastRow() > 1) {
    var custData = custSh.getRange(2, 1, custSh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < custData.length; i++) {
      custOpts += '<option value="' + custData[i][0] + '">' + custData[i][0] + '. ' + custData[i][1] + '</option>';
    }
  }

  return '<style>' +
    'body{font-family:sans-serif;padding:16px}' +
    'label{display:block;margin:12px 0 4px;font-weight:bold;font-size:13px}' +
    'input,select,textarea{width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;font-size:14px}' +
    'textarea{height:60px;resize:vertical}' +
    'button{margin-top:16px;width:100%;padding:12px;background:#34a853;color:white;border:none;border-radius:4px;font-size:15px;cursor:pointer}' +
    'button:hover{background:#2d8e47}' +
    '#result{margin-top:12px;padding:8px;border-radius:4px;display:none}' +
    '.ok{background:#e6f4ea;color:#137333}.ng{background:#fce8e6;color:#c5221f}' +
    '</style>' +
    '<label>顧客</label><select id="customerId">' + custOpts + '</select>' +
    '<label>会社名</label><input id="company" placeholder="新規顧客の場合">' +
    '<label>商談内容 *</label><input id="content" placeholder="例: 初回訪問">' +
    '<label>金額</label><input id="amount" type="number" placeholder="例: 500000">' +
    '<label>ステータス</label><select id="status">' +
    '<option value="リード">リード</option><option value="コンタクト済">コンタクト済</option>' +
    '<option value="提案中">提案中</option><option value="交渉中">交渉中</option>' +
    '<option value="受注">受注</option><option value="失注">失注</option></select>' +
    '<label>メモ</label><textarea id="memo"></textarea>' +
    '<button onclick="submit()">追加</button><div id="result"></div>' +
    '<script>function submit(){var d={customerId:document.getElementById("customerId").value,' +
    'company:document.getElementById("company").value,content:document.getElementById("content").value,' +
    'amount:document.getElementById("amount").value,status:document.getElementById("status").value,' +
    'memo:document.getElementById("memo").value};' +
    'if(!d.content){alert("商談内容を入力してください");return;}' +
    'google.script.run.withSuccessHandler(function(r){var el=document.getElementById("result");el.style.display="block";' +
    'el.className=r.success?"ok":"ng";el.textContent=r.message;}).addDeal(d);}</script>';
}
