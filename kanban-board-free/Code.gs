/**
 * ============================================================
 *  GASTaskBoard 無料版 v1.0
 *  スプレッドシートベースのカンバンボード式タスク管理ツール
 *  TODO / 進行中 / 完了 の3レーン + 期限管理 + メール通知
 * ============================================================
 *
 * © 2026 ゆゆくま / GASおじ All Rights Reserved.
 */

// ========== 定数 ==========
var FREE_MAX_TASKS = 20;
var SH_TASKS = 'タスク一覧';
var SH_BOARD = 'カンバンボード';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';

var ST_TODO = 'TODO';
var ST_DOING = '進行中';
var ST_DONE = '完了';

var PRI_HIGH = '高';
var PRI_MID = '中';
var PRI_LOW = '低';

// ========== メニュー ==========

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 GASTaskBoard')
    .addItem('➕ タスク追加', 'addTask')
    .addItem('▶️ TODO → 進行中', 'moveToInProgress')
    .addItem('✅ 進行中 → 完了', 'moveToDone')
    .addItem('↩️ ステータス変更', 'changeStatus')
    .addSeparator()
    .addItem('📊 カンバンボード更新', 'refreshBoard')
    .addItem('⏰ 期限チェック＆通知', 'checkDeadlines')
    .addItem('⏰ 自動トリガー設定', 'setupTrigger')
    .addItem('⏰ 自動トリガー解除', 'removeTrigger')
    .addSeparator()
    .addItem('🗑️ 完了タスク削除', 'deleteDoneTasks')
    .addItem('ℹ️ 使い方', 'showHelp')
    .addToUi();
}

// ========== タスク追加 ==========

function addTask() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH_TASKS);
  if (!sheet) { ui.alert('setupを先に実行してください。'); return; }

  // 件数チェック
  var lastRow = sheet.getLastRow();
  var activeCount = 0;
  if (lastRow >= 2) {
    var statuses = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    for (var i = 0; i < statuses.length; i++) {
      if (statuses[i][0] && statuses[i][0] !== ST_DONE) activeCount++;
    }
  }
  var totalCount = lastRow >= 2 ? lastRow - 1 : 0;
  if (totalCount >= FREE_MAX_TASKS) {
    ui.alert('⚠️ 無料版はタスク' + FREE_MAX_TASKS + '件までです。\n完了タスクを削除するか、Pro版をご利用ください。');
    return;
  }

  // タスク名
  var nameResult = ui.prompt('➕ タスク追加', 'タスク名を入力:', ui.ButtonSet.OK_CANCEL);
  if (nameResult.getSelectedButton() !== ui.Button.OK || !nameResult.getResponseText().trim()) return;
  var taskName = nameResult.getResponseText().trim();

  // 説明
  var descResult = ui.prompt('📝 説明', '説明を入力（省略可）:', ui.ButtonSet.OK_CANCEL);
  var desc = descResult.getSelectedButton() === ui.Button.OK ? descResult.getResponseText().trim() : '';

  // 期限
  var deadlineResult = ui.prompt('📅 期限', '期限を入力（例: 2026/03/15、省略可）:', ui.ButtonSet.OK_CANCEL);
  var deadline = '';
  if (deadlineResult.getSelectedButton() === ui.Button.OK && deadlineResult.getResponseText().trim()) {
    var d = new Date(deadlineResult.getResponseText().trim());
    if (!isNaN(d.getTime())) deadline = d;
    else { ui.alert('⚠️ 日付形式が不正です。期限なしで登録します。'); }
  }

  // 優先度
  var priResult = ui.prompt('🔥 優先度', '1: 高 / 2: 中 / 3: 低（デフォルト: 中）:', ui.ButtonSet.OK_CANCEL);
  var pri = PRI_MID;
  if (priResult.getSelectedButton() === ui.Button.OK) {
    var p = priResult.getResponseText().trim();
    if (p === '1') pri = PRI_HIGH;
    else if (p === '3') pri = PRI_LOW;
  }

  // ID生成
  var taskId = 'T-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');

  // 書き込み
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 8).setValues([[taskId, taskName, desc, pri, ST_TODO, deadline, new Date(), '']]);

  // 期限セルの書式
  if (deadline) sheet.getRange(newRow, 6).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(newRow, 7).setNumberFormat('yyyy/MM/dd HH:mm:ss');

  // 優先度に色付け
  applyPriorityColor_(sheet, newRow, pri);

  log_('INFO', 'タスク追加: ' + taskName + ' (' + taskId + ')');
  ui.alert('✅ タスクを追加しました！\n\n' + taskName + '\n優先度: ' + pri + '\nステータス: TODO');
  refreshBoard();
}

function applyPriorityColor_(sheet, row, pri) {
  var color = '#FFFFFF';
  if (pri === PRI_HIGH) color = '#FCE4EC';
  else if (pri === PRI_MID) color = '#FFF8E1';
  else if (pri === PRI_LOW) color = '#E8F5E9';
  sheet.getRange(row, 1, 1, 8).setBackground(color);
}

// ========== ステータス移動 ==========

function moveToInProgress() { moveSelectedTask_(ST_TODO, ST_DOING); }
function moveToDone() { moveSelectedTask_(ST_DOING, ST_DONE); }

function moveSelectedTask_(fromStatus, toStatus) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASKS);
  if (!sheet) return;

  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('タスク行を選択してください。'); return; }

  var currentStatus = sheet.getRange(row, 5).getValue();
  if (currentStatus !== fromStatus) {
    ui.alert('⚠️ この操作は「' + fromStatus + '」のタスクにのみ有効です。\n現在のステータス: ' + currentStatus);
    return;
  }

  sheet.getRange(row, 5).setValue(toStatus);
  if (toStatus === ST_DONE) {
    sheet.getRange(row, 8).setValue(new Date());
    sheet.getRange(row, 8).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  }

  var taskName = sheet.getRange(row, 2).getValue();
  log_('INFO', 'ステータス変更: ' + taskName + ' → ' + toStatus);
  ui.alert('✅ 「' + taskName + '」を「' + toStatus + '」に移動しました。');
  refreshBoard();
}

function changeStatus() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASKS);
  if (!sheet) return;

  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('タスク行を選択してください。'); return; }

  var taskName = sheet.getRange(row, 2).getValue();
  var result = ui.prompt('↩️ ステータス変更', '「' + taskName + '」の新しいステータス:\n1: TODO / 2: 進行中 / 3: 完了', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var choice = result.getResponseText().trim();
  var newStatus = '';
  if (choice === '1') newStatus = ST_TODO;
  else if (choice === '2') newStatus = ST_DOING;
  else if (choice === '3') newStatus = ST_DONE;
  else { ui.alert('無効な選択です。'); return; }

  sheet.getRange(row, 5).setValue(newStatus);
  if (newStatus === ST_DONE) {
    sheet.getRange(row, 8).setValue(new Date());
    sheet.getRange(row, 8).setNumberFormat('yyyy/MM/dd HH:mm:ss');
  } else {
    sheet.getRange(row, 8).clearContent();
  }

  log_('INFO', 'ステータス変更: ' + taskName + ' → ' + newStatus);
  ui.alert('✅ 「' + taskName + '」を「' + newStatus + '」に変更しました。');
  refreshBoard();
}

// ========== カンバンボード更新 ==========

function refreshBoard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName(SH_TASKS);
  var boardSheet = ss.getSheetByName(SH_BOARD);
  if (!taskSheet || !boardSheet) return;

  // ボードクリア
  boardSheet.clear();

  // ヘッダー
  boardSheet.getRange('A1').setValue('📝 TODO').setFontWeight('bold').setFontSize(14).setBackground('#BBDEFB').setFontColor('#1565C0');
  boardSheet.getRange('D1').setValue('🔄 進行中').setFontWeight('bold').setFontSize(14).setBackground('#FFF9C4').setFontColor('#F57F17');
  boardSheet.getRange('G1').setValue('✅ 完了').setFontWeight('bold').setFontSize(14).setBackground('#C8E6C9').setFontColor('#2E7D32');

  boardSheet.getRange('A2').setValue('タスク名').setFontWeight('bold');
  boardSheet.getRange('B2').setValue('期限').setFontWeight('bold');
  boardSheet.getRange('C2').setValue('優先度').setFontWeight('bold');
  boardSheet.getRange('D2').setValue('タスク名').setFontWeight('bold');
  boardSheet.getRange('E2').setValue('期限').setFontWeight('bold');
  boardSheet.getRange('F2').setValue('優先度').setFontWeight('bold');
  boardSheet.getRange('G2').setValue('タスク名').setFontWeight('bold');
  boardSheet.getRange('H2').setValue('完了日').setFontWeight('bold');

  // 幅設定
  [200, 100, 60, 200, 100, 60, 200, 120].forEach(function(w, i) { boardSheet.setColumnWidth(i + 1, w); });

  var lastRow = taskSheet.getLastRow();
  if (lastRow < 2) return;

  var data = taskSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  var todo = [], doing = [], done = [];
  for (var i = 0; i < data.length; i++) {
    var task = { name: data[i][1], desc: data[i][2], pri: data[i][3], status: data[i][4], deadline: data[i][5], doneDate: data[i][7] };
    if (task.status === ST_TODO) todo.push(task);
    else if (task.status === ST_DOING) doing.push(task);
    else if (task.status === ST_DONE) done.push(task);
  }

  // 優先度ソート
  var priOrder = function(a, b) {
    var order = { '高': 0, '中': 1, '低': 2 };
    return (order[a.pri] || 1) - (order[b.pri] || 1);
  };
  todo.sort(priOrder);
  doing.sort(priOrder);

  // TODO列
  for (var t = 0; t < todo.length; t++) {
    var r = t + 3;
    boardSheet.getRange(r, 1).setValue(todo[t].name);
    boardSheet.getRange(r, 2).setValue(todo[t].deadline ? formatDate_(todo[t].deadline) : '');
    boardSheet.getRange(r, 3).setValue(todo[t].pri);
    // 期限切れ警告
    if (todo[t].deadline) {
      var dl = new Date(todo[t].deadline); dl.setHours(0, 0, 0, 0);
      if (dl < now) boardSheet.getRange(r, 1, 1, 3).setBackground('#FFCDD2');
      else if (dl.getTime() === now.getTime()) boardSheet.getRange(r, 1, 1, 3).setBackground('#FFE0B2');
    }
  }

  // 進行中列
  for (var d = 0; d < doing.length; d++) {
    var r2 = d + 3;
    boardSheet.getRange(r2, 4).setValue(doing[d].name);
    boardSheet.getRange(r2, 5).setValue(doing[d].deadline ? formatDate_(doing[d].deadline) : '');
    boardSheet.getRange(r2, 6).setValue(doing[d].pri);
    if (doing[d].deadline) {
      var dl2 = new Date(doing[d].deadline); dl2.setHours(0, 0, 0, 0);
      if (dl2 < now) boardSheet.getRange(r2, 4, 1, 3).setBackground('#FFCDD2');
      else if (dl2.getTime() === now.getTime()) boardSheet.getRange(r2, 4, 1, 3).setBackground('#FFE0B2');
    }
  }

  // 完了列
  for (var dn = 0; dn < done.length; dn++) {
    var r3 = dn + 3;
    boardSheet.getRange(r3, 7).setValue(done[dn].name);
    boardSheet.getRange(r3, 8).setValue(done[dn].doneDate ? formatDate_(done[dn].doneDate) : '');
    boardSheet.getRange(r3, 7, 1, 2).setFontColor('#999999');
  }

  log_('INFO', 'カンバンボード更新 - TODO:' + todo.length + ' 進行中:' + doing.length + ' 完了:' + done.length);
}

// ========== 期限チェック＆メール通知 ==========

function checkDeadlines() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH_TASKS);
  if (!sheet || sheet.getLastRow() < 2) return;

  var notifyEmail = getConfig_('通知メール') || Session.getActiveUser().getEmail();
  if (!notifyEmail) {
    log_('WARN', '通知メールが未設定');
    return;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var now = new Date();
  now.setHours(0, 0, 0, 0);
  var tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);

  var overdue = [];
  var dueTomorrow = [];

  for (var i = 0; i < data.length; i++) {
    var status = data[i][4];
    var deadline = data[i][5];
    var taskName = data[i][1];

    if (status === ST_DONE || !deadline || !taskName) continue;

    var dl = new Date(deadline);
    dl.setHours(0, 0, 0, 0);

    if (dl < now) {
      overdue.push({ name: taskName, deadline: formatDate_(dl), status: status });
    } else if (dl.getTime() === tomorrow.getTime()) {
      dueTomorrow.push({ name: taskName, deadline: formatDate_(dl), status: status });
    }
  }

  if (overdue.length === 0 && dueTomorrow.length === 0) {
    log_('INFO', '期限チェック: アラートなし');
    try { SpreadsheetApp.getUi().alert('✅ 期限切れ・期限間近のタスクはありません。'); } catch(e) {}
    return;
  }

  // メール本文作成
  var body = '📋 GASTaskBoard - 期限アラート\n\n';

  if (overdue.length > 0) {
    body += '🚨 期限切れタスク（' + overdue.length + '件）:\n';
    overdue.forEach(function(t) { body += '  ❌ ' + t.name + '（期限: ' + t.deadline + ' / ' + t.status + '）\n'; });
    body += '\n';
  }

  if (dueTomorrow.length > 0) {
    body += '⚠️ 明日が期限のタスク（' + dueTomorrow.length + '件）:\n';
    dueTomorrow.forEach(function(t) { body += '  📅 ' + t.name + '（期限: ' + t.deadline + ' / ' + t.status + '）\n'; });
    body += '\n';
  }

  body += '\nスプレッドシート: ' + ss.getUrl();

  try {
    GmailApp.sendEmail(notifyEmail, '📋 GASTaskBoard 期限アラート', body, { name: 'GASTaskBoard' });
    log_('INFO', '期限通知メール送信: 期限切れ' + overdue.length + '件, 明日期限' + dueTomorrow.length + '件');
  } catch(e) {
    log_('ERROR', 'メール送信エラー: ' + e.message);
  }

  try {
    SpreadsheetApp.getUi().alert(
      '📋 期限チェック完了！\n\n🚨 期限切れ: ' + overdue.length + '件\n⚠️ 明日期限: ' + dueTomorrow.length + '件\n\n通知メールを送信しました。'
    );
  } catch(e) {}
}

// ========== 完了タスク削除 ==========

function deleteDoneTasks() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASKS);
  if (!sheet || sheet.getLastRow() < 2) { ui.alert('タスクがありません。'); return; }

  var result = ui.alert('🗑️ 完了タスク削除', '完了済みタスクをすべて削除しますか？', ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) return;

  var lastRow = sheet.getLastRow();
  var deleted = 0;
  for (var i = lastRow; i >= 2; i--) {
    if (sheet.getRange(i, 5).getValue() === ST_DONE) {
      sheet.deleteRow(i);
      deleted++;
    }
  }

  log_('INFO', '完了タスク削除: ' + deleted + '件');
  ui.alert('✅ ' + deleted + '件の完了タスクを削除しました。');
  refreshBoard();
}

// ========== ヘルプ ==========

function showHelp() {
  SpreadsheetApp.getUi().alert(
    '📋 GASTaskBoard 無料版 — 使い方\n\n' +
    '【タスク管理】\n' +
    '1. メニュー → ➕ タスク追加\n' +
    '2. タスク行を選択 → ▶️ TODO → 進行中\n' +
    '3. タスク行を選択 → ✅ 進行中 → 完了\n\n' +
    '【カンバンボード】\n' +
    '📊 カンバンボード更新 で最新表示\n\n' +
    '【期限通知】\n' +
    '⏰ 期限チェック＆通知 → メールで通知\n' +
    '⏰ 自動トリガー設定 → 1時間ごと自動チェック\n\n' +
    '※ 無料版はタスク' + FREE_MAX_TASKS + '件まで\n' +
    '🔥 Pro版: 無制限タスク＋AI分解＋マルチ通知！'
  );
}
