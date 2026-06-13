/**
 * ============================================================
 *  GASTaskBoard 無料版 - セットアップ
 * ============================================================
 *
 * © 2026 ゆゆくま / GASおじ All Rights Reserved.
 */

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

function formatDate_(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function formatDateTime_(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

// ========== トリガー管理 ==========

function setupTrigger() {
  removeTriggerSilent_();
  ScriptApp.newTrigger('checkDeadlines').timeBased().everyHours(1).create();
  log_('INFO', '自動トリガー設定（1時間間隔）');
  SpreadsheetApp.getUi().alert('⏰ 自動トリガーを設定しました！\n1時間ごとに期限をチェックして通知します。');
}

function removeTrigger() {
  removeTriggerSilent_();
  log_('INFO', '自動トリガー解除');
  SpreadsheetApp.getUi().alert('⏰ 自動トリガーを解除しました。');
}

function removeTriggerSilent_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkDeadlines') ScriptApp.deleteTrigger(trigger);
  });
}

// ========== セットアップ ==========

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // タスク一覧シート
  var taskSheet = ss.getSheetByName(SH_TASKS);
  if (!taskSheet) taskSheet = ss.insertSheet(SH_TASKS);

  taskSheet.getRange(1, 1, 1, 8).setValues([['タスクID', 'タスク名', '説明', '優先度', 'ステータス', '期限', '作成日', '完了日']]);
  taskSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  [120, 250, 300, 60, 80, 100, 140, 140].forEach(function(w, i) { taskSheet.setColumnWidth(i + 1, w); });

  // データバリデーション
  var priRule = SpreadsheetApp.newDataValidation().requireValueInList([PRI_HIGH, PRI_MID, PRI_LOW]).build();
  taskSheet.getRange(2, 4, 100, 1).setDataValidation(priRule);

  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList([ST_TODO, ST_DOING, ST_DONE]).build();
  taskSheet.getRange(2, 5, 100, 1).setDataValidation(statusRule);

  // サンプルタスク
  var now = new Date();
  var sampleDeadline = new Date(now);
  sampleDeadline.setDate(sampleDeadline.getDate() + 3);
  taskSheet.getRange(2, 1, 1, 8).setValues([['T-SAMPLE001', 'サンプルタスク', 'これはサンプルです。編集・削除してOK', PRI_MID, ST_TODO, sampleDeadline, now, '']]);
  taskSheet.getRange(2, 6).setNumberFormat('yyyy/MM/dd');
  taskSheet.getRange(2, 7).setNumberFormat('yyyy/MM/dd HH:mm:ss');

  taskSheet.getRange(22, 1).setValue('※ 無料版はタスク20件まで。Pro版は無制限＋AI分解＋マルチ通知！').setFontColor('#999999').setFontStyle('italic');

  // カンバンボードシート
  var boardSheet = ss.getSheetByName(SH_BOARD);
  if (!boardSheet) boardSheet = ss.insertSheet(SH_BOARD);

  // 設定シート
  var settingsSheet = ss.getSheetByName(SH_SETTINGS);
  if (!settingsSheet) settingsSheet = ss.insertSheet(SH_SETTINGS);
  settingsSheet.getRange(1, 1, 3, 2).setValues([
    ['設定項目', '値'],
    ['通知メール', Session.getActiveUser().getEmail() || 'your-email@gmail.com'],
    ['プロジェクト名', 'マイプロジェクト']
  ]);
  settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  settingsSheet.setColumnWidth(1, 150);
  settingsSheet.setColumnWidth(2, 300);

  // ログシート
  var logSheet = ss.getSheetByName(SH_LOG);
  if (!logSheet) { logSheet = ss.insertSheet(SH_LOG); logSheet.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }

  // ボード初期化
  refreshBoard();

  log_('INFO', 'GASTaskBoard 無料版 セットアップ完了');
  SpreadsheetApp.getUi().alert('✅ GASTaskBoard セットアップ完了！\n\nメニュー → 📋 GASTaskBoard からタスクを管理できます。');
}
