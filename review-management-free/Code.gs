/**
 * ============================================================
 *  GASReviewManager 無料版 v1.0
 *  Googleマップ口コミ管理ツール（手動入力方式・API不要）
 * ============================================================
 * 
 * シート構成:
 * - 「口コミ一覧」: 口コミの登録・管理
 * - 「返信テンプレート」: 星評価別の返信テンプレート
 * - 「設定」: 店舗情報
 * - 「ログ」: 操作ログ
 * 
 * © 2026 ゆゆくま / GASおじ All Rights Reserved.
 */

// ========== 定数 ==========
var SH_REVIEWS = '口コミ一覧';
var SH_REPLY_TEMPLATES = '返信テンプレート';
var SH_SETTINGS = '設定';
var SH_LOG = 'ログ';

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

// ========== メニュー ==========

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⭐ GASReviewManager')
    .addItem('📝 返信テンプレート適用', 'applyReplyTemplate')
    .addItem('📊 口コミサマリー', 'showReviewSummary')
    .addItem('📧 未返信リスト通知（メール）', 'notifyUnrepliedByEmail')
    .addSeparator()
    .addItem('🔄 統計更新', 'updateStats')
    .addItem('ℹ️ 使い方', 'showHelp')
    .addToUi();
}

// ========== 口コミサマリー ==========

function showReviewSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH_REVIEWS);
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('口コミがまだ登録されていません。');
    return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var total = data.filter(function(r) { return r[0]; }).length;
  var ratings = data.filter(function(r) { return r[2]; }).map(function(r) { return Number(r[2]); });
  var avgRating = ratings.length > 0 ? (ratings.reduce(function(a, b) { return a + b; }, 0) / ratings.length).toFixed(1) : 'N/A';
  
  var starCounts = [0, 0, 0, 0, 0];
  ratings.forEach(function(r) { if (r >= 1 && r <= 5) starCounts[r - 1]++; });
  
  var replied = data.filter(function(r) { return r[6] === '✅ 返信済み'; }).length;
  
  log_('INFO', '口コミサマリー表示 - 合計:' + total + ' 平均:' + avgRating);
  
  SpreadsheetApp.getUi().alert(
    '⭐ 口コミサマリー\n\n' +
    '総口コミ数: ' + total + '件\n平均評価: ' + avgRating + '\n\n' +
    '⭐5: ' + starCounts[4] + '件 / ⭐4: ' + starCounts[3] + '件 / ⭐3: ' + starCounts[2] + '件\n' +
    '⭐2: ' + starCounts[1] + '件 / ⭐1: ' + starCounts[0] + '件\n\n' +
    '返信済み: ' + replied + '件\n未返信: ' + (total - replied) + '件'
  );
}

// ========== 返信テンプレート適用 ==========

function applyReplyTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SH_REVIEWS);
  var templateSheet = ss.getSheetByName(SH_REPLY_TEMPLATES);
  
  if (!templateSheet || templateSheet.getLastRow() < 2) {
    ui.alert('返信テンプレートが登録されていません。');
    return;
  }
  
  var row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert('口コミ行を選択してから実行してください。'); return; }
  
  var rating = sheet.getRange(row, 3).getValue();
  var reviewerName = sheet.getRange(row, 1).getValue();
  
  if (!rating) { ui.alert('星評価が入力されていません。'); return; }
  
  var templates = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 3).getValues();
  var matched = templates.filter(function(t) { return Number(t[0]) === Number(rating); });
  
  if (matched.length === 0) {
    ui.alert('星' + rating + 'のテンプレートが見つかりません。');
    return;
  }
  
  var template = matched[Math.floor(Math.random() * matched.length)];
  var shopName = getConfig_('店舗名') || '当店';
  var replyText = template[2]
    .replace(/\{投稿者名\}/g, escHtml_(reviewerName) || 'お客様')
    .replace(/\{店舗名\}/g, escHtml_(shopName));
  
  sheet.getRange(row, 6).setValue(replyText);
  log_('INFO', 'テンプレート適用: 星' + rating + ' → 行' + row);
  ui.alert('✅ 星' + rating + 'のテンプレートを適用しました！\nF列を確認・編集してください。');
}

// ========== メール通知 ==========

function notifyUnrepliedByEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var notifyEmail = getConfig_('通知先メール');
  var sheet = ss.getSheetByName(SH_REVIEWS);
  
  if (!notifyEmail) {
    SpreadsheetApp.getUi().alert('通知先メールアドレスが未設定です。\n「設定」シートに入力してください。');
    return;
  }
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('口コミが登録されていません。');
    return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var unreplied = data.filter(function(r) { return r[0] && r[6] !== '✅ 返信済み'; });
  
  if (unreplied.length === 0) {
    SpreadsheetApp.getUi().alert('🎉 未返信の口コミはありません！');
    return;
  }
  
  var shopName = getConfig_('店舗名') || '店舗';
  var body = '【' + shopName + '】未返信の口コミが' + unreplied.length + '件あります。\n\n';
  unreplied.forEach(function(r, i) {
    body += '---\n' + (i + 1) + '. ' + r[0] + 'さん（⭐' + r[2] + '）\n「' + r[3] + '」\n投稿日: ' + r[1] + '\n\n';
  });
  body += '\nスプレッドシート: ' + ss.getUrl();
  
  GmailApp.sendEmail(notifyEmail, '【口コミ未返信】' + unreplied.length + '件が返信待ちです', body, { name: 'GASReviewManager' });
  log_('INFO', '未返信通知メール送信: ' + unreplied.length + '件');
  SpreadsheetApp.getUi().alert('📧 未返信' + unreplied.length + '件の通知を送信しました！');
}

// ========== 統計更新 ==========

function updateStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SH_REVIEWS);
  var settingsSheet = ss.getSheetByName(SH_SETTINGS);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  var total = data.filter(function(r) { return r[0]; }).length;
  var ratings = data.filter(function(r) { return r[2]; }).map(function(r) { return Number(r[2]); });
  var avg = ratings.length > 0 ? (ratings.reduce(function(a, b) { return a + b; }, 0) / ratings.length).toFixed(1) : 0;
  var replied = data.filter(function(r) { return r[6] === '✅ 返信済み'; }).length;
  
  if (settingsSheet) {
    settingsSheet.getRange('B6').setValue(total);
    settingsSheet.getRange('B7').setValue(avg);
    settingsSheet.getRange('B8').setValue(replied);
    settingsSheet.getRange('B9').setValue(total - replied);
  }
  
  log_('INFO', '統計更新: 合計' + total + ' 平均' + avg + ' 返信済' + replied);
  try { SpreadsheetApp.getUi().alert('📊 統計更新完了！\n口コミ: ' + total + '件 / 平均: ' + avg + ' / 返信済: ' + replied + ' / 未返信: ' + (total - replied)); } catch(e) {}
}

// ========== ヘルプ ==========

function showHelp() {
  SpreadsheetApp.getUi().alert(
    '⭐ GASReviewManager 無料版 — 使い方\n\n' +
    '【口コミ登録】Googleマップから口コミをコピペ入力\n' +
    '【返信案作成】行選択 → メニュー → 📝 返信テンプレート適用\n' +
    '【返信投稿】F列の返信案をコピー → Googleマップで貼り付け投稿\n' +
    '【通知】メニュー → 📧 未返信リスト通知\n' +
    '【統計】メニュー → 📊 口コミサマリー\n\n' +
    '🔥 Pro版: Google API自動取得/AI返信案/Discord・LINE通知/多店舗対応！'
  );
}
