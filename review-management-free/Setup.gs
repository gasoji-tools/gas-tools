/**
 * ============================================================
 *  GASReviewManager Free - 初期セットアップ
 * ============================================================
 * © 2026 ゆゆくま / GASおじ All Rights Reserved.
 */

// ========== セットアップ ==========

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 口コミ一覧
  var reviewSheet = ss.getSheetByName(SH_REVIEWS);
  if (!reviewSheet) reviewSheet = ss.insertSheet(SH_REVIEWS);
  
  reviewSheet.getRange(1, 1, 1, 8).setValues([['投稿者名', '投稿日', '星評価', '口コミ内容', '口コミURL', '返信案', 'ステータス', 'メモ']]);
  reviewSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#FF9800').setFontColor('#FFFFFF');
  [120, 120, 80, 400, 250, 400, 120, 200].forEach(function(w, i) { reviewSheet.setColumnWidth(i + 1, w); });
  
  var ratingRule = SpreadsheetApp.newDataValidation().requireValueInList(['1', '2', '3', '4', '5']).build();
  reviewSheet.getRange(2, 3, 200, 1).setDataValidation(ratingRule);
  
  var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['未対応', '返信案作成中', '✅ 返信済み', '対応不要']).build();
  reviewSheet.getRange(2, 7, 200, 1).setDataValidation(statusRule);
  
  reviewSheet.getRange(2, 1, 1, 8).setValues([['テスト太郎', '2026/02/27', '5', 'とても良いお店でした！また来たいです。', '', '', '未対応', '']]);
  
  // 返信テンプレート
  var templateSheet = ss.getSheetByName(SH_REPLY_TEMPLATES);
  if (!templateSheet) templateSheet = ss.insertSheet(SH_REPLY_TEMPLATES);
  
  templateSheet.getRange(1, 1, 1, 3).setValues([['星評価', 'テンプレート名', '返信文']]);
  templateSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4CAF50').setFontColor('#FFFFFF');
  [80, 200, 500].forEach(function(w, i) { templateSheet.setColumnWidth(i + 1, w); });
  
  templateSheet.getRange(2, 1, 5, 3).setValues([
    [5, '高評価お礼', '{投稿者名}様、素敵な口コミをありがとうございます！✨ スタッフ一同大変嬉しく思います。{店舗名}では、お客様に最高の体験をお届けできるよう日々努めております。またのご来店を心よりお待ちしております！'],
    [4, '高評価お礼', '{投稿者名}様、ご来店ありがとうございます！高い評価をいただき嬉しいです😊 さらにご満足いただけるよう、サービス向上に努めてまいります。'],
    [3, '普通評価', '{投稿者名}様、ご来店ありがとうございます。貴重なご意見に感謝いたします。より良いサービスをご提供できるよう改善してまいります。'],
    [2, '低評価お詫び', '{投稿者名}様、ご来店ありがとうございます。ご期待に添えず申し訳ございません。いただいたご意見を真摯に受け止め、改善に努めてまいります。'],
    [1, '最低評価お詫び', '{投稿者名}様、ご来店ありがとうございます。ご不快な思いをおかけし、深くお詫び申し上げます。具体的なご状況を確認させていただきたく、お店に直接ご連絡いただけますと幸いです。']
  ]);
  
  // 設定
  var settingsSheet = ss.getSheetByName(SH_SETTINGS);
  if (!settingsSheet) settingsSheet = ss.insertSheet(SH_SETTINGS);
  
  settingsSheet.getRange(1, 1, 9, 2).setValues([
    ['設定項目', '値'],
    ['店舗名', 'あなたの店舗名'],
    ['通知先メール', ''],
    ['Googleマップ URL', ''],
    ['--- 統計（自動更新）---', ''],
    ['総口コミ数', ''],
    ['平均評価', ''],
    ['返信済み', ''],
    ['未返信', '']
  ]);
  settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#F44336').setFontColor('#FFFFFF');
  settingsSheet.getRange(5, 1).setFontWeight('bold').setBackground('#9E9E9E').setFontColor('#FFFFFF');
  settingsSheet.setColumnWidth(1, 200);
  settingsSheet.setColumnWidth(2, 400);
  
  // ログ
  var logSheet = ss.getSheetByName(SH_LOG);
  if (!logSheet) { logSheet = ss.insertSheet(SH_LOG); logSheet.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }
  
  log_('INFO', 'GASReviewManager セットアップ完了');
  SpreadsheetApp.getUi().alert('✅ GASReviewManager セットアップ完了！\n\n口コミをコピペで登録して、返信テンプレートを活用してください。');
}
