/**
 * GASCRM 無料版 - セットアップ機能
 * © 2026 ゆゆくま / GASおじ All Rights Reserved.
 */

// ========== 初期セットアップ ==========

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 顧客一覧
  var sh = ss.getSheetByName(SH_CUSTOMERS);
  if (!sh) sh = ss.insertSheet(SH_CUSTOMERS);
  sh.getRange('A1:L1').setValues([['ID', '会社名', '担当者名', 'メール', '電話', '住所', 'ランク', 'ステータス', '最終コンタクト', '次回アクション', '登録日', 'メモ']]);
  sh.setFrozenRows(1);

  var rankRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([RANK_A, RANK_B, RANK_C]).setAllowInvalid(false).build();
  sh.getRange('G2:G200').setDataValidation(rankRule);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([DEAL_LEAD, DEAL_CONTACT, DEAL_PROPOSAL, DEAL_NEGOTIATION, DEAL_WON, DEAL_LOST])
    .setAllowInvalid(false).build();
  sh.getRange('H2:H200').setDataValidation(statusRule);

  // 商談履歴
  var deals = ss.getSheetByName(SH_DEALS);
  if (!deals) deals = ss.insertSheet(SH_DEALS);
  deals.getRange('A1:H1').setValues([['ID', '顧客ID', '会社名', '商談内容', '金額', 'ステータス', '日時', 'メモ']]);
  deals.setFrozenRows(1);

  // 設定
  var set = ss.getSheetByName(SH_SETTINGS);
  if (!set) {
    set = ss.insertSheet(SH_SETTINGS);
    set.getRange('A1:B1').setValues([['項目', '値']]);
    set.getRange('A2:B4').setValues([
      ['通知メール', ''],
      ['フォローアップ日数', '7'],
      ['デフォルトランク', 'B（標準）']
    ]);
  }

  // ログ
  var lg = ss.getSheetByName(SH_LOG);
  if (!lg) { lg = ss.insertSheet(SH_LOG); lg.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }

  log_('INFO', 'セットアップ完了');
  SpreadsheetApp.getUi().alert('GASCRM 無料版のセットアップが完了しました！');
}
