/**
 * Mini Shop Order Tracker Free v1.0.0
 * 個人販売の注文を Google シート1枚で管理し、未発送・未入金を見える化するミニツール
 *
 * © 2026 ゆゆくま / Gasおじラボ
 *
 * できること:
 *   - Orders シート1枚で 注文日 / 購入者メモ / 商品名 / 金額 / 入金状態 / 発送状態 / 発送予定日 / メモ を一覧化
 *   - 手動実行で「未発送一覧」「未入金一覧」をまとめた自分用リマインドの下書きを1通作成
 *   - 宛先は自分（スクリプト実行者のメール）。自動送信はしません。
 *
 * 使い方:
 *   1. initMiniShopOrderTrackerFree を実行してサンプルシートを作る
 *   2. 注文を自分用に書き換える（BOOTH / BASE / メルカリ / DM / イベントなど経路はメモ欄へ）
 *   3. createOrderSummaryDraft を実行して未発送・未入金のリマインド下書きを作る
 *   4. Gmail の下書きを開いて内容を確認し、自分で送信する（または見るだけでもOK）
 *
 * 注意:
 *   - これは個人販売者向けの小さな注文台帳です。法人ECの基幹管理ではありません。
 *   - 顧客の氏名・住所・連絡先などの個人情報は、購入者メモ欄に最小限だけ自己管理してください。
 */

var SHEET_NAME = 'Orders';
var HEADER = ['注文日', '購入者メモ', '商品名', '金額', '入金状態', '発送状態', '発送予定日', 'メモ'];
var PAID_DONE = '入金済み';
var SHIP_DONE = '発送済み';

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('注文台帳 Free')
      .addItem('初期シートを作成', 'initMiniShopOrderTrackerFree')
      .addItem('未発送・未入金リマインドを作成', 'createOrderSummaryDraft')
      .addToUi();
  } catch (e) {
    Logger.log('UI menu skipped: ' + e.message);
  }
}

function initMiniShopOrderTrackerFree() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  sheet.getRange(2, 1, 3, HEADER.length).setValues([
    [todayString_(), 'BOOTH 注文 / にゃさん', 'アクキー 2点', 1600, '入金済み', '未発送', todayString_(), '匿名配送'],
    [todayString_(), 'DM 取り置き / みかんさん', 'ステッカーセット', 800, '未入金', '未発送', '', '入金待ち'],
    [todayString_(), 'イベント手渡し', 'ポストカード 5枚', 500, '入金済み', '発送済み', '', '当日完了']
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADER.length);
  showMessage_('初期シートを作成しました',
    '注文を自分用に書き換えてから、「未発送・未入金リマインドを作成」を実行してください。\n' +
    '入金状態は「未入金 / 入金済み」、発送状態は「未発送 / 梱包済み / 発送済み」で入力します。');
}

function createOrderSummaryDraft() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    showMessage_('シートがありません', '先に initMiniShopOrderTrackerFree を実行してください。');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    showMessage_('注文がありません', 'Orders シートに注文を入力してください。');
    return;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  var today = startOfDay_(new Date());
  var unshipped = [];
  var unpaid = [];

  for (var i = 0; i < values.length; i++) {
    var product = String(values[i][2] || '').trim();
    if (!product) continue;
    var buyer = String(values[i][1] || '').trim();
    var amount = values[i][3];
    var paidStatus = String(values[i][4] || '').trim();
    var shipStatus = String(values[i][5] || '').trim();

    var label = '・' + product + (buyer ? '（' + buyer + '）' : '');

    if (shipStatus !== SHIP_DONE) {
      unshipped.push(label + shipDueLabel_(values[i][6], today, shipStatus));
    }
    if (paidStatus !== PAID_DONE) {
      unpaid.push(label + amountLabel_(amount));
    }
  }

  if (unshipped.length === 0 && unpaid.length === 0) {
    showMessage_('未対応の注文はありません', '未発送・未入金の注文はありませんでした。お疲れさまでした。');
    return;
  }

  var dateText = todayString_();
  var subject = '【注文台帳】未発送' + unshipped.length + '件 / 未入金' + unpaid.length + '件（' + dateText + '）';

  var parts = [dateText + ' 時点の注文台帳リマインド'];
  parts.push('\n■ 未発送（' + unshipped.length + '件）');
  parts.push(unshipped.length ? unshipped.join('\n') : '（なし）');
  parts.push('\n■ 未入金（' + unpaid.length + '件）');
  parts.push(unpaid.length ? unpaid.join('\n') : '（なし）');
  parts.push('\n──────────\nこのメールは「注文台帳 Free」が作成した自分用リマインドです。\n送信前に内容を確認してください。');
  var body = parts.join('\n');

  var to = selfEmail_();
  if (!to) {
    showMessage_('宛先を特定できません',
      '実行者のメールアドレスを取得できませんでした。Googleアカウントでログインした状態で、もう一度実行してください。');
    return;
  }
  GmailApp.createDraft(to, subject, body);

  showMessage_('リマインド下書きを作成しました',
    '未発送 ' + unshipped.length + '件 / 未入金 ' + unpaid.length + '件 のリマインドを下書きにしました。\n' +
    'Gmail の下書きフォルダを開いて内容を確認してください。');
}

function shipDueLabel_(rawDue, today, shipStatus) {
  var statusTag = shipStatus === '梱包済み' ? '[梱包済み]' : '';
  var due = toDate_(rawDue);
  if (!due) return statusTag ? '（' + statusTag.replace(/[\[\]]/g, '') + '）' : '';
  var d = startOfDay_(due);
  var dueText;
  if (d.getTime() < today.getTime()) dueText = '発送予定 超過: ' + formatDate_(d);
  else if (d.getTime() === today.getTime()) dueText = '本日発送予定';
  else dueText = '発送予定: ' + formatDate_(d);
  var inner = statusTag ? statusTag.replace(/[\[\]]/g, '') + ' / ' + dueText : dueText;
  return '（' + inner + '）';
}

function amountLabel_(amount) {
  var n = Number(amount);
  if (isNaN(n) || !amount) return '';
  return '（' + yen_(n) + '）';
}

function yen_(n) {
  return '¥' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function selfEmail_() {
  var email = '';
  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (e) {
    email = '';
  }
  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail() || '';
    } catch (e2) {
      email = '';
    }
  }
  return email;
}

function todayString_() {
  return formatDate_(new Date());
}

function formatDate_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  } catch (e) {
    var y = date.getFullYear();
    var m = ('0' + (date.getMonth() + 1)).slice(-2);
    var d = ('0' + date.getDate()).slice(-2);
    return y + '/' + m + '/' + d;
  }
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toDate_(value) {
  if (value instanceof Date) return value;
  var s = String(value || '').trim();
  if (!s) return null;
  var parsed = new Date(s.replace(/-/g, '/'));
  return isNaN(parsed.getTime()) ? null : parsed;
}

function showMessage_(title, body) {
  try {
    SpreadsheetApp.getUi().alert(title, body, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(title + ': ' + body);
  }
}
