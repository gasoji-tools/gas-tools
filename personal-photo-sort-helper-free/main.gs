/**
 * Personal Photo Sort Helper Free v1.0.0
 * 指定したGoogle Driveフォルダ直下の写真（画像ファイル）を一覧化するツール
 *
 * © ゆゆくま / Gasおじラボ
 *
 * できること:
 *   - フォルダ直下の画像ファイルを「ファイル名 / 作成日 / 種類 / サイズ(MB) / URL」の5列で一覧化
 *   - 作成日（Driveに登録された日時）の新しい順に並べ替え
 * しないこと:
 *   - ファイルの削除・移動・リネーム・権限変更は一切行いません（読み取りのみ）
 *   - サブフォルダの中までは見ません（直下のみ。複数フォルダ・再帰はPro版）
 *
 * 使い方:
 *   1. 下の FOLDER_ID を、写真が入っているフォルダのIDに書き換える
 *   2. listPhotos を実行する（初回は権限の承認が必要）
 *   詳細は README.md を参照
 */

// ===== 設定（ここだけ書き換えてください） =====

// 写真を一覧化したいフォルダのID
// フォルダをブラウザで開いたときのURL https://drive.google.com/drive/folders/XXXXX の XXXXX 部分
var FOLDER_ID = 'YOUR_FOLDER_ID';

// 出力先のシート名（同名シートがあれば中身を上書きします）
var SHEET_NAME = 'Photos';

// ===== ここから下は書き換え不要 =====

var HEADER = ['ファイル名', '作成日', '種類', 'サイズ(MB)', 'URL'];

// 写真として扱うMIMEタイプ
var IMAGE_MIME_LABELS = {
  'image/jpeg': 'JPEG',
  'image/png': 'PNG',
  'image/gif': 'GIF',
  'image/heic': 'HEIC',
  'image/heif': 'HEIF',
  'image/webp': 'WebP',
  'image/tiff': 'TIFF',
  'image/bmp': 'BMP'
};

/**
 * スプレッドシートを開いたときに専用メニューを追加します。
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Photo Sort Helper')
      .addItem('写真一覧を作成', 'listPhotos')
      .addToUi();
  } catch (e) {
    // エディタから直接実行した場合などUIが使えない環境では何もしない
  }
}

/**
 * メイン処理: フォルダ直下の画像ファイル一覧をシートに書き出します。
 */
function listPhotos() {
  if (!FOLDER_ID || FOLDER_ID === 'YOUR_FOLDER_ID') {
    showMessage_(
      '設定が必要です',
      'コード上部の FOLDER_ID を、写真が入っているフォルダのIDに書き換えてから実行してください。'
    );
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
  } catch (e) {
    showMessage_(
      'フォルダが開けません',
      'FOLDER_ID が正しいか、そのフォルダへのアクセス権があるかを確認してください。\n' +
      '指定中のID: ' + FOLDER_ID
    );
    return;
  }

  var rows = [];
  var skipped = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (!isImage_(file)) {
      skipped++;
      continue;
    }
    rows.push([
      file.getName(),
      file.getDateCreated(),
      imageLabel_(file.getMimeType()),
      sizeMb_(file),
      file.getUrl()
    ]);
  }

  // 作成日の新しい順に並べる
  rows.sort(function (a, b) { return b[1] - a[1]; });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  }
  sheet.setFrozenRows(1);

  showMessage_(
    '完了',
    'フォルダ「' + folder.getName() + '」直下の写真 ' + rows.length + ' 件を' +
    'シート「' + SHEET_NAME + '」に書き出しました。' +
    (skipped > 0 ? '\n（画像以外の ' + skipped + ' 件はスキップしました）' : '')
  );
}

/**
 * 画像ファイルかどうかを判定します。
 */
function isImage_(file) {
  return !!IMAGE_MIME_LABELS[file.getMimeType()];
}

/**
 * 画像のMIMEタイプを短い種類名に変換します。
 */
function imageLabel_(mime) {
  return IMAGE_MIME_LABELS[mime] || mime;
}

/**
 * ファイルサイズをMB単位（小数2桁）で返します。取得できない場合は空文字。
 */
function sizeMb_(file) {
  try {
    var bytes = file.getSize();
    if (!bytes) return '';
    return Math.round(bytes / 1024 / 1024 * 100) / 100;
  } catch (e) {
    return '';
  }
}

/**
 * 実行環境に応じてダイアログまたはログでメッセージを表示します。
 */
function showMessage_(title, body) {
  try {
    SpreadsheetApp.getUi().alert(title, body, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(title + ': ' + body);
  }
}
