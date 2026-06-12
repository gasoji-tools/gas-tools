/**
 * DriveFileList Free
 *
 * Lists files directly inside one Google Drive folder into the active
 * spreadsheet. This free version does not scan subfolders.
 */
function onOpen() {
  getOrCreateSettingsSheet_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi()
    .createMenu('DriveFileList Free')
    .addItem('ファイル一覧を更新', 'updateDriveFileList')
    .addToUi();
}

function updateDriveFileList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getOrCreateSettingsSheet_(ss);
  const folderId = String(settings.getRange('B2').getValue()).trim();

  if (!folderId) {
    ss.toast('設定シートのB2に、一覧化したいGoogle DriveフォルダIDを入力してください。', 'DriveFileList Free', 8);
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const rows = [];

  while (files.hasNext()) {
    const file = files.next();
    rows.push([
      file.getName(),
      file.getLastUpdated(),
      file.getMimeType(),
      file.getUrl(),
      getOwnerEmail_(file)
    ]);
  }

  writeFileList_(ss, rows);
  ss.toast(rows.length + '件のファイルを一覧化しました。', 'DriveFileList Free', 5);
}

function getOrCreateSettingsSheet_(ss) {
  const sheet = ss.getSheetByName('設定') || ss.insertSheet('設定');
  sheet.getRange('A1:B3').setValues([
    ['項目', '値'],
    ['フォルダID', sheet.getRange('B2').getValue()],
    ['メモ', 'Free版は指定フォルダ直下のファイルだけを一覧化します。']
  ]);
  sheet.getRange('A1:B1').setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
  return sheet;
}

function writeFileList_(ss, rows) {
  const sheet = ss.getSheetByName('ファイル一覧') || ss.insertSheet('ファイル一覧');
  sheet.clear();
  sheet.getRange(1, 1, 1, 5).setValues([[
    'ファイル名',
    '最終更新日',
    '種類',
    'URL',
    '所有者メール（取得できる場合）'
  ]]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);
}

function getOwnerEmail_(file) {
  try {
    const owner = file.getOwner();
    return owner ? owner.getEmail() : '';
  } catch (e) {
    return '';
  }
}
