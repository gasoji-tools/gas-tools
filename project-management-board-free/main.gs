// © 2026 ゆゆくま / GASおじ All Rights Reserved.
/**
 * ============================================================
 *  GAS Kanban v1.0 — カンバンボードツール（Trello代替）
 *  Setup.gs — 初期設定・WebApp・ユーティリティ
 * ============================================================
 */

// ── シート名定数 ──
var SH_CONFIG  = '設定';
var SH_BOARDS  = 'ボード';
var SH_COLUMNS = 'カラム';
var SH_CARDS   = 'カード';
var SH_LABELS  = 'ラベル';
var SH_LOG     = 'ログ';

// ── メニュー ──
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 Kanban')
    .addItem('🔧 初期設定', 'initialSetup')
    .addItem('🌐 ボードを開く', 'openWebApp')
    .addToUi();
}

// ── 初期設定 ──
function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var cfg = getOrCreateSheet_(ss, SH_CONFIG);
  if (cfg.getLastRow() < 2) {
    cfg.getRange('A1:B1').setValues([['項目', '値']]);
    cfg.getRange('A2:B4').setValues([
      ['デフォルトボード', 'メインボード'],
      ['カードの最大数/カラム', '20'],
      ['通知ON', 'ON']
    ]);
  }

  var boards = getOrCreateSheet_(ss, SH_BOARDS);
  if (boards.getLastRow() < 2) {
    boards.getRange('A1:C1').setValues([['ボードID', 'ボード名', '作成日時']]);
    boards.getRange('A2:C2').setValues([['BD-001', 'メインボード', new Date()]]);
  }

  var cols = getOrCreateSheet_(ss, SH_COLUMNS);
  if (cols.getLastRow() < 2) {
    cols.getRange('A1:D1').setValues([['カラムID', 'ボードID', 'カラム名', '表示順']]);
    cols.getRange('A2:D5').setValues([
      ['CL-001', 'BD-001', 'ToDo', 1],
      ['CL-002', 'BD-001', '進行中', 2],
      ['CL-003', 'BD-001', 'レビュー', 3],
      ['CL-004', 'BD-001', '完了', 4]
    ]);
  }

  var cards = getOrCreateSheet_(ss, SH_CARDS);
  if (cards.getLastRow() < 1) {
    cards.getRange('A1:J1').setValues([[
      'カードID', 'ボードID', 'カラムID', 'タイトル', '説明',
      '担当者', 'ラベル', '期限', '表示順', '作成日時'
    ]]);
  }

  var labels = getOrCreateSheet_(ss, SH_LABELS);
  if (labels.getLastRow() < 2) {
    labels.getRange('A1:C1').setValues([['ラベル名', '色', 'ボードID']]);
    labels.getRange('A2:C5').setValues([
      ['バグ', '#ea4335', 'BD-001'],
      ['機能', '#4285f4', 'BD-001'],
      ['改善', '#34a853', 'BD-001'],
      ['緊急', '#ff6d01', 'BD-001']
    ]);
  }

  var lg = getOrCreateSheet_(ss, SH_LOG);
  if (lg.getLastRow() < 1) {
    lg.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]);
  }

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('WEBHOOK_URL')) props.setProperty('WEBHOOK_URL', '');
  if (!props.getProperty('CARD_COUNTER')) props.setProperty('CARD_COUNTER', '0');
  if (!props.getProperty('COLUMN_COUNTER')) props.setProperty('COLUMN_COUNTER', '4');
  if (!props.getProperty('BOARD_COUNTER')) props.setProperty('BOARD_COUNTER', '1');

  log_('INFO', '初期設定完了');
  SpreadsheetApp.getUi().alert('✅ 初期設定が完了しました！\n\nデフォルトボード「メインボード」を作成しました。');
}

// ── WebApp ──
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('GAS Kanban — カンバンボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openWebApp() {
  SpreadsheetApp.getUi().alert('WebApp URL:\n' + ScriptApp.getService().getUrl());
}

// ── Public Wrappers ──
function getBoards() { return getBoards_(); }
function getBoard(boardId) { return getBoard_(boardId); }
function addBoard(name) { return addBoard_(name); }
function getColumns(boardId) { return getColumns_(boardId); }
function addColumn(boardId, name) { return addColumn_(boardId, name); }
function deleteColumn(colId) { return deleteColumn_(colId); }
function getCards(boardId) { return getCards_(boardId); }
function addCard(boardId, colId, title, desc, assignee, label, deadline) { return addCard_(boardId, colId, title, desc, assignee, label, deadline); }
function updateCard(cardId, title, desc, assignee, label, deadline) { return updateCard_(cardId, title, desc, assignee, label, deadline); }
function moveCard(cardId, newColId) { return moveCard_(cardId, newColId); }
function deleteCard(cardId) { return deleteCard_(cardId); }
function getLabels(boardId) { return getLabels_(boardId); }
function addLabel(boardId, name, color) { return addLabel_(boardId, name, color); }
function getDashboard(boardId) { return getDashboard_(boardId); }

// ── ユーティリティ ──
function getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function getConfig_(key) {
  if (key === 'Webhook URL') {
    return PropertiesService.getScriptProperties().getProperty('WEBHOOK_URL') || '';
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CONFIG);
  if (!sh || sh.getLastRow() < 2) return '';
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

function log_(level, msg) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SH_LOG);
    if (!sh) { sh = ss.insertSheet(SH_LOG); sh.getRange('A1:C1').setValues([['日時', 'レベル', 'メッセージ']]); }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[new Date(), level, msg]]);
    var lastRow = sh.getLastRow();
    if (lastRow > 501) sh.deleteRows(2, lastRow - 501);
  } catch(e) { Logger.log(level + ': ' + msg); }
}

function fetchWithRetry_(url, options, retries) {
  retries = retries || 3;
  for (var i = 0; i < retries; i++) {
    try {
      Utilities.sleep(350);
      var res = UrlFetchApp.fetch(url, options);
      var code = res.getResponseCode();
      if (code >= 200 && code < 300) return res;
      if (i === retries - 1) throw new Error('HTTP ' + code);
    } catch(e) {
      if (i === retries - 1) throw e;
      Utilities.sleep(1000 * (i + 1));
    }
  }
}

function escHtml_(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function nextCardId_() {
  var props = PropertiesService.getScriptProperties();
  var c = parseInt(props.getProperty('CARD_COUNTER') || '0') + 1;
  props.setProperty('CARD_COUNTER', String(c));
  return 'CD-' + ('00000' + c).slice(-5);
}

function nextColumnId_() {
  var props = PropertiesService.getScriptProperties();
  var c = parseInt(props.getProperty('COLUMN_COUNTER') || '0') + 1;
  props.setProperty('COLUMN_COUNTER', String(c));
  return 'CL-' + ('000' + c).slice(-3);
}

function nextBoardId_() {
  var props = PropertiesService.getScriptProperties();
  var c = parseInt(props.getProperty('BOARD_COUNTER') || '0') + 1;
  props.setProperty('BOARD_COUNTER', String(c));
  return 'BD-' + ('000' + c).slice(-3);
}

function notifyWebhook_(message) {
  var url = PropertiesService.getScriptProperties().getProperty('WEBHOOK_URL') || '';
  var notifyOn = getConfig_('通知ON');
  if (!url || notifyOn !== 'ON') return;
  try {
    fetchWithRetry_(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '📋 Kanban: ' + message }),
      muteHttpExceptions: true
    });
  } catch(e) { log_('WARN', 'Webhook失敗: ' + e.message); }
}
/**
 * ============================================================
 *  GAS Kanban v1.0 — カンバンボードツール（Trello代替）
 *  Code.gs — ビジネスロジック
 * ============================================================
 */

// ══════════════════════════════════════════════════
//  ボード管理
// ══════════════════════════════════════════════════

function getBoards_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_BOARDS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return data.filter(function(r) { return r[0]; }).map(function(r) {
    return { id: r[0], name: r[1], created: r[2] };
  });
}

function getBoard_(boardId) {
  boardId = String(boardId || '').trim();
  if (!boardId) throw new Error('ボードIDは必須です');
  var boards = getBoards_();
  for (var i = 0; i < boards.length; i++) {
    if (boards[i].id === boardId) return boards[i];
  }
  throw new Error('ボードが見つかりません: ' + boardId);
}

function addBoard_(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('ボード名は必須です');

  var boards = getBoards_();
  for (var i = 0; i < boards.length; i++) {
    if (boards[i].name === name) throw new Error('同名のボードが既に存在します');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_BOARDS);
  var id = nextBoardId_();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[id, name, new Date()]]);

  // デフォルトカラム作成
  var colSh = ss.getSheetByName(SH_COLUMNS);
  var defaults = [['ToDo', 1], ['進行中', 2], ['レビュー', 3], ['完了', 4]];
  defaults.forEach(function(d) {
    var colId = nextColumnId_();
    colSh.getRange(colSh.getLastRow() + 1, 1, 1, 4).setValues([[colId, id, d[0], d[1]]]);
  });

  log_('INFO', 'ボード作成: ' + name + ' (' + id + ')');
  return { success: true, id: id };
}

// ══════════════════════════════════════════════════
//  カラム管理
// ══════════════════════════════════════════════════

function getColumns_(boardId) {
  boardId = String(boardId || '').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_COLUMNS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var cols = data.filter(function(r) { return r[0] && (!boardId || r[1] === boardId); })
    .map(function(r) { return { id: r[0], boardId: r[1], name: r[2], order: r[3] }; });
  cols.sort(function(a, b) { return a.order - b.order; });
  return cols;
}

function addColumn_(boardId, name) {
  boardId = String(boardId || '').trim();
  if (!boardId) throw new Error('ボードIDは必須です');
  name = String(name || '').trim();
  if (!name) throw new Error('カラム名は必須です');

  var cols = getColumns_(boardId);
  for (var i = 0; i < cols.length; i++) {
    if (cols[i].name === name) throw new Error('同名のカラムが既に存在します');
  }

  var order = cols.length > 0 ? cols[cols.length - 1].order + 1 : 1;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_COLUMNS);
  var id = nextColumnId_();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[id, boardId, name, order]]);
  log_('INFO', 'カラム追加: ' + name);
  return { success: true, id: id };
}

function deleteColumn_(colId) {
  colId = String(colId || '').trim();
  if (!colId) throw new Error('カラムIDは必須です');

  // カラム内のカードチェック
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cardSh = ss.getSheetByName(SH_CARDS);
  if (cardSh && cardSh.getLastRow() >= 2) {
    var cards = cardSh.getRange(2, 1, cardSh.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < cards.length; i++) {
      if (cards[i][2] === colId) throw new Error('カード' + cards[i][0] + 'が存在するカラムは削除できません');
    }
  }

  var sh = ss.getSheetByName(SH_COLUMNS);
  if (!sh || sh.getLastRow() < 2) throw new Error('カラムが見つかりません');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === colId) {
      sh.deleteRow(i + 2);
      log_('INFO', 'カラム削除: ' + colId);
      return { success: true };
    }
  }
  throw new Error('カラムが見つかりません');
}

// ══════════════════════════════════════════════════
//  カード管理
// ══════════════════════════════════════════════════

function getCards_(boardId) {
  boardId = String(boardId || '').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CARDS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  return data.filter(function(r) { return r[0] && (!boardId || r[1] === boardId); })
    .map(function(r) {
      return {
        id: r[0], boardId: r[1], colId: r[2], title: r[3], desc: r[4],
        assignee: r[5], label: r[6], deadline: r[7], order: r[8], created: r[9]
      };
    });
}

function addCard_(boardId, colId, title, desc, assignee, label, deadline) {
  boardId = String(boardId || '').trim();
  if (!boardId) throw new Error('ボードIDは必須です');
  colId = String(colId || '').trim();
  if (!colId) throw new Error('カラムIDは必須です');
  title = String(title || '').trim();
  if (!title) throw new Error('タイトルは必須です');
  desc = String(desc || '').trim();
  assignee = String(assignee || '').trim();
  label = String(label || '').trim();

  // カラム存在チェック
  var cols = getColumns_(boardId);
  var colExists = false;
  for (var i = 0; i < cols.length; i++) {
    if (cols[i].id === colId) { colExists = true; break; }
  }
  if (!colExists) throw new Error('カラムが見つかりません: ' + colId);

  // カラム内カード数チェック
  var maxCards = parseInt(getConfig_('カードの最大数/カラム')) || 20;
  var cards = getCards_(boardId).filter(function(c) { return c.colId === colId; });
  if (cards.length >= maxCards) throw new Error('このカラムのカード数上限（' + maxCards + '）に達しています');

  var order = cards.length + 1;
  var deadlineDt = deadline ? new Date(deadline) : '';
  if (deadline && isNaN(new Date(deadline).getTime())) throw new Error('期限の形式が不正です');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CARDS);
  var id = nextCardId_();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 10).setValues([[
    id, boardId, colId, title, desc, assignee, label, deadlineDt, order, new Date()
  ]]);
  log_('INFO', 'カード追加: ' + title + ' (' + id + ')');
  notifyWebhook_('新カード: ' + title);
  return { success: true, id: id };
}

function updateCard_(cardId, title, desc, assignee, label, deadline) {
  cardId = String(cardId || '').trim();
  if (!cardId) throw new Error('カードIDは必須です');
  title = String(title || '').trim();
  if (!title) throw new Error('タイトルは必須です');

  var deadlineDt = deadline ? new Date(deadline) : '';
  if (deadline && isNaN(new Date(deadline).getTime())) throw new Error('期限の形式が不正です');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CARDS);
  if (!sh || sh.getLastRow() < 2) throw new Error('カードが見つかりません');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === cardId) {
      sh.getRange(i + 2, 4, 1, 5).setValues([[
        title, desc || '', assignee || '', label || '', deadlineDt
      ]]);
      log_('INFO', 'カード更新: ' + title);
      return { success: true };
    }
  }
  throw new Error('カードが見つかりません: ' + cardId);
}

function moveCard_(cardId, newColId) {
  cardId = String(cardId || '').trim();
  if (!cardId) throw new Error('カードIDは必須です');
  newColId = String(newColId || '').trim();
  if (!newColId) throw new Error('移動先カラムIDは必須です');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CARDS);
  if (!sh || sh.getLastRow() < 2) throw new Error('カードが見つかりません');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === cardId) {
      var oldColId = data[i][2];
      if (oldColId === newColId) return { success: true };

      // 移動先カラム存在チェック
      var cols = getColumns_(data[i][1]);
      var found = false, colName = '';
      for (var j = 0; j < cols.length; j++) {
        if (cols[j].id === newColId) { found = true; colName = cols[j].name; break; }
      }
      if (!found) throw new Error('カラムが見つかりません: ' + newColId);

      sh.getRange(i + 2, 3).setValue(newColId);
      log_('INFO', 'カード移動: ' + data[i][3] + ' → ' + colName);
      notifyWebhook_(data[i][3] + ' → ' + colName);
      return { success: true };
    }
  }
  throw new Error('カードが見つかりません: ' + cardId);
}

function deleteCard_(cardId) {
  cardId = String(cardId || '').trim();
  if (!cardId) throw new Error('カードIDは必須です');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_CARDS);
  if (!sh || sh.getLastRow() < 2) throw new Error('カードが見つかりません');
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === cardId) {
      sh.deleteRow(i + 2);
      log_('INFO', 'カード削除: ' + cardId);
      return { success: true };
    }
  }
  throw new Error('カードが見つかりません: ' + cardId);
}

// ══════════════════════════════════════════════════
//  ラベル管理
// ══════════════════════════════════════════════════

function getLabels_(boardId) {
  boardId = String(boardId || '').trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_LABELS);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  return data.filter(function(r) { return r[0] && (!boardId || r[2] === boardId); })
    .map(function(r) { return { name: r[0], color: r[1], boardId: r[2] }; });
}

function addLabel_(boardId, name, color) {
  boardId = String(boardId || '').trim();
  if (!boardId) throw new Error('ボードIDは必須です');
  name = String(name || '').trim();
  if (!name) throw new Error('ラベル名は必須です');
  color = String(color || '#808080').trim();

  var labels = getLabels_(boardId);
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].name === name) throw new Error('同名のラベルが既に存在します');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH_LABELS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[name, color, boardId]]);
  log_('INFO', 'ラベル追加: ' + name);
  return { success: true };
}

// ══════════════════════════════════════════════════
//  ダッシュボード
// ══════════════════════════════════════════════════

function getDashboard_(boardId) {
  boardId = String(boardId || '').trim();
  if (!boardId) {
    var boards = getBoards_();
    if (boards.length === 0) return { columns: [], cards: [], labels: [], boards: [] };
    boardId = boards[0].id;
  }

  var columns = getColumns_(boardId);
  var cards = getCards_(boardId);
  var labels = getLabels_(boardId);
  var boards = getBoards_();

  // カラム別カード整理
  var colCards = {};
  columns.forEach(function(c) { colCards[c.id] = []; });
  cards.forEach(function(c) {
    if (colCards[c.colId]) colCards[c.colId].push(c);
  });

  // 期限切れカードカウント
  var now = new Date();
  var overdue = cards.filter(function(c) {
    return c.deadline && new Date(c.deadline) < now && c.colId !== columns[columns.length - 1].id;
  }).length;

  return {
    boardId: boardId,
    boards: boards,
    columns: columns,
    cards: cards,
    colCards: colCards,
    labels: labels,
    totalCards: cards.length,
    overdueCards: overdue
  };
}
