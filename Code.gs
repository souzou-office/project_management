// ============================================
// 設定
// ============================================
const CONFIG = {
  // API key must be stored in Script Properties, not in source code.
  // Set it via: File > Project settings > Script properties > Add: CLAUDE_API_KEY = your-key
  get CLAUDE_API_KEY() {
    return PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  },
  SPREADSHEET_ID: '14tvW6j9OlgeDslG035CrGA6Cyui2C39hX7byYl-dNyI',
  DRIVE_FOLDER_ID: '1ciQnDa4WOLY6K5eY-EPWc7MtxUoGxP1X',
  PROCESSED_LABEL: '見積依頼処理済',
  SEARCH_QUERY: 'to:info@souzou-office.jp -label:見積依頼処理済 newer_than:3d',
  MASTER_FOLDER_ID: '1ciQnDa4WOLY6K5eY-EPWc7MtxUoGxP1X',
  TRANSFER_WEB_APP_URL: 'https://script.google.com/a/macros/souzou-office.jp/s/AKfycbzjCXA1fPb_j0mdB4rqUrf1cZxG_ne3cFzxE3FVhkSdcqOnN-NjOhI_RdC_D2k_fFiLQg/exec',
  STAFF_LIST: ['古江', '菖蒲', '石橋', '江口', '中村', '池田', '浦'],
};

// ============================================
// カラム定義
// ============================================
// メイン:      会社名(1) 案件要約(2) 件名(3) 日時(4) 資料(5) メール(6) 担当者(7) URL(8) チェック(9) 処理済(10) 除外(11) threadId(12)
// 担当者タブ:  会社名(1) 案件要約(2) 件名(3) 日時(4) 資料(5) メール(6) 担当者(7) URL(8) チェック(9) 処理済(10) 戻す(11) threadId(12)
// 処理済タブ:  会社名(1) 案件要約(2) 件名(3) 日時(4) 資料(5) メール(6) 担当者(7) URL(8) チェック(9) 戻す(10) threadId(11)
// 除外タブ:    会社名(1) 案件要約(2) 件名(3) 日時(4) 資料(5) メール(6) 担当者(7) URL(8) チェック(9) 戻す(10) threadId(11)

// ============================================
// メイン処理：見積依頼メールを処理
// ============================================
function processQuoteEmails() {
  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, 200);
  Logger.log('ヒット件数: ' + threads.length);
  if (threads.length === 0) return;

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('メイン') || ss.getSheets()[0];
  const label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['会社名', '案件要約', '件名', '日時', '資料', 'メール', '担当者', 'URL', 'チェック', '処理済', '除外', 'threadId']);
    sheet.hideColumns(12);
    formatHeader(sheet, 12);
  }

  // 全シートから既存threadIdを収集
  let existingThreadIds = new Set();
  const allSheetNames = ['メイン', '処理済', '除外'].concat(CONFIG.STAFF_LIST);
  allSheetNames.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const lastRow = s.getLastRow();
    if (lastRow <= 1) return;
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const tidCol = headers.indexOf('threadId') + 1;
    if (tidCol === 0) return;
    const ids = s.getRange(2, tidCol, lastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingThreadIds.add(r[0]); });
  });
  Logger.log('既存ID数: ' + existingThreadIds.size);

  for (const thread of threads) {
    const threadId = thread.getId();
    Logger.log('チェック中: ' + threadId);

    if (existingThreadIds.has(threadId)) {
      Logger.log('スキップ（既存）: ' + threadId);
      thread.addLabel(label);
      continue;
    }

    const messages = thread.getMessages();
    const firstMessage = messages[0];

    const daysAgo = new Date();
    daysAgo.setDate(daysAgo.getDate() - 7);
    if (firstMessage.getDate() < daysAgo) {
      Logger.log('スキップ（古い）: ' + firstMessage.getDate());
      thread.addLabel(label);
      continue;
    }

    const subject = firstMessage.getSubject();
    const from = firstMessage.getFrom();
    if (from.includes('@souzou-office.jp')) {
      thread.addLabel(label);
      continue;
    }

    const date = firstMessage.getDate();
    const body = messages.map(m => m.getPlainBody()).join('\n---\n');

    const analysis = analyzeWithClaude(subject, body, from);
    Logger.log('件名: ' + subject);
    Logger.log('判定結果: ' + JSON.stringify(analysis));

    if (!analysis.is_quote_request) {
      thread.addLabel(label);
      continue;
    }

    Logger.log('新規追加: ' + threadId);

    const attachments = messages.flatMap(m => m.getAttachments());
    let folderLink = '';
    if (attachments.length > 0) {
      const folderName = subject.substring(0, 50).replace(/[\/\\?%*:|"<>]/g, '') + '_' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
      const subFolder = folder.createFolder(folderName);
      attachments.forEach(att => subFolder.createFile(att));
      folderLink = '=HYPERLINK("' + subFolder.getUrl() + '","📂開く")';
    }

    const rfc822Id = firstMessage.getHeader('Message-ID');
    const mailLink = '=HYPERLINK("https://mail.google.com/mail/u/0/#search/rfc822msgid:' + encodeURIComponent(rfc822Id) + '","📧開く")';

    const shortDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yy/M/d H:mm');
    const shortSubject = subject.length > 30 ? subject.substring(0, 30) + '...' : subject;

    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 12).setValues([[
      analysis.company_name || '',
      analysis.summary || '',
      shortSubject,
      shortDate,
      '',
      '',
      '',
      '',
      false,
      false,
      false,
      threadId
    ]]);
    if (folderLink) sheet.getRange(2, 5).setFormula(folderLink);
    sheet.getRange(2, 6).setFormula(mailLink);
    sheet.getRange(2, 9).insertCheckboxes();
    sheet.getRange(2, 10).insertCheckboxes();
    sheet.getRange(2, 11).insertCheckboxes();
    const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STAFF_LIST, true).setAllowInvalid(false).build();
    sheet.getRange(2, 7).setDataValidation(staffRule);

    thread.addLabel(label);
    existingThreadIds.add(threadId);
  }

  sortByDate(sheet);
  clearDataFormat(sheet);
}

// ============================================
// Claude APIで見積依頼か判定
// ============================================
function analyzeWithClaude(subject, body, from) {
  const safeBody = body || '';
  const safeSubject = subject || '';
  const safeFrom = from || '';

  const prompt = 'あなたは司法書士事務所の受付担当です。以下のメールが「見積依頼」かどうか判定してください。\n\n【判定基準】\n- 登記、相続、会社設立などの見積もり・費用を知りたい → true\n- 売買の決済で司法書士を依頼したい → true\n- 「見積」「お見積り」「費用」「ご依頼」という言葉がある → 基本的にtrue\n- 広告・営業・セミナー案内メール → false\n- こちらから見積もりを送った返信で、単なる確認・お礼 → false\n- FAXや添付ファイルの中身が不明で判断できない → false\n\n【重要】迷ったらtrueにしてください。取りこぼし防止が優先。\n\n【メール情報】\n送信者: ' + safeFrom + '\n件名: ' + safeSubject + '\n本文:\n' + safeBody.substring(0, 3000) + '\n\n【出力形式】以下のJSONのみを出力:\n{"is_quote_request": true/false, "company_name": "送信者の会社名", "contact_person": "担当者名", "summary": "何の依頼か100字以内で要約"}';

  const apiKey = CONFIG.CLAUDE_API_KEY;
  if (!apiKey) {
    Logger.log('Error: CLAUDE_API_KEY is not set in Script Properties');
    return {is_quote_request: false, company_name: null, contact_person: null, summary: null};
  }

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 500,
      messages: [{role: 'user', content: prompt}]
    }),
    muteHttpExceptions: true
  });

  try {
    const result = JSON.parse(response.getContentText());
    const text = result.content[0].text;
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {
    Logger.log('Claude parse error: ' + e);
  }

  return {is_quote_request: false, company_name: null, contact_person: null, summary: null};
}

// ============================================
// チェックボックス監視（編集時に即座に移動）
// ============================================
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const value = e.value;

  if (row === 1) return;
  if (value !== 'TRUE') return;

  const ss = e.source;

  // --- メインタブ ---
  if (sheetName === 'メイン') {
    if (col === 10) {
      const staff = sheet.getRange(row, 7).getValue();
      if (!staff) {
        sheet.getRange(row, 10).setValue(false);
        return;
      }
      const staffSheet = getOrCreateStaffSheet(ss, staff);
      moveRow(sheet, row, staffSheet, 'main_to_staff');
    } else if (col === 11) {
      const excludeSheet = ss.getSheetByName('除外');
      moveRow(sheet, row, excludeSheet, 'main_to_exclude');
    }
  }
  // --- 担当者タブ ---
  else if (CONFIG.STAFF_LIST.includes(sheetName)) {
    if (col === 10) {
      const doneSheet = ss.getSheetByName('処理済');
      moveRow(sheet, row, doneSheet, 'staff_to_done');
    } else if (col === 11) {
      const mainSheet = ss.getSheetByName('メイン');
      moveRow(sheet, row, mainSheet, 'staff_to_main');
    }
  }
  // --- 処理済タブ ---
  else if (sheetName === '処理済') {
    if (col === 10) {
      const staff = sheet.getRange(row, 7).getValue();
      if (staff && CONFIG.STAFF_LIST.includes(staff)) {
        const staffSheet = getOrCreateStaffSheet(ss, staff);
        moveRow(sheet, row, staffSheet, 'done_to_staff');
      } else {
        const mainSheet = ss.getSheetByName('メイン');
        moveRow(sheet, row, mainSheet, 'done_to_main');
      }
    }
  }
  // --- 除外タブ ---
  else if (sheetName === '除外') {
    if (col === 10) {
      const mainSheet = ss.getSheetByName('メイン');
      moveRow(sheet, row, mainSheet, 'exclude_to_main');
    }
  }
  // --- 手動追加タブ ---
  else if (sheetName === '手動追加') {
    if (col === 1 && e.value) {
      processManualAddFromUrl(ss, sheet, row);
    }
  }
}
// ============================================
// 行移動（汎用）
// ============================================
function moveRow(fromSheet, row, toSheet, direction) {
  const numCols = fromSheet.getLastColumn();
  const data = fromSheet.getRange(row, 1, 1, numCols).getDisplayValues()[0];
  const formulas = fromSheet.getRange(row, 5, 1, 2).getFormulas()[0];

  const fromHeaders = fromSheet.getRange(1, 1, 1, numCols).getValues()[0];
  const fromTidCol = fromHeaders.indexOf('threadId');
  const threadId = (fromTidCol >= 0) ? fromSheet.getRange(row, fromTidCol + 1).getValue() : '';

  const urlValue = data[7] || '';

  if (direction === 'main_to_staff') {
    toSheet.appendRow([data[0], data[1], data[2], data[3], '', '', data[6], urlValue, false, false, false, threadId]);
    const newRow = toSheet.getLastRow();
    if (formulas[0]) toSheet.getRange(newRow, 5).setFormula(formulas[0]);
    if (formulas[1]) toSheet.getRange(newRow, 6).setFormula(formulas[1]);
    toSheet.getRange(newRow, 9).insertCheckboxes();
    toSheet.getRange(newRow, 10).insertCheckboxes();
    toSheet.getRange(newRow, 11).insertCheckboxes();
  }
  else if (direction === 'main_to_exclude') {
    toSheet.appendRow([data[0], data[1], data[2], data[3], '', '', data[6], urlValue, false, false, threadId]);
    const newRow = toSheet.getLastRow();
    if (formulas[0]) toSheet.getRange(newRow, 5).setFormula(formulas[0]);
    if (formulas[1]) toSheet.getRange(newRow, 6).setFormula(formulas[1]);
    toSheet.getRange(newRow, 9).insertCheckboxes();
    toSheet.getRange(newRow, 10).insertCheckboxes();
  }
  else if (direction === 'staff_to_done') {
    toSheet.appendRow([data[0], data[1], data[2], data[3], '', '', data[6], urlValue, false, false, threadId]);
    const newRow = toSheet.getLastRow();
    if (formulas[0]) toSheet.getRange(newRow, 5).setFormula(formulas[0]);
    if (formulas[1]) toSheet.getRange(newRow, 6).setFormula(formulas[1]);
    toSheet.getRange(newRow, 9).insertCheckboxes();
    toSheet.getRange(newRow, 10).insertCheckboxes();
  }
  else if (direction === 'done_to_staff') {
    toSheet.appendRow([data[0], data[1], data[2], data[3], '', '', data[6], urlValue, false, false, false, threadId]);
    const newRow = toSheet.getLastRow();
    if (formulas[0]) toSheet.getRange(newRow, 5).setFormula(formulas[0]);
    if (formulas[1]) toSheet.getRange(newRow, 6).setFormula(formulas[1]);
    toSheet.getRange(newRow, 9).insertCheckboxes();
    toSheet.getRange(newRow, 10).insertCheckboxes();
    toSheet.getRange(newRow, 11).insertCheckboxes();
  }
  else if (direction === 'staff_to_main' || direction === 'done_to_main' || direction === 'exclude_to_main') {
    toSheet.appendRow([data[0], data[1], data[2], data[3], '', '', data[6], urlValue, false, false, false, threadId]);
    const newRow = toSheet.getLastRow();
    if (formulas[0]) toSheet.getRange(newRow, 5).setFormula(formulas[0]);
    if (formulas[1]) toSheet.getRange(newRow, 6).setFormula(formulas[1]);
    toSheet.getRange(newRow, 9).insertCheckboxes();
    toSheet.getRange(newRow, 10).insertCheckboxes();
    toSheet.getRange(newRow, 11).insertCheckboxes();
    const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STAFF_LIST, true).setAllowInvalid(false).build();
    toSheet.getRange(newRow, 7).setDataValidation(staffRule);
  }

  fromSheet.deleteRow(row);

  if (direction === 'staff_to_main' || direction === 'done_to_main' || direction === 'exclude_to_main') {
    sortByDate(toSheet);
  }
}

// ============================================
// 担当者タブの取得または作成
// ============================================
function getOrCreateStaffSheet(ss, staffName) {
  let sheet = ss.getSheetByName(staffName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(staffName);
  sheet.appendRow(['会社名', '案件要約', '件名', '日時', '資料', 'メール', '担当者', 'URL', 'チェック', '処理済', '戻す', 'threadId']);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 600);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 200);
  sheet.setColumnWidth(9, 60);
  sheet.setColumnWidth(10, 60);
  sheet.setColumnWidth(11, 60);
  sheet.hideColumns(12);
  formatHeader(sheet, 12);

  return sheet;
}

// ============================================
// ラベル取得または作成
// ============================================
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// ============================================
// ヘッダー書式設定
// ============================================
function formatHeader(sheet, cols) {
  const header = sheet.getRange(1, 1, 1, cols);
  header.setBackground('#1a73e8');
  header.setFontColor('white');
  header.setFontWeight('bold');
  header.setHorizontalAlignment('center');
}

// ============================================
// データ部分の書式クリア
// ============================================
function clearDataFormat(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    dataRange.setBackground('white');
    dataRange.setFontColor('black');
    dataRange.setFontWeight('normal');
  }
}

// ============================================
// 初期セットアップ（v4）
// ============================================
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const staffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STAFF_LIST, true)
    .setAllowInvalid(false)
    .build();

  let mainSheet = ss.getSheetByName('メイン');
  if (!mainSheet) {
    mainSheet = ss.getSheets()[0];
    mainSheet.setName('メイン');
  }
  mainSheet.clear();
  mainSheet.appendRow(['会社名', '案件要約', '件名', '日時', '資料', 'メール', '担当者', 'URL', 'チェック', '処理済', '除外', 'threadId']);
  mainSheet.setFrozenRows(1);
  mainSheet.setColumnWidth(1, 150);
  mainSheet.setColumnWidth(2, 600);
  mainSheet.setColumnWidth(3, 180);
  mainSheet.setColumnWidth(4, 90);
  mainSheet.setColumnWidth(5, 60);
  mainSheet.setColumnWidth(6, 60);
  mainSheet.setColumnWidth(7, 80);
  mainSheet.setColumnWidth(8, 200);
  mainSheet.setColumnWidth(9, 60);
  mainSheet.setColumnWidth(10, 60);
  mainSheet.setColumnWidth(11, 60);
  mainSheet.hideColumns(12);
  formatHeader(mainSheet, 12);
  mainSheet.getRange('A:L').setVerticalAlignment('middle');
  mainSheet.getRange('A:L').setWrap(true);
  mainSheet.getRange('G2:G1000').setDataValidation(staffRule);

  let doneSheet = ss.getSheetByName('処理済');
  if (!doneSheet) {
    doneSheet = ss.insertSheet('処理済');
  }
  doneSheet.clear();
  doneSheet.appendRow(['会社名', '案件要約', '件名', '日時', '資料', 'メール', '担当者', 'URL', 'チェック', '戻す', 'threadId']);
  doneSheet.setFrozenRows(1);
  doneSheet.setColumnWidth(1, 150);
  doneSheet.setColumnWidth(2, 600);
  doneSheet.setColumnWidth(3, 180);
  doneSheet.setColumnWidth(4, 90);
  doneSheet.setColumnWidth(5, 60);
  doneSheet.setColumnWidth(6, 60);
  doneSheet.setColumnWidth(7, 80);
  doneSheet.setColumnWidth(8, 200);
  doneSheet.setColumnWidth(9, 60);
  doneSheet.setColumnWidth(10, 60);
  doneSheet.hideColumns(11);
  formatHeader(doneSheet, 11);

  let excludeSheet = ss.getSheetByName('除外');
  if (!excludeSheet) {
    excludeSheet = ss.insertSheet('除外');
  }
  excludeSheet.clear();
  excludeSheet.appendRow(['会社名', '案件要約', '件名', '日時', '資料', 'メール', '担当者', 'URL', 'チェック', '戻す', 'threadId']);
  excludeSheet.setFrozenRows(1);
  excludeSheet.setColumnWidth(1, 150);
  excludeSheet.setColumnWidth(2, 600);
  excludeSheet.setColumnWidth(3, 180);
  excludeSheet.setColumnWidth(4, 90);
  excludeSheet.setColumnWidth(5, 60);
  excludeSheet.setColumnWidth(6, 60);
  excludeSheet.setColumnWidth(7, 80);
  excludeSheet.setColumnWidth(8, 200);
  excludeSheet.setColumnWidth(9, 60);
  excludeSheet.setColumnWidth(10, 60);
  excludeSheet.hideColumns(11);
  formatHeader(excludeSheet, 11);

  CONFIG.STAFF_LIST.forEach(name => {
    getOrCreateStaffSheet(ss, name);
  });
}

// ============================================
// 日付でソート（新しい順）
// ============================================
function sortByDate(sheet) {
  if (!sheet) {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    sheet = ss.getSheetByName('メイン');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({column: 4, ascending: false});
  }
}

// ============================================
// 書式修正
// ============================================
function fixFormat() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const allNames = ['メイン', '処理済', '除外'].concat(CONFIG.STAFF_LIST);
  allNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) clearDataFormat(sheet);
  });
}

// ============================================
// 空行削除
// ============================================
function deleteEmptyRows() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const allNames = ['メイン', '処理済', '除外'].concat(CONFIG.STAFF_LIST);
  allNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === '' && data[i][1] === '') {
        sheet.deleteRow(i + 1);
      }
    }
  });
}

// ============================================
// テスト・ユーティリティ
// ============================================
function testDropdown() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('メイン');
  const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STAFF_LIST, true).setAllowInvalid(false).build();
  sheet.getRange('G2:G100').setDataValidation(staffRule);
}

function testSearch() {
  const threads = GmailApp.search('to:info@souzou-office.jp -label:見積依頼処理済 newer_than:3d', 0, 50);
  Logger.log('ヒット件数: ' + threads.length);
  for (let i = 0; i < Math.min(threads.length, 10); i++) {
    Logger.log(threads[i].getFirstMessageSubject());
  }
}

function fixAllRows() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('メイン');
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const staffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STAFF_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, lastRow - 1, 1).setDataValidation(staffRule);

  sheet.getRange(2, 9, lastRow - 1, 1).insertCheckboxes();
  sheet.getRange(2, 10, lastRow - 1, 1).insertCheckboxes();
  sheet.getRange(2, 11, lastRow - 1, 1).insertCheckboxes();
}

// ============================================
// 手動追加：件名で検索してメインに追加
// ============================================
function processManualAdd() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const mainSheet = ss.getSheetByName('メイン');
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);

  // 既存threadId収集
  let existingThreadIds = new Set();
  const allSheetNames = ['メイン', '処理済', '除外'].concat(CONFIG.STAFF_LIST);
  allSheetNames.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const lastRow = s.getLastRow();
    if (lastRow <= 1) return;
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const tidCol = headers.indexOf('threadId') + 1;
    if (tidCol === 0) return;
    const ids = s.getRange(2, tidCol, lastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingThreadIds.add(r[0]); });
  });

  // --- 手動追加シートから処理 ---
  const addSheet = ss.getSheetByName('手動追加');
  if (addSheet && addSheet.getLastRow() > 1) {
    const subjects = addSheet.getRange(2, 1, addSheet.getLastRow() - 1, 1).getValues();
    for (let i = subjects.length - 1; i >= 0; i--) {
      const subject = subjects[i][0];
      if (!subject) continue;

      try {
        const threads = GmailApp.search('to:info@souzou-office.jp subject:' + subject, 0, 5);
        let added = false;

        for (const thread of threads) {
          const threadId = thread.getId();
          if (existingThreadIds.has(threadId)) continue;

          addToMain(thread, mainSheet, folder, processedLabel, existingThreadIds);
          added = true;
          break;
        }
      } catch (e) {
        Logger.log('手動追加エラー: ' + e);
      }

      addSheet.deleteRow(i + 2);
    }
  }

  sortByDate(mainSheet);
  clearDataFormat(mainSheet);
}

// ============================================
// メインシートに1件追加（共通処理）
// ============================================
function addToMain(thread, mainSheet, folder, processedLabel, existingThreadIds) {
  const messages = thread.getMessages();
  const firstMessage = messages[0];
  const subject = firstMessage.getSubject();
  const from = firstMessage.getFrom();
  const date = firstMessage.getDate();
  const body = messages.map(m => m.getPlainBody()).join('\n---\n');
  const threadId = thread.getId();

  const analysis = analyzeWithClaude(subject, body, from);

  const attachments = messages.flatMap(m => m.getAttachments());
  let folderLink = '';
  if (attachments.length > 0) {
    const folderName = subject.substring(0, 50).replace(/[\/\\?%*:|"<>]/g, '') + '_' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
    const subFolder = folder.createFolder(folderName);
    attachments.forEach(att => subFolder.createFile(att));
    folderLink = '=HYPERLINK("' + subFolder.getUrl() + '","📂開く")';
  }

  const rfc822Id = firstMessage.getHeader('Message-ID');
  const mailLink = '=HYPERLINK("https://mail.google.com/mail/u/0/#search/rfc822msgid:' + encodeURIComponent(rfc822Id) + '","📧開く")';

  const shortDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yy/M/d H:mm');
  const shortSubject = subject.length > 30 ? subject.substring(0, 30) + '...' : subject;

  mainSheet.insertRowAfter(1);
  mainSheet.getRange(2, 1, 1, 12).setValues([[
    analysis.company_name || '',
    analysis.summary || '',
    shortSubject,
    shortDate,
    '',
    '',
    '',
    '',
    false,
    false,
    false,
    threadId
  ]]);
  if (folderLink) mainSheet.getRange(2, 5).setFormula(folderLink);
  mainSheet.getRange(2, 6).setFormula(mailLink);
  mainSheet.getRange(2, 9).insertCheckboxes();
  mainSheet.getRange(2, 10).insertCheckboxes();
  mainSheet.getRange(2, 11).insertCheckboxes();
  const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STAFF_LIST, true).setAllowInvalid(false).build();
  mainSheet.getRange(2, 7).setDataValidation(staffRule);

  thread.addLabel(processedLabel);
  existingThreadIds.add(threadId);
}
// ============================================
// 手動追加：URLからメインに追加
// ============================================
function processManualAddFromUrl(ss, addSheet, row) {
  const url = addSheet.getRange(row, 1).getValue();
  if (!url) return;

  // URLからスレッドIDを抽出
  const match = url.match(/\/([A-Za-z0-9_-]+)\s*$/);
  if (!match) return;
  const threadId = match[1];

  // 既存チェック
  let exists = false;
  const allSheetNames = ['メイン', '処理済', '除外'].concat(CONFIG.STAFF_LIST);
  allSheetNames.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) return;
    const lastRow = s.getLastRow();
    if (lastRow <= 1) return;
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    const tidCol = headers.indexOf('threadId') + 1;
    if (tidCol === 0) return;
    const ids = s.getRange(2, tidCol, lastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0] == threadId) exists = true; });
  });

  if (exists) {
    addSheet.deleteRow(row);
    return;
  }

  // メールを取得
  try {
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      addSheet.deleteRow(row);
      return;
    }

    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const subject = firstMessage.getSubject();
    const from = firstMessage.getFrom();
    const date = firstMessage.getDate();
    const body = messages.map(m => m.getPlainBody()).join('\n---\n');

    const analysis = analyzeWithClaude(subject, body, from);

    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const attachments = messages.flatMap(m => m.getAttachments());
    let folderLink = '';
    if (attachments.length > 0) {
      const folderName = subject.substring(0, 50).replace(/[\/\\?%*:|"<>]/g, '') + '_' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMdd');
      const subFolder = folder.createFolder(folderName);
      attachments.forEach(att => subFolder.createFile(att));
      folderLink = '=HYPERLINK("' + subFolder.getUrl() + '","📂開く")';
    }

    const rfc822Id = firstMessage.getHeader('Message-ID');
    const mailLink = '=HYPERLINK("https://mail.google.com/mail/u/0/#search/rfc822msgid:' + encodeURIComponent(rfc822Id) + '","📧開く")';

    const shortDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yy/M/d H:mm');
    const shortSubject = subject.length > 30 ? subject.substring(0, 30) + '...' : subject;

    const mainSheet = ss.getSheetByName('メイン');
    mainSheet.insertRowAfter(1);
    mainSheet.getRange(2, 1, 1, 12).setValues([[
      analysis.company_name || '',
      analysis.summary || '',
      shortSubject,
      shortDate,
      '',
      '',
      '',
      '',
      false,
      false,
      false,
      threadId
    ]]);
    if (folderLink) mainSheet.getRange(2, 5).setFormula(folderLink);
    mainSheet.getRange(2, 6).setFormula(mailLink);
    mainSheet.getRange(2, 9).insertCheckboxes();
    mainSheet.getRange(2, 10).insertCheckboxes();
    mainSheet.getRange(2, 11).insertCheckboxes();
    const staffRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STAFF_LIST, true).setAllowInvalid(false).build();
    mainSheet.getRange(2, 7).setDataValidation(staffRule);

    // 処理済ラベル追加
    const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
    thread.addLabel(processedLabel);

  } catch (e) {
    Logger.log('手動追加エラー: ' + e);
  }

  addSheet.deleteRow(row);
}
function createManualAddSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('手動追加');
  if (sheet) return;
  sheet = ss.insertSheet('手動追加');
  sheet.appendRow(['件名（コピペ）']);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 600);
  formatHeader(sheet, 1);
}
function testSubjectSearch() {
  const subject = '大野城市若草二丁目土地の件です';
  const threads = GmailApp.search('to:info@souzou-office.jp subject:' + subject, 0, 5);
  Logger.log('ヒット数: ' + threads.length);
  for (const t of threads) {
    Logger.log('件名: ' + t.getFirstMessageSubject());
    Logger.log('ID: ' + t.getId());
  }
}
