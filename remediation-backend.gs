/**
 * Remedial System - Google Apps Script 後台
 * v3.1 Backend — 儲存批改記錄、考卷檢核結果至 Google 試算表，
 *                考卷照片存至 Google Drive（試算表僅存連結）
 *
 * 使用方式：
 * 1. 開啟 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 4. 複製網址，貼入系統「設定」→「Google Apps Script 網址」
 *
 * ★ 從 v3.0 升級：貼上新程式碼後，必須「部署 → 管理部署作業 →
 *   編輯 → 版本：新版本」重新部署，新功能才會生效（網址不變）。
 *   首次儲存照片時 Apps Script 會要求授權 Google Drive 權限。
 */

// ── 試算表設定 ──────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SUBMISSIONS_SHEET = '批改記錄';
const EXAMS_SHEET = '考卷設定';
const VALIDATIONS_SHEET = '考卷檢核';
const PHOTO_FOLDER = 'Remedial System 考卷照片';

// ★ 欄位定義 — 與 HTML submission 物件對齊
const SUB_COLS = [
  'id', 'examId', 'studentName', 'seatNo',
  'score', 'totalQuestions', 'percentage',
  'misconceptions', 'wrongQuestions', 'teacherNotes',
  'gradedAt', 'fileName', 'pageCount',
  'errorDetails', 'photoUrls'
];

// ★ 欄位定義 — 與 HTML exam 物件對齊
const EXAM_COLS = [
  'id', 'name', 'subject', 'grade',
  'year', 'sem', 'createdAt', 'roster',
  'answerKeyUrl'
];

// ★ 欄位定義 — 空白考卷檢核結果
const VAL_COLS = [
  'id', 'checkedAt', 'grade', 'subject', 'examTitle',
  'summary', 'issueCount',
  'contentIssues', 'duplicates', 'inappropriate',
  'difficulty', 'teacherAdvice',
  'photoUrl', 'fileName'
];

// ── 初始化試算表 ────────────────────────────────
function ensureSheet(name, cols) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < cols.length) {
    // 舊版工作表升級：補上新增欄位的標題
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  }
  return sh;
}
function setupSheets() {
  return {
    subSheet: ensureSheet(SUBMISSIONS_SHEET, SUB_COLS),
    examSheet: ensureSheet(EXAMS_SHEET, EXAM_COLS),
    valSheet: ensureSheet(VALIDATIONS_SHEET, VAL_COLS),
  };
}

// ── Google Drive 照片儲存 ───────────────────────
function getPhotoFolder() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function saveImages(images, prefix) {
  if (!images || !images.length) return [];
  const folder = getPhotoFolder();
  return images.map((img, i) => {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(img.b64), 'image/jpeg',
        prefix + '_' + (i + 1) + '_' + (img.name || 'photo.jpg'));
      return folder.createFile(blob).getUrl();
    } catch (err) {
      return '儲存失敗: ' + err;
    }
  });
}
// 讀取既有列的指定欄位（避免同步時把照片連結洗掉）
function fieldMap(sheet, cols, fields) {
  const data = sheet.getDataRange().getValues();
  const m = {};
  data.slice(1).forEach(r => {
    const o = {};
    fields.forEach(f => { o[f] = r[cols.indexOf(f)] || ''; });
    m[String(r[0])] = o;
  });
  return m;
}

// ── GET 處理 ────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'all';

  if (action === 'ping') {
    return jsonResp({ ok: true, message: 'Remedial System 後台連線正常', time: new Date().toISOString() });
  }

  if (action === 'all') {
    const { subSheet, examSheet, valSheet } = setupSheets();
    const submissions = sheetToObjects(subSheet, SUB_COLS).map(parseSub);
    const exams = sheetToObjects(examSheet, EXAM_COLS);
    const validations = sheetToObjects(valSheet, VAL_COLS).map(parseVal);
    exams.forEach(ex => {
      try { ex.roster = JSON.parse(ex.roster || '[]'); } catch { ex.roster = []; }
    });
    return jsonResp({ ok: true, submissions, exams, validations });
  }

  if (action === 'submissions') {
    const { subSheet } = setupSheets();
    const submissions = sheetToObjects(subSheet, SUB_COLS).map(parseSub);
    return jsonResp({ ok: true, submissions });
  }

  if (action === 'validations') {
    const { valSheet } = setupSheets();
    const validations = sheetToObjects(valSheet, VAL_COLS).map(parseVal);
    return jsonResp({ ok: true, validations });
  }

  return jsonResp({ ok: false, error: 'Unknown action' });
}

function parseSub(s) {
  try { s.misconceptions = JSON.parse(s.misconceptions || '[]'); } catch { s.misconceptions = []; }
  try { s.wrongQuestions = JSON.parse(s.wrongQuestions || '[]'); } catch { s.wrongQuestions = []; }
  try { s.errorDetails = JSON.parse(s.errorDetails || '[]'); } catch { s.errorDetails = []; }
  return s;
}
function parseVal(v) {
  ['contentIssues', 'duplicates', 'inappropriate', 'teacherAdvice'].forEach(k => {
    try { v[k] = JSON.parse(v[k] || '[]'); } catch { v[k] = []; }
  });
  try { v.difficulty = JSON.parse(v.difficulty || 'null'); } catch { v.difficulty = null; }
  return v;
}

// ── POST 處理 ───────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch { return jsonResp({ ok: false, error: 'Invalid JSON' }); }

  const action = body.action;

  if (action === 'saveSubmission') {
    const sub = body.submission;
    if (!sub?.id) return jsonResp({ ok: false, error: '缺少 submission.id' });
    const { subSheet } = setupSheets();
    // 學生考卷照片存 Drive，試算表存連結；未附照片時保留既有連結
    const urls = saveImages(body.images, '學生考卷_' + sub.id);
    if (urls.length) sub.photoUrls = urls.join('\n');
    else {
      const keep = fieldMap(subSheet, SUB_COLS, ['photoUrls'])[String(sub.id)];
      if (keep) sub.photoUrls = keep.photoUrls;
    }
    upsertRow(subSheet, SUB_COLS, sub.id, subToRow(sub));
    return jsonResp({ ok: true, message: '批改記錄已儲存', photoUrls: urls });
  }

  if (action === 'saveExam') {
    const exam = body.exam;
    if (!exam?.id) return jsonResp({ ok: false, error: '缺少 exam.id' });
    const { examSheet } = setupSheets();
    // 答案卷照片存 Drive，試算表存連結
    if (exam.answerKeyB64) {
      const urls = saveImages([{ name: '答案卷.jpg', b64: exam.answerKeyB64 }], '答案卷_' + exam.id);
      exam.answerKeyUrl = urls[0] || '';
    }
    upsertRow(examSheet, EXAM_COLS, exam.id, examToRow(exam));
    return jsonResp({ ok: true, message: '考卷設定已儲存' });
  }

  if (action === 'saveValidation') {
    const v = body.validation;
    if (!v?.id) return jsonResp({ ok: false, error: '缺少 validation.id' });
    const { valSheet } = setupSheets();
    const urls = saveImages(body.images, '空白考卷_' + v.id);
    upsertRow(valSheet, VAL_COLS, v.id, valToRow(v, urls.join('\n')));
    return jsonResp({ ok: true, message: '檢核結果已儲存', photoUrls: urls });
  }

  if (action === 'syncAll') {
    const subs = body.submissions || [];
    const exams = body.exams || [];
    const { subSheet, examSheet } = setupSheets();
    // 同步不帶照片，保留試算表中既有的 Drive 連結
    const subKeep = fieldMap(subSheet, SUB_COLS, ['photoUrls']);
    const examKeep = fieldMap(examSheet, EXAM_COLS, ['answerKeyUrl']);
    subs.forEach(s => {
      if (!s.photoUrls && subKeep[String(s.id)]) s.photoUrls = subKeep[String(s.id)].photoUrls;
      upsertRow(subSheet, SUB_COLS, s.id, subToRow(s));
    });
    exams.forEach(ex => {
      if (!ex.answerKeyUrl && examKeep[String(ex.id)]) ex.answerKeyUrl = examKeep[String(ex.id)].answerKeyUrl;
      upsertRow(examSheet, EXAM_COLS, ex.id, examToRow(ex));
    });
    return jsonResp({ ok: true, message: `同步完成：${subs.length} 筆記錄，${exams.length} 份考卷` });
  }

  if (action === 'deleteSubmission') {
    const id = body.id;
    const { subSheet } = setupSheets();
    deleteRow(subSheet, id);
    return jsonResp({ ok: true, message: '記錄已刪除' });
  }

  return jsonResp({ ok: false, error: `Unknown action: ${action}` });
}

// ── 工具函式 ────────────────────────────────────

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(sheet, cols) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => {
    const obj = {};
    cols.forEach((col, i) => { obj[col] = row[i] ?? ''; });
    return obj;
  }).filter(o => o.id);
}

function upsertRow(sheet, cols, id, rowData) {
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { targetRow = i + 1; break; }
  }
  const row = cols.map(c => rowData[c] !== undefined ? rowData[c] : '');
  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, cols.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function deleteRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; }
  }
}

// ★ 對應 HTML submission 物件欄位
function subToRow(s) {
  return {
    id:             s.id || '',
    examId:         s.examId || '',
    studentName:    s.studentName || '',
    seatNo:         s.seatNo || '',
    score:          s.score ?? '',
    totalQuestions: s.totalQuestions ?? '',
    percentage:     s.percentage ?? '',
    misconceptions: JSON.stringify(s.misconceptions || []),
    wrongQuestions: JSON.stringify(s.wrongQuestions || []),
    teacherNotes:   s.teacherNotes || '',
    gradedAt:       s.gradedAt || new Date().toLocaleDateString('zh-TW'),
    fileName:       s.fileName || '',
    pageCount:      s.pageCount ?? 1,
    errorDetails:   JSON.stringify(s.errorDetails || []),
    photoUrls:      s.photoUrls || '',
  };
}

// ★ 對應 HTML exam 物件欄位
function examToRow(ex) {
  return {
    id:           ex.id || '',
    name:         ex.name || '',
    subject:      ex.subject || '',
    grade:        ex.grade || '',
    year:         ex.year || '',
    sem:          ex.sem || '',
    createdAt:    ex.createdAt || new Date().toLocaleDateString('zh-TW'),
    roster:       JSON.stringify(ex.roster || []),
    answerKeyUrl: ex.answerKeyUrl || '',
  };
}

// ★ 空白考卷檢核結果欄位
function valToRow(v, photoUrl) {
  return {
    id:            v.id || '',
    checkedAt:     v.checkedAt || new Date().toLocaleString('zh-TW'),
    grade:         v.grade || '',
    subject:       v.subject || '',
    examTitle:     v.examTitle || '',
    summary:       v.summary || '',
    issueCount:    v.issueCount ?? 0,
    contentIssues: JSON.stringify(v.contentIssues || []),
    duplicates:    JSON.stringify(v.duplicates || []),
    inappropriate: JSON.stringify(v.inappropriate || []),
    difficulty:    JSON.stringify(v.difficulty || null),
    teacherAdvice: JSON.stringify(v.teacherAdvice || []),
    photoUrl:      photoUrl || '',
    fileName:      v.fileName || '',
  };
}
