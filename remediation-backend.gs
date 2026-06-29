/**
 * Remedial System - Google Apps Script 後台
 * v3.0 Backend — 儲存批改記錄至 Google 試算表
 *
 * 使用方式：
 * 1. 開啟 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 4. 複製網址，貼入系統「設定」→「Google Apps Script 網址」
 */

// ── 試算表設定 ──────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SUBMISSIONS_SHEET = '批改記錄';
const EXAMS_SHEET = '考卷設定';

// ★ 欄位定義 — 與 HTML submission 物件對齊
const SUB_COLS = [
  'id', 'examId', 'studentName', 'seatNo',
  'score', 'totalQuestions', 'percentage',
  'misconceptions', 'wrongQuestions', 'teacherNotes',
  'gradedAt', 'fileName', 'pageCount'
];

// ★ 欄位定義 — 與 HTML exam 物件對齊
const EXAM_COLS = [
  'id', 'name', 'subject', 'grade',
  'year', 'sem', 'createdAt', 'roster'
];

// ── 初始化試算表 ────────────────────────────────
function setupSheets() {
  let subSheet = SS.getSheetByName(SUBMISSIONS_SHEET);
  if (!subSheet) {
    subSheet = SS.insertSheet(SUBMISSIONS_SHEET);
    subSheet.getRange(1, 1, 1, SUB_COLS.length).setValues([SUB_COLS]);
    subSheet.getRange(1, 1, 1, SUB_COLS.length).setFontWeight('bold');
    subSheet.setFrozenRows(1);
  }
  let examSheet = SS.getSheetByName(EXAMS_SHEET);
  if (!examSheet) {
    examSheet = SS.insertSheet(EXAMS_SHEET);
    examSheet.getRange(1, 1, 1, EXAM_COLS.length).setValues([EXAM_COLS]);
    examSheet.getRange(1, 1, 1, EXAM_COLS.length).setFontWeight('bold');
    examSheet.setFrozenRows(1);
  }
  return { subSheet, examSheet };
}

// ── GET 處理 ────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'all';

  if (action === 'ping') {
    return jsonResp({ ok: true, message: 'Remedial System 後台連線正常', time: new Date().toISOString() });
  }

  if (action === 'all') {
    const { subSheet, examSheet } = setupSheets();
    const submissions = sheetToObjects(subSheet, SUB_COLS);
    const exams = sheetToObjects(examSheet, EXAM_COLS);
    submissions.forEach(s => {
      try { s.misconceptions = JSON.parse(s.misconceptions || '[]'); } catch { s.misconceptions = []; }
      try { s.wrongQuestions = JSON.parse(s.wrongQuestions || '[]'); } catch { s.wrongQuestions = []; }
    });
    exams.forEach(ex => {
      try { ex.roster = JSON.parse(ex.roster || '[]'); } catch { ex.roster = []; }
    });
    return jsonResp({ ok: true, submissions, exams });
  }

  if (action === 'submissions') {
    const { subSheet } = setupSheets();
    const submissions = sheetToObjects(subSheet, SUB_COLS);
    submissions.forEach(s => {
      try { s.misconceptions = JSON.parse(s.misconceptions || '[]'); } catch { s.misconceptions = []; }
      try { s.wrongQuestions = JSON.parse(s.wrongQuestions || '[]'); } catch { s.wrongQuestions = []; }
    });
    return jsonResp({ ok: true, submissions });
  }

  return jsonResp({ ok: false, error: 'Unknown action' });
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
    upsertRow(subSheet, SUB_COLS, sub.id, subToRow(sub));
    return jsonResp({ ok: true, message: '批改記錄已儲存' });
  }

  if (action === 'saveExam') {
    const exam = body.exam;
    if (!exam?.id) return jsonResp({ ok: false, error: '缺少 exam.id' });
    const { examSheet } = setupSheets();
    upsertRow(examSheet, EXAM_COLS, exam.id, examToRow(exam));
    return jsonResp({ ok: true, message: '考卷設定已儲存' });
  }

  if (action === 'syncAll') {
    const subs = body.submissions || [];
    const exams = body.exams || [];
    const { subSheet, examSheet } = setupSheets();
    subs.forEach(s => upsertRow(subSheet, SUB_COLS, s.id, subToRow(s)));
    exams.forEach(ex => upsertRow(examSheet, EXAM_COLS, ex.id, examToRow(ex)));
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
  };
}

// ★ 對應 HTML exam 物件欄位
function examToRow(ex) {
  return {
    id:        ex.id || '',
    name:      ex.name || '',
    subject:   ex.subject || '',
    grade:     ex.grade || '',
    year:      ex.year || '',
    sem:       ex.sem || '',
    createdAt: ex.createdAt || new Date().toLocaleDateString('zh-TW'),
    roster:    JSON.stringify(ex.roster || []),
  };
}
