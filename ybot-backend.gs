/**
 * Ybot 個人助理 — Google Apps Script 後台
 * 待辦／筆記／提醒的雲端儲存 + Gmail／日曆彙整 + 每日主動簡報 + 到點提醒信
 *
 * 使用方式：
 * 1. 開啟一份 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 執行一次 setupDailyBrief()（授權後排程「每日簡報」）
 *    再執行一次 setupReminderWatch()（排程「到點提醒」，每 30 分鐘檢查一次）
 * 4. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 5. 複製部署網址，貼入 ybot.html「⚙️ 設定 → 雲端後台」→ 測試連線
 *
 * 授權提醒：Gmail／日曆彙整需要額外授權 Gmail 與 Calendar 權限，
 * 重新部署後首次執行會跳出授權視窗，請同意（僅你本人帳號讀取，不外傳）。
 *
 * 隱私：所有資料僅存於「你自己的」Google 試算表，由你本人的帳號執行；
 * Gmail／日曆內容只在你自己部署的後台內彙整、回傳給你自己的前端。
 */

// ── 試算表設定 ──────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const NOTE_SHEET = '待辦與筆記';
const KV_SHEET = '個人設定';
const BRIEF_SHEET = '每日簡報存檔';

const NOTE_COLS = ['id', 'type', 'content', 'dueAt', 'done', 'createdAt', 'notifiedAt'];
// type: note（瑣事筆記）／ todo（待辦）／ reminder（有時間點的提醒，到點會主動寄信）

// 通知信箱（留空則自動用試算表擁有者信箱）
const NOTIFY_EMAIL = '';

// ── AI 設定（選填，供每日簡報生成友善摘要用）────
// 執行一次 setAiConfig('gemini','你的KEY') 或 setAiConfig('openai','sk-...')
function setAiConfig(provider, key) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AI_PROVIDER', provider || 'gemini');
  props.setProperty('AI_KEY', key || '');
  return '✅ AI 設定完成：' + (provider || 'gemini');
}
function getAiConfig() {
  const props = PropertiesService.getScriptProperties();
  return { provider: props.getProperty('AI_PROVIDER') || 'gemini', key: props.getProperty('AI_KEY') || '' };
}
function callAI(prompt) {
  const { provider, key } = getAiConfig();
  if (!key) return '';
  try {
    if (provider === 'openai') {
      const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({ model: 'gpt-4o-mini', messages: [{ role: 'user', content: prompt }], temperature: 0.6 })
      });
      const d = JSON.parse(res.getContentText());
      return d.choices && d.choices[0] ? d.choices[0].message.content : '';
    } else {
      const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + key;
      const res = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
      });
      const d = JSON.parse(res.getContentText());
      return d.candidates && d.candidates[0] ? d.candidates[0].content.parts[0].text : '';
    }
  } catch (err) { return ''; }
}

// ── 初始化試算表 ────────────────────────────────
function setupSheets() {
  const note = ensureSheet(NOTE_SHEET, NOTE_COLS);
  const kv = ensureSheet(KV_SHEET, ['key', 'value']);
  const brief = ensureSheet(BRIEF_SHEET, ['generatedAt', 'summary']);
  return { note, kv, brief };
}
function ensureSheet(name, cols) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    sh.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── GET 處理 ────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'context';

  if (action === 'ping') {
    return jsonResp({ ok: true, message: 'Ybot 後台連線正常', time: new Date().toISOString() });
  }
  if (action === 'notes') {
    const { note } = setupSheets();
    return jsonResp({ ok: true, notes: sheetToObjects(note, NOTE_COLS) });
  }
  if (action === 'context') {
    return jsonResp({ ok: true, context: buildContext() });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── POST 處理 ───────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResp({ ok: false, error: 'Invalid JSON' }); }
  const action = body.action;

  if (action === 'addNote') {
    const { note } = setupSheets();
    const n = body.note || {};
    const id = n.id || Utilities.getUuid();
    appendObj(note, NOTE_COLS, {
      id, type: n.type || 'note', content: n.content || '', dueAt: n.dueAt || '',
      done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(), notifiedAt: ''
    });
    return jsonResp({ ok: true, message: '已新增', id });
  }
  if (action === 'updateNote') {
    const { note } = setupSheets();
    updatePartial(note, NOTE_COLS, body.id, body.patch || {});
    return jsonResp({ ok: true, message: '已更新' });
  }
  if (action === 'deleteNote') {
    const { note } = setupSheets();
    deleteRow(note, body.id);
    return jsonResp({ ok: true, message: '已刪除' });
  }
  if (action === 'syncAll') {
    const { note } = setupSheets();
    (body.notes || []).forEach(n => {
      if (!n.id) return;
      upsertRow(note, NOTE_COLS, n.id, {
        id: n.id, type: n.type || 'note', content: n.content || '', dueAt: n.dueAt || '',
        done: n.done ? 'true' : '', createdAt: n.createdAt || new Date().toISOString(), notifiedAt: n.notifiedAt || ''
      });
    });
    return jsonResp({ ok: true, message: `同步完成：${(body.notes || []).length} 筆` });
  }
  if (action === 'saveLinkedBackends') {
    writeKv({ remediationBackendUrl: body.remediationBackendUrl || '', healthBackendUrl: body.healthBackendUrl || '' });
    return jsonResp({ ok: true, message: '已儲存整合設定' });
  }
  if (action === 'saveAiConfig') {
    setAiConfig(body.provider, body.key);
    return jsonResp({ ok: true, message: '已儲存 AI 設定' });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── 彙整上下文：Gmail + 日曆 + 待辦提醒 + 其他系統摘要 ──
function buildContext() {
  return {
    generatedAt: new Date().toISOString(),
    notes: sheetToObjects(setupSheets().note, NOTE_COLS),
    gmail: getGmailDigest(),
    calendar: getCalendarDigest(),
    linked: getLinkedSummaries()
  };
}

// Gmail 彙整：近 2 天未讀信件，僅取主旨／寄件者／時間／摘要片段（不含全文）
function getGmailDigest() {
  try {
    const threads = GmailApp.search('is:unread newer_than:2d', 0, 15);
    const out = [];
    threads.forEach(t => {
      const msgs = t.getMessages();
      const last = msgs[msgs.length - 1];
      out.push({
        subject: t.getFirstMessageSubject(),
        from: last.getFrom(),
        date: last.getDate().toISOString(),
        snippet: (last.getPlainBody() || '').slice(0, 120).replace(/\s+/g, ' ')
      });
    });
    return out;
  } catch (err) { return []; }
}

// 日曆彙整：未來 7 天的行程
function getCalendarDigest() {
  try {
    const now = new Date();
    const until = new Date(now.getTime() + 7 * 86400000);
    const events = CalendarApp.getDefaultCalendar().getEvents(now, until);
    return events.slice(0, 20).map(ev => ({
      title: ev.getTitle(),
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString(),
      allDay: ev.isAllDayEvent()
    }));
  } catch (err) { return []; }
}

// 串接考卷批改／健康預測兩套系統的後台（若已在設定中填入網址）
function getLinkedSummaries() {
  const kv = readKv();
  const out = {};
  if (kv.remediationBackendUrl) {
    try {
      const res = UrlFetchApp.fetch(kv.remediationBackendUrl + '?action=all', { muteHttpExceptions: true });
      const d = JSON.parse(res.getContentText());
      if (d.ok) {
        const subs = d.submissions || [];
        const recent = subs.slice(-10);
        const avg = recent.length ? Math.round(recent.reduce((s, r) => s + (Number(r.percentage) || 0), 0) / recent.length) : null;
        const misconceptions = {};
        recent.forEach(r => (r.misconceptions || []).forEach(m => { misconceptions[m] = (misconceptions[m] || 0) + 1; }));
        out.remediation = {
          totalSubmissions: subs.length,
          recentAvgPercentage: avg,
          topMisconceptions: Object.entries(misconceptions).sort((a, b) => b[1] - a[1]).slice(0, 5).map(x => x[0])
        };
      }
    } catch (err) { /* 忽略連線失敗 */ }
  }
  if (kv.healthBackendUrl) {
    try {
      const res = UrlFetchApp.fetch(kv.healthBackendUrl + '?action=latest', { muteHttpExceptions: true });
      const d = JSON.parse(res.getContentText());
      if (d.ok && d.snapshot) {
        out.health = {
          date: d.snapshot.date, score: d.snapshot.score, bio: d.snapshot.bio,
          cvdLevel: d.snapshot.cvdLevel, dmLevel: d.snapshot.dmLevel
        };
      }
    } catch (err) { /* 忽略連線失敗 */ }
  }
  return out;
}

// ── 自動化：每日主動簡報 ────────────────────────
function setupDailyBrief() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyBrief') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBrief').timeBased().everyDays(1).atHour(7).create();
  return '✅ 已排程每日早上 7 點自動產生 Ybot 簡報';
}

function dailyBrief() {
  const ctx = buildContext();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const today = new Date().toISOString().slice(0, 10);
  const overdue = ctx.notes.filter(n => n.type !== 'note' && n.done !== 'true' && n.dueAt && n.dueAt.slice(0, 10) < today);
  const dueToday = ctx.notes.filter(n => n.type !== 'note' && n.done !== 'true' && n.dueAt && n.dueAt.slice(0, 10) === today);
  const openTodos = ctx.notes.filter(n => n.type === 'todo' && n.done !== 'true' && !n.dueAt);

  let lines = [`【Ybot 每日簡報】${new Date().toLocaleDateString('zh-TW')}`, ''];

  if (ctx.calendar.length) {
    lines.push('📅 未來行程：');
    ctx.calendar.slice(0, 8).forEach(ev => {
      const t = ev.allDay ? '全天' : new Date(ev.start).toLocaleString('zh-TW', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      lines.push(`　${t}　${ev.title}`);
    });
    lines.push('');
  }
  if (overdue.length || dueToday.length) {
    lines.push('⏰ 待辦與提醒：');
    overdue.forEach(n => lines.push(`　🔴 已逾期：${n.content}（${n.dueAt.slice(0, 10)}）`));
    dueToday.forEach(n => lines.push(`　🟡 今天到期：${n.content}`));
    lines.push('');
  }
  if (openTodos.length) {
    lines.push(`📝 未完成待辦（無期限）共 ${openTodos.length} 項：`);
    openTodos.slice(0, 8).forEach(n => lines.push(`　・${n.content}`));
    lines.push('');
  }
  if (ctx.gmail.length) {
    lines.push(`📬 未讀重要信件（近 2 天，共 ${ctx.gmail.length} 封）：`);
    ctx.gmail.slice(0, 5).forEach(m => lines.push(`　・${m.subject}（${m.from}）`));
    lines.push('');
  }
  if (ctx.linked.remediation) {
    const r = ctx.linked.remediation;
    lines.push(`🎯 考卷批改：累計 ${r.totalSubmissions} 筆，近期平均 ${r.recentAvgPercentage ?? '—'}%，常見迷思：${(r.topMisconceptions || []).join('、') || '無'}`);
  }
  if (ctx.linked.health) {
    const h = ctx.linked.health;
    lines.push(`🫀 健康：最新健康分 ${h.score}，生理年齡 ${h.bio}（${h.date}）`);
  }
  lines.push('');

  const aiText = aiDailyNarrative(ctx, overdue, dueToday);
  if (aiText) lines.push(aiText);

  lines.push('（本簡報由 Ybot 後台每日自動產生。）');
  const summary = lines.join('\n');

  appendObj(setupSheets().brief, ['generatedAt', 'summary'], { generatedAt: new Date().toISOString(), summary });
  try { MailApp.sendEmail(email, '🤖 Ybot 每日簡報', summary); } catch (err) { }
}

function aiDailyNarrative(ctx, overdue, dueToday) {
  const { key } = getAiConfig();
  if (!key) return '';
  const prompt = `你是使用者 Young 的個人助理 Ybot，個性溫暖直接。以下是今天的彙整資料：\n` +
    `行程：${JSON.stringify(ctx.calendar.slice(0, 8))}\n` +
    `逾期：${overdue.map(n => n.content).join('、') || '無'}\n` +
    `今天到期：${dueToday.map(n => n.content).join('、') || '無'}\n` +
    `未讀重要信件數：${ctx.gmail.length}\n` +
    `其他系統摘要：${JSON.stringify(ctx.linked)}\n\n` +
    `請用繁體中文寫一段 150 字內的「今日提醒」，語氣自然像朋友提醒、不要條列重複上面已列的內容，重點放在「今天最該優先做的 1-2 件事」與一句鼓勵。`;
  const out = callAI(prompt);
  return out ? ('💬 Ybot 想跟你說：\n' + out + '\n') : '';
}

// ── 自動化：到點提醒（每 30 分鐘檢查一次）──────
function setupReminderWatch() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'reminderWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reminderWatch').timeBased().everyMinutes(30).create();
  return '✅ 已排程每 30 分鐘檢查一次到點提醒';
}

function reminderWatch() {
  const { note } = setupSheets();
  const all = sheetToObjects(note, NOTE_COLS);
  const now = new Date();
  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const due = all.filter(n => n.type === 'reminder' && n.done !== 'true' && !n.notifiedAt && n.dueAt && new Date(n.dueAt) <= now);
  due.forEach(n => {
    try {
      MailApp.sendEmail(email, '⏰ Ybot 提醒：' + n.content, `到了你設定的提醒時間：\n\n${n.content}\n\n（設定時間：${n.dueAt}）`);
      updatePartial(note, NOTE_COLS, n.id, { notifiedAt: new Date().toISOString() });
    } catch (err) { /* 忽略單筆寄送失敗 */ }
  });
}

function getOwnerEmail() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// ── KV 設定 ─────────────────────────────────────
function readKv() {
  const { kv } = setupSheets();
  const data = kv.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) obj[data[i][0]] = data[i][1]; }
  return obj;
}
function writeKv(map) {
  const { kv } = setupSheets();
  Object.keys(map).forEach(k => upsertRow(kv, ['key', 'value'], k, { key: k, value: map[k] }));
}

// ── 通用工具 ────────────────────────────────────
function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheetToObjects(sheet, cols) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => {
    const o = {}; cols.forEach((c, i) => o[c] = row[i] ?? ''); return o;
  }).filter(o => o.id);
}
function appendObj(sheet, cols, obj) {
  sheet.appendRow(cols.map(c => obj[c] !== undefined ? obj[c] : ''));
}
function upsertRow(sheet, cols, id, rowData) {
  const data = sheet.getDataRange().getValues();
  let target = -1;
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { target = i + 1; break; } }
  const row = cols.map(c => rowData[c] !== undefined ? rowData[c] : '');
  if (target > 0) sheet.getRange(target, 1, 1, cols.length).setValues([row]);
  else sheet.appendRow(row);
}
function updatePartial(sheet, cols, id, patch) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const rowData = {};
      cols.forEach((c, j) => rowData[c] = data[i][j]);
      Object.assign(rowData, patch);
      const row = cols.map(c => rowData[c] !== undefined ? rowData[c] : '');
      sheet.getRange(i + 1, 1, 1, cols.length).setValues([row]);
      return;
    }
  }
}
function deleteRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; } }
}
