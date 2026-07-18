/**
 * LifeSpan 健康預測 — Google Apps Script 後台
 * 健康資料的雲端儲存、比對、管理與自動化中心
 *
 * 使用方式：
 * 1. 開啟一份 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上此程式碼（取代所有內容）
 * 3. 執行一次 setupWeeklyTrigger()（授權後會自動排程每週健康週報）
 * 4. 部署 → 新增部署作業 → 網頁應用程式
 *    - 以下列身分執行：我（Me）
 *    - 誰可以存取：所有人（Anyone）
 * 5. 複製部署網址，貼入 health.html「設定 → 雲端後台」→ 測試連線
 *
 * 隱私：資料僅存於「你自己的」Google 試算表，由你本人的帳號執行。
 */

// ── 試算表設定 ──────────────────────────────────
const SS = SpreadsheetApp.getActiveSpreadsheet();
const REC_SHEET = '健康紀錄';   // 每次評估的快照
const LAB_SHEET = '檢驗匯入';   // 健康存摺 / 健檢 / 穿戴裝置匯入的原始檢驗值
const KV_SHEET  = '個人設定';   // 個人檔案 key/value
const REPORT_SHEET = '健康週報'; // 自動化週報存檔

// ★ 健康紀錄欄位 — 與 health.html snapshot 物件對齊
const REC_COLS = [
  'id', 'date', 'sex', 'age', 'height', 'weight', 'waist', 'hr',
  'sbp', 'dbp', 'glucose', 'chol', 'hdl',
  'smoke', 'exercise', 'alcohol', 'sleep', 'sitting', 'diet', 'stress',
  'bmi', 'whtr', 'score', 'bio', 'life', 'cvdLevel', 'dmLevel', 'metsCount', 'source'
];

// ★ 檢驗匯入欄位
const LAB_COLS = ['id', 'importedAt', 'source', 'metric', 'value', 'unit', 'measuredAt'];

// 通知信箱（留空則自動用試算表擁有者信箱）
const NOTIFY_EMAIL = '';

// ── 初始化試算表 ────────────────────────────────
function setupSheets() {
  const rec = ensureSheet(REC_SHEET, REC_COLS);
  const lab = ensureSheet(LAB_SHEET, LAB_COLS);
  const kv  = ensureSheet(KV_SHEET, ['key', 'value']);
  const rep = ensureSheet(REPORT_SHEET, ['generatedAt', 'summary', 'score', 'bio', 'delta']);
  return { rec, lab, kv, rep };
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
  const action = e?.parameter?.action || 'all';

  if (action === 'ping') {
    return jsonResp({ ok: true, message: 'LifeSpan 健康後台連線正常', time: new Date().toISOString() });
  }
  if (action === 'all') {
    const { rec } = setupSheets();
    return jsonResp({ ok: true, snapshots: sheetToObjects(rec, REC_COLS), profile: readProfile() });
  }
  if (action === 'snapshots') {
    const { rec } = setupSheets();
    return jsonResp({ ok: true, snapshots: sheetToObjects(rec, REC_COLS) });
  }
  if (action === 'latest') {
    const { rec } = setupSheets();
    const all = sheetToObjects(rec, REC_COLS);
    return jsonResp({ ok: true, snapshot: all.length ? all[all.length - 1] : null });
  }
  if (action === 'compare') {
    return jsonResp({ ok: true, comparison: compareLatest() });
  }
  if (action === 'labs') {
    const { lab } = setupSheets();
    return jsonResp({ ok: true, labs: sheetToObjects(lab, LAB_COLS) });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── POST 處理 ───────────────────────────────────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResp({ ok: false, error: 'Invalid JSON' }); }
  const action = body.action;

  if (action === 'saveSnapshot') {
    const s = body.snapshot;
    if (!s?.id) return jsonResp({ ok: false, error: '缺少 snapshot.id' });
    const { rec } = setupSheets();
    upsertRow(rec, REC_COLS, s.id, snapToRow(s));
    return jsonResp({ ok: true, message: '健康紀錄已儲存', comparison: compareLatest() });
  }
  if (action === 'syncAll') {
    const snaps = body.snapshots || [];
    const { rec } = setupSheets();
    snaps.forEach(s => { if (s.id) upsertRow(rec, REC_COLS, s.id, snapToRow(s)); });
    if (body.profile) writeProfile(body.profile);
    return jsonResp({ ok: true, message: `同步完成：${snaps.length} 筆健康紀錄`, comparison: compareLatest() });
  }
  if (action === 'saveProfile') {
    writeProfile(body.profile || {});
    return jsonResp({ ok: true, message: '個人檔案已儲存' });
  }
  if (action === 'importLabs') {
    const labs = body.labs || [];
    const { lab } = setupSheets();
    const stamp = new Date().toISOString();
    labs.forEach(l => {
      const id = l.id || Utilities.getUuid();
      appendObj(lab, LAB_COLS, {
        id, importedAt: stamp, source: l.source || 'import',
        metric: l.metric || '', value: l.value ?? '', unit: l.unit || '', measuredAt: l.measuredAt || ''
      });
    });
    return jsonResp({ ok: true, message: `已匯入 ${labs.length} 筆檢驗值`, latest: latestLabValues() });
  }
  if (action === 'deleteSnapshot') {
    const { rec } = setupSheets();
    deleteRow(rec, body.id);
    return jsonResp({ ok: true, message: '紀錄已刪除' });
  }
  if (action === 'clearAll') {
    clearSheetBody(REC_SHEET); clearSheetBody(LAB_SHEET);
    return jsonResp({ ok: true, message: '雲端資料已清除' });
  }
  return jsonResp({ ok: false, error: 'Unknown action: ' + action });
}

// ── 比對邏輯 ────────────────────────────────────
function compareLatest() {
  const { rec } = setupSheets();
  const all = sheetToObjects(rec, REC_COLS);
  if (all.length < 2) return { hasComparison: false, count: all.length };
  const cur = all[all.length - 1], prev = all[all.length - 2];
  const metrics = [
    { k: 'score', label: '健康分', up: true },
    { k: 'bio', label: '生理年齡', up: false },
    { k: 'weight', label: '體重', up: false },
    { k: 'bmi', label: 'BMI', up: false },
    { k: 'sbp', label: '收縮壓', up: false },
    { k: 'glucose', label: '空腹血糖', up: false }
  ];
  const diffs = [];
  metrics.forEach(m => {
    const a = parseFloat(prev[m.k]), b = parseFloat(cur[m.k]);
    if (isNaN(a) || isNaN(b)) return;
    const delta = Math.round((b - a) * 100) / 100;
    const better = m.up ? delta > 0 : delta < 0;
    diffs.push({ label: m.label, prev: a, cur: b, delta, better: delta === 0 ? null : better });
  });
  return { hasComparison: true, from: prev.date, to: cur.date, diffs };
}

// ── 自動化：每週健康週報 ────────────────────────
function setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyAutoReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyAutoReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  return '✅ 已排程每週一早上 8 點自動產生健康週報';
}

function weeklyAutoReport() {
  const { rec, rep } = setupSheets();
  const all = sheetToObjects(rec, REC_COLS);
  if (!all.length) return;
  const cur = all[all.length - 1];
  const cmp = compareLatest();
  const email = NOTIFY_EMAIL || getOwnerEmail();

  // 距上次評估天數 → 提醒
  const daysSince = Math.floor((Date.now() - new Date(cur.date).getTime()) / 86400000);

  let lines = [];
  lines.push(`【LifeSpan 健康週報】${new Date().toLocaleDateString('zh-TW')}`);
  lines.push('');
  lines.push(`綜合健康分：${cur.score}　生理年齡：${cur.bio}（實際 ${cur.age}）`);
  lines.push(`預期壽命估算：${cur.life} 歲　心血管風險：${cur.cvdLevel}　糖尿病風險：${cur.dmLevel}`);
  lines.push('');
  if (cmp.hasComparison) {
    lines.push(`與上次（${cmp.from}）比較：`);
    cmp.diffs.forEach(d => {
      const arrow = d.delta > 0 ? '▲' : d.delta < 0 ? '▼' : '＝';
      const tag = d.better === true ? '✅改善' : d.better === false ? '⚠️退步' : '持平';
      lines.push(`  ${d.label}：${d.prev} → ${d.cur}（${arrow}${Math.abs(d.delta)}）${tag}`);
    });
    lines.push('');
  }
  if (daysSince >= 7) lines.push(`📌 提醒：距上次健康評估已 ${daysSince} 天，建議更新一次數據。`);
  lines.push('');
  lines.push('（本週報由 LifeSpan 健康後台自動產生，僅供自我管理參考，非醫療診斷。）');
  const summary = lines.join('\n');

  appendObj(rep, ['generatedAt', 'summary', 'score', 'bio', 'delta'], {
    generatedAt: new Date().toISOString(), summary,
    score: cur.score, bio: cur.bio,
    delta: cmp.hasComparison ? JSON.stringify(cmp.diffs) : ''
  });

  if (email) {
    try { MailApp.sendEmail(email, '🫀 LifeSpan 健康週報', summary); } catch (err) {}
  }
}

function getOwnerEmail() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// ── 自動化：每日主動守護（危險門檻 + 惡化趨勢 → 主動寄警告）──
function setupDailyWatch() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyHealthWatch') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyHealthWatch').timeBased().everyDays(1).atHour(9).create();
  return '✅ 已排程每日早上 9 點自動健康守護（偵測到警示才寄信）';
}

function dailyHealthWatch() {
  const { rec } = setupSheets();
  const all = sheetToObjects(rec, REC_COLS);
  if (!all.length) return;
  const cur = all[all.length - 1];
  const alerts = evaluateAlerts(cur, all);
  if (!alerts.length) return; // 沒有警示就不打擾

  const email = NOTIFY_EMAIL || getOwnerEmail();
  if (!email) return;

  const rank = { crit: '🔴 立即注意', warn: '🟠 需改善', watch: '🟡 趨勢留意' };
  let lines = ['【LifeSpan 健康守護警示】' + new Date().toLocaleDateString('zh-TW'), ''];
  lines.push('系統偵測到以下需要關注的項目：', '');
  alerts.forEach(a => {
    lines.push(`${rank[a.level]}　${a.title}`);
    lines.push(`　${a.desc}`);
    lines.push(`　👉 建議：${a.action}`, '');
  });
  lines.push('（本警示由 LifeSpan 健康後台每日自動產生，僅供自我管理參考，非醫療診斷。危急症狀請立即就醫。）');
  const body = lines.join('\n');

  // 避免同一天重複寄送
  const props = PropertiesService.getScriptProperties();
  const today = new Date().toISOString().slice(0, 10);
  const sig = today + '|' + alerts.map(a => a.level + a.title).join(';');
  if (props.getProperty('lastAlertSig') === sig) return;
  props.setProperty('lastAlertSig', sig);

  const topLevel = alerts[0].level === 'crit' ? '🔴' : alerts[0].level === 'warn' ? '🟠' : '🟡';
  try { MailApp.sendEmail(email, `${topLevel} LifeSpan 健康守護：發現 ${alerts.length} 項提醒`, body); } catch (err) {}
}

// 伺服器端警示引擎（與前端 analyzeAlerts 門檻一致）
function evaluateAlerts(cur, all) {
  const A = [];
  const n = v => { const x = parseFloat(v); return isNaN(x) ? null : x; };
  const add = (level, title, desc, action) => A.push({ level, title, desc, action });
  const sbp = n(cur.sbp), dbp = n(cur.dbp), glu = n(cur.glucose), bmi = n(cur.bmi), mets = n(cur.metsCount);

  if (sbp >= 180 || dbp >= 110) add('crit', '血壓危急', `血壓 ${cur.sbp}/${cur.dbp} mmHg 達高血壓危象範圍。`, '請儘速就醫評估。');
  else if (sbp >= 140 || dbp >= 90) add('warn', '高血壓', `血壓 ${cur.sbp}/${cur.dbp} mmHg 屬高血壓。`, '減鈉、規律運動、減重；必要時就醫。');
  if (glu >= 200) add('crit', '血糖過高', `空腹血糖 ${cur.glucose} mg/dL 明顯偏高。`, '儘速就醫做進一步檢查。');
  else if (glu >= 126) add('warn', '血糖達糖尿病範圍', `空腹血糖 ${cur.glucose} mg/dL（≥126）。`, '就醫確認並控糖。');
  else if (glu >= 100) add('watch', '血糖偏高（前期）', `空腹血糖 ${cur.glucose} mg/dL 屬前期。`, '減糖、增纖維與運動，可逆轉。');
  if (bmi >= 35) add('warn', '重度肥胖', `BMI ${bmi.toFixed(1)}。`, '循序減重 5-10%。');
  if (String(cur.smoke) === 'current') add('warn', '目前吸菸', '吸菸是最強的可修正致命因子。', '尋求戒菸門診。');
  if (mets >= 3) add('warn', '符合代謝症候群', `已符合 ${mets}/5 項。`, '多管齊下改善腰圍/血壓/血糖/血脂。');

  // 惡化趨勢
  const trend = key => {
    const pts = all.filter(s => s[key] !== '' && s[key] != null).map(s => ({ t: new Date(s.date).getTime(), v: parseFloat(s[key]) })).filter(o => !isNaN(o.v));
    if (pts.length < 3) return null;
    pts.sort((a, b) => a.t - b.t);
    const months = Math.max((pts[pts.length - 1].t - pts[0].t) / 2.592e9, 0.1);
    return { first: pts[0].v, last: pts[pts.length - 1].v, perMonth: (pts[pts.length - 1].v - pts[0].v) / months };
  };
  [['weight', '體重', 'kg', 0.7], ['sbp', '收縮壓', 'mmHg', 2], ['glucose', '空腹血糖', 'mg/dL', 2]].forEach(([k, lbl, u, rate]) => {
    const t = trend(k);
    if (t && t.perMonth >= rate) add('watch', `${lbl}持續上升`, `${lbl}由 ${t.first} → ${t.last}${u}（每月約 +${t.perMonth.toFixed(1)}）。`, '及早調整生活型態，密切監測。');
  });

  A.sort((a, b) => ({ crit: 1, warn: 2, watch: 3 }[a.level] - { crit: 1, warn: 2, watch: 3 }[b.level]));
  return A;
}

// ── 個人檔案 KV ─────────────────────────────────
function readProfile() {
  const { kv } = setupSheets();
  const data = kv.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { try { obj[data[i][0]] = JSON.parse(data[i][1]); } catch { obj[data[i][0]] = data[i][1]; } }
  }
  return obj;
}
function writeProfile(profile) {
  const { kv } = setupSheets();
  Object.keys(profile).forEach(k => {
    const v = typeof profile[k] === 'object' ? JSON.stringify(profile[k]) : profile[k];
    upsertRow(kv, ['key', 'value'], k, { key: k, value: v });
  });
}

// 匯總每個檢驗指標的最新值（供前端自動帶入）
function latestLabValues() {
  const { lab } = setupSheets();
  const all = sheetToObjects(lab, LAB_COLS);
  const latest = {};
  all.forEach(l => {
    const t = new Date(l.measuredAt || l.importedAt).getTime() || 0;
    if (!latest[l.metric] || t >= latest[l.metric]._t) latest[l.metric] = { value: l.value, unit: l.unit, _t: t };
  });
  const out = {};
  Object.keys(latest).forEach(k => out[k] = latest[k].value);
  return out;
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
function deleteRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) { if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; } }
}
function clearSheetBody(name) {
  const sh = SS.getSheetByName(name);
  if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}
function snapToRow(s) {
  const r = {};
  REC_COLS.forEach(c => r[c] = s[c] !== undefined && s[c] !== null ? s[c] : '');
  r.id = s.id;
  r.date = s.date || new Date().toISOString();
  r.source = s.source || 'app';
  return r;
}
