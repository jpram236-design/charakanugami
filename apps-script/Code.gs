/**
 * चरकानुगामी — सामुदायिक मंच का डेटाबेस (Google Apps Script)
 *
 * एक बार के चरण:
 *  1. script.google.com → New project → यह पूरा कोड चिपकाएँ → Save
 *  2. ऊपर फ़ंक्शन "setup" चुनकर ▶ Run (अनुमति दें)
 *  3. Deploy → New deployment → ⚙ → Web app
 *       Execute as: Me    |    Who has access: Anyone   → Deploy
 *  4. मिला "Web app URL" (…/exec) वेबसाइट में डालने के लिए भेजें।
 *
 * देखरेख: "मंच-डेटा" शीट में किसी पंक्ति को हटाने से वह पोस्ट/उत्तर साइट से हट जाता है।
 */

const COLLECTIONS = ['questions', 'answers', 'experiences', 'expComments',
  'jPosts', 'jComments', 'jLikes', 'sPosts', 'sComments', 'sLikes', 'quizzes', 'quizSubs', 'yatharth'];
const ADMIN_ONLY = ['yatharth']; // केवल संचालक लिख सकते हैं
// किसकी "मूल" प्रविष्टि का लेखक इसे हटा सकता है: [क्षेत्र, मूल संग्रह]
const PARENT = {
  answers: ['qid', 'questions'], expComments: ['eid', 'experiences'],
  jComments: ['pid', 'jPosts'], sComments: ['pid', 'sPosts'], quizSubs: ['qzid', 'quizzes']
};
const CHUNK = 45000, MAX_CHUNKS = 6, MAX_DOC = CHUNK * MAX_CHUNKS;
const HEAD = ['संग्रह', 'id', 'लेखक-id', 'लेखक का नाम', 'झलक', 'बना', 'अद्यतन'];
const NCOL = HEAD.length + MAX_CHUNKS;
const RATE_PER_MIN = 40;

/* ---------------- एक बार चलाएँ ---------------- */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (!id) {
    const ss = SpreadsheetApp.create('चरकानुगामी — मंच-डेटा');
    const d = ss.getSheets()[0];
    d.setName('मंच-डेटा');
    d.getRange(1, 1, 1, NCOL).setValues([HEAD.concat(Array.from({ length: MAX_CHUNKS }, (_, i) => 'डेटा-' + (i + 1)))]);
    d.setFrozenRows(1);
    const u = ss.insertSheet('सदस्य-नाम');
    u.getRange(1, 1, 1, 4).setValues([['id', 'नाम', 'कुंजी', 'पहली बार']]);
    u.setFrozenRows(1);
    id = ss.getId();
    props.setProperty('SHEET_ID', id);
  }
  let key = props.getProperty('ADMIN_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('ADMIN_KEY', key); }
  console.log('✅ तैयार। डेटा-शीट: ' + SpreadsheetApp.openById(id).getUrl());
  console.log('🔑 संचालक-कुंजी (गोपनीय रखें): ' + key);
  console.log('अगला चरण: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone');
}

/* ---------------- पढ़ना ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const out = JSON.parse(readAll_());
  out.admin = !!(p.key && p.key === adminKey_());
  return json_(out);
}

function readAll_() {
  const cache = CacheService.getScriptCache();
  const ver = cache.get('ver') || '0';
  const n = parseInt(cache.get('n_' + ver) || '0', 10);
  if (n) {
    const keys = []; for (let i = 0; i < n; i++) keys.push('c_' + ver + '_' + i);
    const got = cache.getAll(keys);
    if (Object.keys(got).length === n) return keys.map((k) => got[k]).join('');
  }
  const s = sheets_();
  const rows = s.d.getLastRow() > 1 ? s.d.getRange(2, 1, s.d.getLastRow() - 1, NCOL).getValues() : [];
  const docs = [];
  rows.forEach((r) => {
    if (!r[0] || !r[1]) return;
    const raw = r.slice(HEAD.length).join('');
    try { docs.push({ c: String(r[0]), id: String(r[1]), d: JSON.parse(raw) }); } catch (err) {}
  });
  const names = {};
  if (s.u.getLastRow() > 1) s.u.getRange(2, 1, s.u.getLastRow() - 1, 2).getValues().forEach((r) => { if (r[0]) names[r[0]] = String(r[1] || ''); });
  const str = JSON.stringify({ docs: docs, names: names, t: Date.now() });
  try {
    const parts = {}; let i = 0;
    for (let k = 0; k < str.length; k += 90000) parts['c_' + ver + '_' + (i++)] = str.slice(k, k + 90000);
    if (i <= 80) { cache.putAll(parts, 60); cache.put('n_' + ver, String(i), 60); }
  } catch (err) {}
  return str;
}

/* ---------------- लिखना ---------------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, code: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, code: 'busy' });
  try {
    const res = handle_(req);
    if (res.ok) bump_();
    return json_(res);
  } catch (err) {
    return json_({ ok: false, code: 'error', message: String(err) });
  } finally { lock.releaseLock(); }
}

function handle_(req) {
  const uid = String(req.uid || ''), token = String(req.token || '');
  if (!/^u_[A-Za-z0-9]{10,40}$/.test(uid) || token.length < 16) return { ok: false, code: 'permission_denied' };
  const admin = !!(req.key && req.key === adminKey_());
  const s = sheets_();

  // पहचान: पहली बार कुंजी दर्ज, बाद में मिलान
  const users = s.u.getLastRow() > 1 ? s.u.getRange(2, 1, s.u.getLastRow() - 1, 4).getValues() : [];
  const hash = sha_(token);
  let urow = -1;
  for (let i = 0; i < users.length; i++) if (users[i][0] === uid) { urow = i + 2; break; }
  if (urow > 0 && users[urow - 2][2] !== hash) return { ok: false, code: 'permission_denied' };
  const name = req.name !== undefined ? String(req.name).replace(/[<>]/g, '').trim().slice(0, 40) : null;
  if (urow < 0) {
    s.u.appendRow([uid, name || '', hash, new Date()]);
    urow = s.u.getLastRow();
  } else if (name !== null) s.u.getRange(urow, 2).setValue(name);
  if (req.action === 'name') return { ok: true };

  // दर-सीमा
  const cache = CacheService.getScriptCache();
  const rk = 'rl_' + uid + '_' + Math.floor(Date.now() / 60000);
  const cnt = parseInt(cache.get(rk) || '0', 10) + 1;
  cache.put(rk, String(cnt), 120);
  if (cnt > RATE_PER_MIN && !admin) return { ok: false, code: 'rate_limited' };

  const ops = Array.isArray(req.ops) ? req.ops.slice(0, 20) : [];
  const last = s.d.getLastRow();
  const rows = last > 1 ? s.d.getRange(2, 1, last - 1, NCOL).getValues() : [];
  const index = {};
  rows.forEach((r, i) => { if (r[0]) index[r[0] + '/' + r[1]] = i + 2; });
  const docAt = (row) => { try { return JSON.parse(rows[row - 2].slice(HEAD.length).join('')); } catch (err) { return {}; } };
  const parentAuthor = (c, d) => {
    const p = PARENT[c]; if (!p || !d) return null;
    const row = index[p[1] + '/' + d[p[0]]]; return row ? (docAt(row).author || rows[row - 2][2]) : null;
  };
  const toDelete = [];
  const nameOf = name !== null ? name : String(s.u.getRange(urow, 2).getValue() || '');

  for (const op of ops) {
    const c = String(op.c || ''), id = String(op.id || '');
    if (COLLECTIONS.indexOf(c) < 0 || !/^[A-Za-z0-9_\-]{1,120}$/.test(id)) return { ok: false, code: 'permission_denied' };
    if (ADMIN_ONLY.indexOf(c) >= 0 && !admin) return { ok: false, code: 'permission_denied' };
    const row = index[c + '/' + id];
    const old = row ? docAt(row) : null;
    const oldAuthor = row ? (old.author || rows[row - 2][2]) : null;

    if (op.op === 'delete') {
      if (!row) continue;
      if (!(admin || oldAuthor === uid || parentAuthor(c, old) === uid)) return { ok: false, code: 'permission_denied' };
      toDelete.push(row);
      continue;
    }
    if (op.op !== 'set' && op.op !== 'update') return { ok: false, code: 'bad_request' };
    let d = (op.d && typeof op.d === 'object') ? op.d : {};
    if (row) {
      const may = admin || oldAuthor === uid || (c === 'quizSubs' && parentAuthor(c, old) === uid);
      if (!may) return { ok: false, code: 'permission_denied' };
      d = op.op === 'update' ? Object.assign({}, old, d) : d;
      d.author = oldAuthor;
    } else {
      if (/Likes$/.test(c) && !id.endsWith('__' + uid)) return { ok: false, code: 'permission_denied' };
      d.author = uid;
    }
    const str = JSON.stringify(d);
    if (str.length > MAX_DOC) return { ok: false, code: 'too_large' };
    const chunks = []; for (let k = 0; k < MAX_CHUNKS; k++) chunks.push(str.slice(k * CHUNK, (k + 1) * CHUNK));
    const preview = String(d.title || d.text || d.body || d.vyadhi || '').replace(/\s+/g, ' ').slice(0, 120);
    const authorName = row ? rows[row - 2][3] : nameOf;
    const values = [c, id, d.author, authorName, preview, row ? rows[row - 2][5] : new Date(), new Date()].concat(chunks);
    if (row) { s.d.getRange(row, 1, 1, NCOL).setValues([values]); rows[row - 2] = values; }
    else { s.d.appendRow(values); const r = s.d.getLastRow(); index[c + '/' + id] = r; rows[r - 2] = values; }
  }
  toDelete.sort((a, b) => b - a).forEach((r) => s.d.deleteRow(r));
  return { ok: true };
}

/* ---------------- सहायक ---------------- */
function sheets_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('पहले setup चलाएँ');
  const ss = SpreadsheetApp.openById(id);
  return { d: ss.getSheetByName('मंच-डेटा'), u: ss.getSheetByName('सदस्य-नाम') };
}
function adminKey_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '\u0000'; }
function bump_() { CacheService.getScriptCache().put('ver', String(Date.now()), 21600); }
function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s).map((b) => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
