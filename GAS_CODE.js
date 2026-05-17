/**
 * Google Apps Script Backend for Examiner Filter Pro (v3.0 - Optimized)
 * Paste this code into your Google Apps Script editor (Extensions > Apps Script).
 * 
 * 1. Replace the SPREADSHEET_ID below with your actual ID.
 * 2. Click "Deploy" > "New Deployment" > "Web App".
 * 3. Set "Execute as" to "Me" and "Who has access" to "Anyone".
 * 4. Copy the Web App URL and paste it into the "GAS_DEPLOYMENT_URL" secret in AI Studio.
 */

/*************** CONFIG (Default fallbacks) ***************/
let SPREADSHEET_ID = '1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU'; 
let SHEET_NAME     = 'Examiner Information';

// Security: Use dynamic config if passed from Express server
function setConfig_(e, content) {
  const qId = e?.parameter?.ssId || content?.ssId;
  const qSh = e?.parameter?.sheetName || content?.sheetName;
  if (qId) SPREADSHEET_ID = qId;
  if (qSh) SHEET_NAME = qSh;
}

const ALLOW = {
  ENGLISH: 55,
  BANGLA: 48,
  PHYSICS: 48,
  CHEMISTRY: 48,
  MATH: 48,
  BIOLOGY: 48,
  ICT: 48
};

const BLANK_LABEL = '(Blank)';
const BLANK_KEY   = '__blank__';
const ALL_SUBJECT_KEYS = ['english','bangla','physics','chemistry','math','biology','ict'];

// ── Column indices (0-based) ──────────────────────────────
const COL = {
  SL:          0,   // A
  NAME:        1,   // B (Nick Name)
  STATUS:      2,   // C
  TPIN:        3,   // D
  INST:        4,   // E
  DEPT:        5,   // F
  BATCH:       6,   // G (HSC Batch)
  RM:          7,   // H
  REMARKED_BY: 8,   // I
  MOB1:        9,   // J (Mobile Number)
  ALT:         10,  // K (Alternate)
  NAGAD:       11,  // L (Mobile Banking / Nagad)
  
  EN:          61,  // BJ English(%)
  BN:          64,  // BM Bangla(%)
  PHY:         67,  // BP Physics(%)
  CHEM:        70,  // BS Chemistry(%)
  MATH:        73,  // BV Math(%)
  BIO:         76,  // BY Biology(%)
  ICT:         79,  // CB ICT(%)
  
  TRAIN:       82,  // CE Training Report
  TRAIN_DATE:  83,  // CF Training Date
  CAMPUS:      88,  // CK Campus
  REMARK_RAW:  92   // CQ Remark
};

/*************** MEMORY CACHE ***************/
let MEM_STORE = {
  loadedAt:     0,
  dataTtlMs:    7200000,  // 2 hours (Internal sheet data persistence)
  optTtlMs:     14400000, // 4 hours (Filter options persistence)
  lastRowCheck: 0,
  lastRowTtl:   60000,    // 60s for row check
  lastRowCount: 0,
  header:       null,
  body:         null,
  options:      null,
  optLoadedAt:  0,
  // Indexes for O(1) filtering
  instIdx:      null,
  deptIdx:      null,
  batchIdx:     null,
  trainIdx:     null,
  campusIdx:    null,
  tpinIdx:      null
};

/*************** ENTRY POINTS ***************/

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  try {
    const action = (e && e.parameter) ? e.parameter.action : null;
    let content = {};
    
    if (e && e.postData && e.postData.contents) {
      try { content = JSON.parse(e.postData.contents); } catch (err) {}
    }

    setConfig_(e, content);
    const finalAction = action || content.action;
    
    if (finalAction === 'ping') {
      return respond_({ success: true, pong: true, version: "3.0.0", time: new Date().toISOString() });
    }

    const filters  = content.filters  || (e.parameter.filters ? JSON.parse(e.parameter.filters) : null);
    const page     = content.page     || e.parameter.page;
    const pageSize = content.pageSize || e.parameter.pageSize;
    const query    = content.query    || e.parameter.query;

    let result;

    if (finalAction === 'options' || finalAction === 'filterOptions') {
      result = getFilterOptionsFast();
    } else if (finalAction === 'filter') {
      result = getFilteredDataFast(filters, page, pageSize);
    } else if (finalAction === 'lookup') {
      result = lookupByQuery(query);
    } else if (finalAction === 'sync') {
      result = getSheetRowCount();
    } else if (finalAction === 'clearCache') {
      result = clearFastCache();
    } else {
      if (!finalAction) return respond_({ success: true, message: "GAS Ready" });
      result = { success: false, error: 'Unknown action: ' + finalAction };
    }
    
    return respond_(result);
  } catch (err) {
    return respond_({ success: false, error: err.toString(), hint: "Verify Spreadsheet ID." });
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/*************** HELPERS ***************/
function normalize(str) {
  if (str === null || str === undefined) return '';
  return String(str).trim().toLowerCase();
}

function isBlankish_(v) {
  return !v || String(v).trim() === '';
}

function openSheetStrict_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  return sh;
}

function toNum_(v) {
  const n = Number(String(v ?? '').trim());
  return Number.isFinite(n) ? n : NaN;
}

function cleanMobile_(v) {
  let m = String(v ?? '').trim().replace(/\D/g, '');
  if (m.length === 10) m = '0' + m;
  return m;
}

function subjectThresholdDynamic_(k, allowEnglish, allowOthers) {
  const en  = Number.isFinite(allowEnglish) ? allowEnglish : ALLOW.ENGLISH;
  const oth = Number.isFinite(allowOthers)  ? allowOthers  : 48;
  return k === 'english' ? en : oth;
}

function subjectValue_(r, k) {
  const colMap = {
    english: COL.EN,
    bangla: COL.BN,
    physics: COL.PHY,
    chemistry: COL.CHEM,
    math: COL.MATH,
    biology: COL.BIO,
    ict: COL.ICT
  };
  return toNum_(r[colMap[k]]);
}

function isAllowedBySubjectsDynamic_(r, subjectsSelected, subjectLogic, allowEnglish, allowOthers) {
  let keys = (subjectsSelected || []).filter(Boolean);
  if (keys.length === 0) keys = ALL_SUBJECT_KEYS;
  const mode = (subjectLogic === 'all') ? 'all' : 'any';

  if (mode === 'all') {
    for (const k of keys) {
      const th = subjectThresholdDynamic_(k, allowEnglish, allowOthers);
      const v  = subjectValue_(r, k);
      if (isNaN(v) || v < th) return false;
    }
    return true;
  }
  for (const k of keys) {
    const th = subjectThresholdDynamic_(k, allowEnglish, allowOthers);
    const v  = subjectValue_(r, k);
    if (!isNaN(v) && v >= th) return true;
  }
  return false;
}

/*************** INDEX BUILDER ***************/
function buildIndexes_(body) {
  const instIdx = new Map(), deptIdx = new Map(), batchIdx = new Map(), trainIdx = new Map(), campusIdx = new Map(), tpinIdx = new Map();
  const add = (map, val, i) => {
    const key = isBlankish_(val) ? BLANK_KEY : normalize(val);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(i);
  };
  for (let i = 0; i < body.length; i++) {
    const r = body[i];
    add(instIdx, r[COL.INST], i);
    add(deptIdx, r[COL.DEPT], i);
    add(batchIdx, r[COL.BATCH], i);
    add(trainIdx, r[COL.TRAIN], i);
    add(campusIdx, r[COL.CAMPUS], i);
    add(tpinIdx, r[COL.TPIN], i);
  }
  return { instIdx, deptIdx, batchIdx, trainIdx, campusIdx, tpinIdx };
}

/*************** RESULT COLUMNS ***************/
function buildKeepColsAndHeader_(subjectKeys) {
  const keepCols = [COL.SL, COL.NAME, COL.TPIN, COL.INST, COL.DEPT, COL.BATCH, COL.RM, COL.MOB1, COL.ALT];
  const header   = ['SL', 'Nick Name', 'T-PIN', 'Inst.', 'Dept.', 'HSC Batch', 'Rm', 'Mobile Number', 'Alternate'];
  const SUBJECT_MAP = { english: COL.EN, bangla: COL.BN, physics: COL.PHY, chemistry: COL.CHEM, math: COL.MATH, biology: COL.BIO, ict: COL.ICT };

  for (const k of subjectKeys) {
    if (SUBJECT_MAP[k]) { keepCols.push(SUBJECT_MAP[k]); header.push(k.charAt(0).toUpperCase() + k.slice(1) + '(%)'); }
  }
  keepCols.push(COL.TRAIN, COL.CAMPUS, -1);
  header.push('Training Report', 'Physical Campus', 'Allow Status');
  return { keepCols, header };
}

/*************** SHEET STORE ***************/
function getSheetStore_() {
  const now = Date.now(), sh = openSheetStrict_();
  let currentRowCount = MEM_STORE.lastRowCount;
  
  // Only call getLastRow if TTL expired
  if (now - MEM_STORE.lastRowCheck > MEM_STORE.lastRowTtl) {
    currentRowCount = sh.getLastRow();
    MEM_STORE.lastRowCount = currentRowCount;
    MEM_STORE.lastRowCheck = now;
  }

  if (MEM_STORE.body && (now - MEM_STORE.loadedAt) < MEM_STORE.dataTtlMs && MEM_STORE.lastRowCount === currentRowCount) {
    return { header: MEM_STORE.header, body: MEM_STORE.body };
  }
  const values = sh.getRange(1, 1, currentRowCount, sh.getLastColumn()).getValues();
  const header = values[0], body = values.slice(1);
  for (let i = 0; i < body.length; i++) {
    body[i][COL.MOB1] = cleanMobile_(body[i][COL.MOB1]);
    body[i][COL.ALT]  = cleanMobile_(body[i][COL.ALT]);
  }
  const idxs = buildIndexes_(body);
  Object.assign(MEM_STORE, { header, body, loadedAt: now, lastRowCount: currentRowCount, ...idxs, options: null });
  return { header, body };
}

/*************** OPTIONS ***************/
function getFilterOptionsFast() {
  try {
    const now = Date.now();
    if (MEM_STORE.options && (now - MEM_STORE.optLoadedAt) < MEM_STORE.optTtlMs) return MEM_STORE.options;
    const { body } = getSheetStore_();
    const sets = { institutes: new Set(), departments: new Set(), batches: new Set(), trainings: new Set(), campuses: new Set(), tpins: new Set() };
    for (const r of body) {
      if (r[COL.INST]) sets.institutes.add(String(r[COL.INST]).trim());
      if (r[COL.DEPT]) sets.departments.add(String(r[COL.DEPT]).trim());
      if (r[COL.BATCH]) sets.batches.add(String(r[COL.BATCH]).trim());
      sets.trainings.add(String(r[COL.TRAIN] || BLANK_LABEL).trim());
      sets.campuses.add(String(r[COL.CAMPUS] || BLANK_LABEL).trim());
      sets.tpins.add(String(r[COL.TPIN] || BLANK_LABEL).trim());
    }
    const out = {
      success: true,
      institutes: [...sets.institutes].sort(),
      departments: [...sets.departments].sort(),
      batches: [...sets.batches].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})),
      trainings: [...sets.trainings].sort(),
      campuses: [...sets.campuses].sort(),
      tpins: [...sets.tpins].sort(),
      rowCount: body.length,
      allow: ALLOW,
      subjects: ALL_SUBJECT_KEYS.map(k => ({ key: k, label: k.charAt(0).toUpperCase() + k.slice(1) + '(%)' }))
    };
    MEM_STORE.options = out; MEM_STORE.optLoadedAt = now;
    return out;
  } catch (e) { return { success: false, error: e.message }; }
}

function getSheetRowCount() {
  try { return { success: true, rowCount: openSheetStrict_().getLastRow() - 1 }; } catch (e) { return { success: false, error: e.message }; }
}

/*************** FILTERING ***************/
function getFilteredDataFast(filters, page, pageSize) {
  try {
    const { body } = getSheetStore_();
    
    const normalizeList = (arr) => {
      if (!arr || !Array.isArray(arr)) return [];
      const normBlank = normalize(BLANK_LABEL);
      return arr.map(v => {
        const nv = normalize(v);
        return (nv === normBlank || v === BLANK_LABEL) ? BLANK_KEY : nv;
      }).filter(Boolean);
    };

    const nf = {
      inst: normalizeList(filters?.institute),
      dept: normalizeList(filters?.department),
      batch: normalizeList(filters?.batch),
      train: normalizeList(filters?.trainingsSelected),
      camp: normalizeList(filters?.campusesSelected),
      onlyAllowed: filters?.onlyAllowed !== false
    };

    let candidates = null;
    const intersect = (map, arr) => {
      if (!arr || !arr.length) return candidates;
      const union = new Set();
      for (const val of arr) (map.get(val) || []).forEach(i => union.add(i));
      
      if (candidates === null) return union;
      return new Set([...candidates].filter(i => union.has(i)));
    };

    candidates = intersect(MEM_STORE.instIdx, nf.inst);
    candidates = intersect(MEM_STORE.deptIdx, nf.dept);
    candidates = intersect(MEM_STORE.batchIdx, nf.batch);
    candidates = intersect(MEM_STORE.trainIdx, nf.train);
    candidates = intersect(MEM_STORE.campusIdx, nf.camp);

    const rows = [], ps = pageSize || 200, start = ((page || 1)-1) * ps, end = start + ps;
    let total = 0;

    const process = (i) => {
      const src = body[i];
      const ok = isAllowedBySubjectsDynamic_(src, filters?.subjectsSelected, filters?.subjectLogic, filters?.allowEnglish, filters?.allowOthers);
      if (nf.onlyAllowed && !ok) return;
      total++;
      if (total > start && total <= end) {
        const { keepCols } = buildKeepColsAndHeader_(filters?.subjectsSelected || ALL_SUBJECT_KEYS);
        rows.push(keepCols.map(c => c === -1 ? (ok ? 'ALLOWED' : 'NOT ALLOWED') : src[c]));
      }
    };

    if (candidates === null) {
      for (let i=0; i<body.length; i++) process(i);
    } else {
      candidates.forEach(process);
    }

    const { header } = buildKeepColsAndHeader_(filters?.subjectsSelected || ALL_SUBJECT_KEYS);
    return { 
      success: true, 
      header, 
      rows, 
      total, 
      page: page||1, 
      totalPages: Math.ceil(total/ps),
      debug: { candidatesSize: candidates ? candidates.size : 'null', filterApplied: nf }
    };
  } catch (e) { return { success: false, error: e.message }; }
}

function lookupByQuery(query) {
  try {
    const q = normalize(query);
    if (!q) return { success: true, found: false };
    const { body } = getSheetStore_();
    for (const r of body) {
      if (normalize(r[COL.TPIN]) === q || normalize(r[COL.MOB1]).slice(-10) === q.slice(-10) || normalize(r[COL.ALT]).slice(-10) === q.slice(-10)) {
        const { keepCols, header } = buildKeepColsAndHeader_(ALL_SUBJECT_KEYS);
        return { success: true, found: true, header, row: keepCols.map(c => c === -1 ? 'LOOKUP' : r[c]) };
      }
    }
    return { success: true, found: false };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function clearFastCache() {
  MEM_STORE = { loadedAt: 0, dataTtlMs: 600000, optTtlMs: 1800000, lastRowCount: 0, header: null, body: null, options: null, optLoadedAt: 0 };
  return { success: true };
}
