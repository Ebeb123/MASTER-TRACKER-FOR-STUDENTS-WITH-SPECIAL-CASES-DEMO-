/**
 * Code.gs
 * Student Special Case Master Tracker  (clean web app)
 * PUBLIC DEMO COPY: see Demo.gs (fictional cases, visitor sandbox, nightly reset).
 * ----------------------------------------------------------------
 * Files:  Code.gs · Index.html · App.html
 * Data:   tab "CASES" (row 1 = headers, one case per row)
 * Photos: one Drive folder per case inside ROOT_FOLDER (Col P = photo refs, Col Q = folder link)
 *
 * Settings (categories, modalities, statuses, KPI cards, titles, checklist) live in Script Properties
 * and are pushed into the sheet every time they are saved:
 *   - Categories  → column E: a case can have several ("MH – Mental Health, FC – Family Crisis").
 *                   Renamed categories are updated in every row. (No dropdown: Apps Script cannot make
 *                   multi-select dropdowns — pick categories in the tracker.)
 *   - Modalities  → column G dropdown (warning only, never blocks typing) + renames updated in every row
 *   - Statuses    → column L dropdown + cell colors + renames updated in every row
 *   - Checklist   → header/checkbox columns (I–K built-in, R onward for new items)
 */

// ===== SETTINGS =====
const ROOT_FOLDER = '';   // demo copy: setupDemo() creates its own photo folder
const SHEET_NAME = 'CASES';

const OLD_TRACKER_URL = '';
const OLD_TRACKER_TAB = 'MAIN TRACKER';

// ===== LAYOUT (do not change) =====
const HEADERS = [
  'No.', 'Student Name', 'Grade & Section', 'Adviser', 'Category', 'Date Tagged',
  'Learning Modality', 'Date Approved', 'Adviser Notified', 'Health Declaration Submitted',
  'Medical Clearance on File', 'Current Status', 'Last Monitoring Update', 'Next Review Date',
  'Remarks', 'Photos (app data)', 'Case Folder'
];
const APP_COLS = 16;          // A–P written by the app
const PHOTO_META_COL = 16;    // P
const CASE_FOLDER_COL = 17;   // Q
const CATEGORY_COL = 5;       // E
const MODALITY_COL = 7;       // G
const STATUS_COL = 12;        // L
const PENDING_FOLDER_NAME = '_Pending uploads (unsaved cases)';
const CONFIG_KEY = 'TRACKER_CONFIG';
const CONFIG_CHUNK = 8000;            // Script Properties allow ~9 KB per value, so big settings are split
const CONFIG_MAX = 100000;
const SYNC_ERROR_KEY = 'TRACKER_SYNC_ERROR';

// Checklist: built-in items = columns I–K (0-based 8–10), new items = column R onward (0-based 17+)
const CHECKLIST_START_COL = 17;
const CHECKLIST_MAX_COL = 60;
const DEFAULT_CHECKLIST_LABELS = { 8: 'Adviser Notified', 9: 'Health Declaration Submitted', 10: 'Medical Clearance on File' };

// Same ids as the app's defaults (used to detect renames)
const DEFAULT_CATEGORIES = [
  { id: 'cat-1', code: 'MH', full: 'MH – Mental Health' },
  { id: 'cat-2', code: 'MC', full: 'MC – Medical Condition' },
  { id: 'cat-3', code: 'AC', full: 'AC – Accident/Injury Recovery' },
  { id: 'cat-4', code: 'PG', full: 'PG – Pregnancy' },
  { id: 'cat-5', code: 'TP', full: 'TP – Temporary Physical Limitation' },
  { id: 'cat-6', code: 'FC', full: 'FC – Family Crisis' },
  { id: 'cat-7', code: 'LD', full: 'LD – Learning/Developmental Condition' },
  { id: 'cat-8', code: 'OT', full: 'OT – Others' }
];
const DEFAULT_MODALITIES = [
  { id: 'mod-1', code: 'F2F', full: 'F2F – Face-to-Face' },
  { id: 'mod-2', code: 'ALM', full: 'ALM – Alternative Learning Modality' },
  { id: 'mod-3', code: 'HYB', full: 'HYB – Hybrid' },
  { id: 'mod-4', code: 'TFA', full: 'TFA – Temporarily Not Attending' }
];
const DEFAULT_STATUSES = [
  { id: 'st-green', label: 'Green – Cleared for Face-to-Face', color: '#10B981' },
  { id: 'st-yellow', label: 'Yellow – Under Monitoring', color: '#F59E0B' },
  { id: 'st-blue', label: 'Blue – Alternative Learning Modality', color: '#3B82F6' },
  { id: 'st-red', label: 'Red – Restricted or Pending Documents', color: '#EF4444' },
  { id: 'st-purple', label: 'Purple – WITHDRAWN / Case Closed', color: '#9333EA' }
];

// ===== WEB APP =====
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Student Special Case Master Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ===== MENU =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master Tracker')
    .addItem('📋 Open Tracker', 'showTracker')
    .addItem('Open Tracker in a new tab', 'showTrackerLink')
    .addSeparator()
    .addItem('Sync settings to sheet (lists, colors & checklist)', 'syncDropdownLists')
    .addItem('Open case photos folder', 'showRootFolderLink')
    .addToUi();
}

function showTracker() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate().setWidth(1400).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Student Special Case Master Tracker');
}

function showTrackerLink() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Not deployed yet. Use "Open Tracker", or deploy: Deploy → New deployment → Web app.');
    return;
  }
  showLinkDialog_(url, 'Open Master Tracker in a new tab');
}

function showRootFolderLink() {
  showLinkDialog_(getRootFolder_().getUrl(), 'Open case photos folder');
}

function showLinkDialog_(url, label) {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;font-size:14px"><a href="' + url + '" target="_blank" onclick="google.script.host.close()">' +
      label.replace(/[<>&"]/g, '') + '</a></p>'
    ).setWidth(380).setHeight(90),
    'Master Tracker'
  );
}

// ===== ONE-TIME SETUP =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  const sh = getSheet_();
  formatSheet_(sh);
  runSettingsSync_('');

  const root = getRootFolder_();
  const me = Session.getEffectiveUser().getEmail();
  ss.getEditors().map(u => u.getEmail()).filter(e => e && e !== me).forEach(e => {
    try { root.addEditor(e); } catch (err) { Logger.log('Could not share with ' + e + ': ' + err); }
  });

  syncCaseFolders();
  Logger.log('Setup complete.\nSpreadsheet: ' + ss.getName() + '\nTab: ' + sh.getName() + '\nPhotos folder: ' + root.getUrl());
}

function formatSheet_(sh) {
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#FFF4F1').setFontColor('#523B36').setWrap(true);
  sh.setFrozenRows(1);
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  // Checkbox rule only (insertCheckboxes() writes FALSE into every row, which breaks getLastRow)
  sh.getRange(2, 9, rows, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange(2, 6, rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 8, rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 13, rows, 2).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, CATEGORY_COL, rows, 1).setWrap(true);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(CATEGORY_COL, 220);
  sh.setColumnWidth(PHOTO_META_COL, 140);
  sh.setColumnWidth(CASE_FOLDER_COL, 260);
}

function syncCaseFolders() {
  const sh = getSheet_();
  const last = lastCaseRow_(sh);
  let count = 0;
  for (let i = 1; i < last; i++) {
    if (ensureCaseFolder_(sh, i, true)) count++;
  }
  Logger.log('Case folders checked: ' + count);
}

// ===== SETTINGS → SHEET SYNC =====
// Menu item: Master Tracker → Sync settings to sheet
function syncDropdownLists() {
  const errors = runSettingsSync_('');
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      errors.length ? 'Sync problem: ' + errors.join('; ') : 'Categories, modalities, statuses (with colors) and checklist columns now match Settings.',
      'Master Tracker', 6
    );
  } catch (e) { /* not opened from the sheet */ }
}

// Pushes everything from Settings into the sheet. oldJson = settings before this save (to detect renames/deletes)
function runSettingsSync_(oldJson) {
  const errors = [];
  try { syncChecklistColumns_(oldJson); } catch (err) { errors.push('checklist columns – ' + (err.message || err)); }
  try { syncListColumns_(oldJson); } catch (err) { errors.push('categories/modalities/statuses – ' + (err.message || err)); }
  const props = PropertiesService.getScriptProperties();
  if (errors.length) {
    props.setProperty(SYNC_ERROR_KEY, errors.join('; ').slice(0, 500));
    Logger.log('Settings sync failed: ' + errors.join('; '));
  } else {
    props.deleteProperty(SYNC_ERROR_KEY);
  }
  return errors;
}

// ----- config storage (split into parts when large) -----
function readConfigJson_() {
  const props = PropertiesService.getScriptProperties();
  const parts = Number(props.getProperty(CONFIG_KEY + '_PARTS') || 0);
  if (!parts) return props.getProperty(CONFIG_KEY) || '';
  let json = '';
  for (let i = 0; i < parts; i++) json += props.getProperty(CONFIG_KEY + '_' + i) || '';
  return json;
}

function writeConfigJson_(json) {
  const props = PropertiesService.getScriptProperties();
  const oldParts = Number(props.getProperty(CONFIG_KEY + '_PARTS') || 0);
  if (json.length <= CONFIG_CHUNK) {
    props.setProperty(CONFIG_KEY, json);
    props.deleteProperty(CONFIG_KEY + '_PARTS');
    for (let i = 0; i < oldParts; i++) props.deleteProperty(CONFIG_KEY + '_' + i);
    return;
  }
  const obj = {};
  let n = 0;
  for (let i = 0; i < json.length; i += CONFIG_CHUNK) obj[CONFIG_KEY + '_' + (n++)] = json.slice(i, i + CONFIG_CHUNK);
  obj[CONFIG_KEY + '_PARTS'] = String(n);
  props.setProperties(obj);
  props.deleteProperty(CONFIG_KEY);
  for (let i = n; i < oldParts; i++) props.deleteProperty(CONFIG_KEY + '_' + i);
}

function parseConfig_(json) {
  try {
    return JSON.parse(json || '{}') || {};
  } catch (e) {
    return {};
  }
}

function readConfig_() {
  return parseConfig_(readConfigJson_());
}

// ----- list helpers -----
function normKey_(s) {
  return String(s === null || s === undefined ? '' : s)
    .toLowerCase().replace(/[–—]/g, '-').replace(/\s*-\s*/g, ' - ').replace(/\s+/g, ' ').trim();
}

function codeOf_(s) {
  return String(s || '').split(/\s*[–—-]\s*/)[0].trim().toUpperCase();
}

function optionList_(list, fallback) {
  const src = (Array.isArray(list) && list.length) ? list : fallback;
  const seen = {};
  const out = [];
  src.forEach(o => {
    const code = String((o && o.code) || '').trim();
    if (!code) return;
    const label = String((o && o.label) || '').trim();
    const full = String((o && o.full) || '').trim() || (label ? code + ' – ' + label : code);
    if (seen[full]) return;
    seen[full] = true;
    out.push({ id: String((o && o.id) || ''), code: code, full: full });
  });
  return out;
}

function statusList_(list) {
  const src = (Array.isArray(list) && list.length) ? list : DEFAULT_STATUSES;
  const seen = {};
  const out = [];
  src.forEach(s => {
    const full = String((s && s.label) || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, 100);
    if (!full || seen[normKey_(full)]) return;
    seen[normKey_(full)] = true;
    const color = /^#[0-9a-f]{6}$/i.test(String(s.color || '')) ? s.color : '#8A736E';
    out.push({ id: String(s.id || ''), full: full, color: color });
  });
  return out;
}

function listsFrom_(cfg) {
  return {
    categories: optionList_(cfg.categoriesList, DEFAULT_CATEGORIES),
    modalities: optionList_(cfg.modalitiesList, DEFAULT_MODALITIES),
    statuses: statusList_(cfg.statusesList)
  };
}

// old value (normalized) → new value, for items whose id stayed the same but name changed
function renameMap_(oldList, newList) {
  const map = {};
  oldList.forEach(o => {
    if (!o.id) return;
    const n = newList.filter(x => x.id === o.id)[0];
    if (n && n.full !== o.full) map[normKey_(o.full)] = n.full;
  });
  return map;
}

// Updates categories, modalities and statuses in the sheet (values, dropdowns, colors)
function syncListColumns_(oldJson) {
  withLock_(() => {
    const sh = getSheet_();
    const now = listsFrom_(readConfig_());
    const before = oldJson ? listsFrom_(parseConfig_(oldJson)) : null;
    const rows = lastCaseRow_(sh) - 1;
    const allRows = Math.max(sh.getMaxRows() - 1, 1);

    // ---- Column E: categories (several per case, comma separated) ----
    const catRenames = before ? renameMap_(before.categories, now.categories) : {};
    const catByName = {}, catByCode = {};
    now.categories.forEach(c => {
      catByName[normKey_(c.full)] = c.full;
      if (!catByCode[c.code.toUpperCase()]) catByCode[c.code.toUpperCase()] = c.full;
    });
    if (rows > 0) {
      rewriteColumn_(sh, CATEGORY_COL, rows, v => {
        const parts = String(v).split(/\s*[,;\n]\s*/).map(p => p.trim()).filter(Boolean);
        const out = [];
        parts.forEach(p => {
          const k = normKey_(p);
          const val = catRenames[k] || catByName[k] || catByCode[codeOf_(p)] || p;
          if (out.indexOf(val) === -1) out.push(val);
        });
        return out.join(', ');
      });
    }
    sh.getRange(2, CATEGORY_COL, allRows, 1).clearDataValidations();   // multi-select → no single-value dropdown

    // ---- Column G: modality ----
    const modRenames = before ? renameMap_(before.modalities, now.modalities) : {};
    const modByName = {}, modByCode = {};
    now.modalities.forEach(m => {
      modByName[normKey_(m.full)] = m.full;
      if (!modByCode[m.code.toUpperCase()]) modByCode[m.code.toUpperCase()] = m.full;
    });
    if (rows > 0) {
      rewriteColumn_(sh, MODALITY_COL, rows, v => {
        const k = normKey_(v);
        return modRenames[k] || modByName[k] || modByCode[codeOf_(v)] || v;
      });
    }
    setListRule_(sh, MODALITY_COL, now.modalities.map(m => m.full),
      'Pick a modality from the list (manage them in the tracker: Settings → Modalities).');

    // ---- Column L: status (+ colors) ----
    const stRenames = before ? renameMap_(before.statuses, now.statuses) : {};
    const stByName = {};
    now.statuses.forEach(s => { stByName[normKey_(s.full)] = s.full; });
    if (rows > 0) {
      rewriteColumn_(sh, STATUS_COL, rows, v => {
        const k = normKey_(v);
        return stRenames[k] || stByName[k] || v;
      });
    }
    setListRule_(sh, STATUS_COL, now.statuses.map(s => s.full),
      'Pick a status from the list (manage them in the tracker: Settings → Statuses & Colors).');
    applyStatusColors_(sh, now.statuses);
  });
}

// Rewrites non-empty cells of one column (rows 2..rows+1) with fn(value); only writes if something changed
function rewriteColumn_(sh, col, rows, fn) {
  const range = sh.getRange(2, col, rows, 1);
  const vals = range.getValues();
  let changed = false;
  const out = vals.map(r => {
    const v = r[0];
    if (v === '' || v === null || v === undefined) return r;
    const nv = fn(String(v).trim());
    if (nv !== v) { changed = true; return [nv]; }
    return r;
  });
  if (changed) range.setValues(out);
}

// Dropdown in WARNING mode (invalid values get a red corner, never a "There was a problem" popup)
function setListRule_(sh, col, values, helpText) {
  const rows = sh.getMaxRows() - 1;
  if (rows < 1 || !values.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .setHelpText(helpText)
    .build();
  sh.getRange(2, col, rows, 1).setDataValidation(rule);
}

// Colors each status cell in column L with the status color (replaces only the column-L rules)
function applyStatusColors_(sh, statuses) {
  const rows = sh.getMaxRows() - 1;
  if (rows < 1) return;
  const range = sh.getRange(2, STATUS_COL, rows, 1);
  const others = sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().every(r => r.getColumn() === STATUS_COL && r.getNumColumns() === 1)
  );
  const mine = statuses.map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(s.full)
    .setBackground(mixHex_(s.color, 255, 0.8))
    .setFontColor(mixHex_(s.color, 0, 0.45))
    .setBold(true)
    .setRanges([range])
    .build());
  sh.setConditionalFormatRules(others.concat(mine));
}

function mixHex_(hex, target, amt) {
  const n = parseInt(String(hex).replace('#', ''), 16);
  const rgb = isNaN(n) ? [168, 146, 140] : [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  return '#' + rgb.map(c => Math.round(c + (target - c) * amt).toString(16).padStart(2, '0')).join('');
}

// ===== CHECKLIST ITEMS → SHEET COLUMNS =====
function normalizeChecklist_(list) {
  const src = Array.isArray(list) ? list : [];
  const out = [];
  [8, 9, 10].forEach(col => {
    const found = src.filter(x => x && Number(x.col) === col)[0];
    const label = cleanLabel_(found && found.label) || DEFAULT_CHECKLIST_LABELS[col];
    out.push({ id: 'builtin-' + col, col: col, label: label, builtin: true });
  });
  const usedCols = {};
  src.forEach(x => {
    if (!x || !x.id) return;
    const col = Number(x.col);
    const label = cleanLabel_(x.label);
    if (!label || !(col >= CHECKLIST_START_COL && col <= CHECKLIST_MAX_COL) || usedCols[col]) return;
    usedCols[col] = true;
    out.push({ id: String(x.id), col: col, label: label, builtin: false });
  });
  return out;
}

function cleanLabel_(s) {
  return String(s || '').trim().replace(/^[=+\-@]+/, '').trim().slice(0, 80);
}

function getChecklistExtras_() {
  return normalizeChecklist_(readConfig_().checklistList).filter(x => !x.builtin);
}

function syncChecklistColumns_(oldJson) {
  withLock_(() => {
    const sh = getSheet_();
    const oldExtras = oldJson
      ? normalizeChecklist_(parseConfig_(oldJson).checklistList).filter(x => !x.builtin)
      : [];
    const newList = normalizeChecklist_(readConfig_().checklistList);
    const newIds = {};
    newList.forEach(x => { newIds[x.id] = true; });

    oldExtras.forEach(x => {
      if (!newIds[x.id]) clearChecklistColumn_(sh, x.col);   // deleted item → erase its column
    });
    newList.forEach(x => applyChecklistColumn_(sh, x));     // headers + checkboxes
  });
}

function applyChecklistColumn_(sh, item) {
  const c = item.col + 1;
  if (c > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
  sh.getRange(1, c).setValue(item.label)
    .setFontWeight('bold').setBackground('#FFF4F1').setFontColor('#523B36').setWrap(true);
  if (!item.builtin) {
    const rows = sh.getMaxRows() - 1;
    if (rows > 0) {
      sh.getRange(2, c, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
    sh.setColumnWidth(c, 130);
  }
}

function clearChecklistColumn_(sh, col) {
  const c = col + 1;
  if (col < CHECKLIST_START_COL || c > sh.getMaxColumns()) return;
  const range = sh.getRange(1, c, sh.getMaxRows(), 1);
  range.clear();
  range.clearDataValidations();
}

function writeExtras_(sh, index, row) {
  if (!Array.isArray(row)) return;
  getChecklistExtras_().forEach(x => {
    if (x.col + 1 > sh.getMaxColumns()) return;
    const v = row[x.col];
    if (v === undefined || v === null) return;
    sh.getRange(index + 1, x.col + 1).setValue(v === true || String(v).toLowerCase() === 'true');
  });
}

// ===== HELPERS =====
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME, 0);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 1-based number of the last row that has a student name (1 = only the header)
function lastCaseRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const names = sh.getRange(2, 2, last - 1, 1).getValues();
  for (let i = names.length - 1; i >= 0; i--) {
    if (String(names[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

function parseFolderId_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  const m = s.match(/folders\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return /^[a-zA-Z0-9_-]{15,}$/.test(s) ? s : '';
}

function rootFolderId_() {
  return parseFolderId_(ROOT_FOLDER) || (typeof demoRootFolderId_ === 'function' ? demoRootFolderId_() : '');
}

function getRootFolder_() {
  const id = rootFolderId_();
  if (!id) throw new Error('ROOT_FOLDER is empty. Paste your Drive folder link at the top of Code.gs.');
  try {
    const f = DriveApp.getFolderById(id);
    f.getName();
    return f;
  } catch (e) {
    throw new Error('No access to the case photos folder. Ask the owner to share it with you.');
  }
}

function getPendingFolder_() {
  const root = getRootFolder_();
  const it = root.getFoldersByName(PENDING_FOLDER_NAME);
  return it.hasNext() ? it.next() : root.createFolder(PENDING_FOLDER_NAME);
}

function isInTracker_(item) {
  const rootId = rootFolderId_();
  let level = [item];
  for (let depth = 0; depth < 3; depth++) {
    const next = [];
    for (const node of level) {
      const parents = node.getParents();
      while (parents.hasNext()) {
        const p = parents.next();
        if (p.getId() === rootId) return true;
        next.push(p);
      }
    }
    if (!next.length) return false;
    level = next;
  }
  return false;
}

function isInFolder_(file, folder) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folder.getId()) return true;
  }
  return false;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function prepareRow_(row) {
  return (row || []).map(v => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') {
      if (v.length > 49000) v = v.slice(0, 49000);
      if (/^[=+@]/.test(v) || /^-[^\d]/.test(v)) v = "'" + v;
    }
    return v;
  });
}

function readPhotoIds_(metaCell) {
  if (!metaCell || typeof metaCell !== 'string') return [];
  try {
    const list = JSON.parse(metaCell);
    return Array.isArray(list) ? list.map(p => p && p.driveId).filter(Boolean) : [];
  } catch (e) {
    return [];
  }
}

function trashFiles_(ids) {
  (ids || []).forEach(id => {
    try {
      const f = DriveApp.getFileById(id);
      if (isInTracker_(f)) f.setTrashed(true);
    } catch (e) { /* already gone */ }
  });
}

function writeRow_(sh, index, row) {
  const vals = prepareRow_(row).slice(0, APP_COLS);
  while (vals.length < APP_COLS) vals.push('');
  sh.getRange(index + 1, 1, 1, APP_COLS).setValues([vals]);
}

function isCaseRow_(row) {
  return Boolean(row) && String(row[0]).trim() !== '__TRACKER_CONFIG__' && String(row[1] || '').trim() !== '';
}

function caseFolderName_(row) {
  const no = row[0];
  const num = (no !== '' && !isNaN(Number(no))) ? String(Number(no)).padStart(3, '0') : String(no || '').trim();
  const name = String(row[1] || '').trim();
  const section = String(row[2] || '').trim();
  const label = (num ? num + ' – ' : '') + name + (section ? ' (' + section + ')' : '');
  return label.replace(/[\\\/:*?"<>|]/g, '-').slice(0, 120);
}

function ensureCaseFolder_(sh, index, movePhotos) {
  if (index < 1) return null;
  const row = sh.getRange(index + 1, 1, 1, CASE_FOLDER_COL).getValues()[0];
  if (!isCaseRow_(row)) return null;

  let folder = null;
  const existingId = parseFolderId_(row[CASE_FOLDER_COL - 1]);
  if (existingId) {
    try {
      folder = DriveApp.getFolderById(existingId);
      if (folder.isTrashed()) folder = null;
    } catch (e) {
      folder = null;
    }
  }

  const wantName = caseFolderName_(row);
  if (!folder) {
    const same = getRootFolder_().getFoldersByName(wantName);
    folder = same.hasNext() ? same.next() : getRootFolder_().createFolder(wantName);
    sh.getRange(index + 1, CASE_FOLDER_COL).setValue(folder.getUrl());
    movePhotos = true;
  }
  if (folder.getName() !== wantName) folder.setName(wantName);

  if (movePhotos) {
    readPhotoIds_(row[PHOTO_META_COL - 1]).forEach(id => {
      try {
        const f = DriveApp.getFileById(id);
        if (!isInFolder_(f, folder)) f.moveTo(folder);
      } catch (e) { /* skip */ }
    });
  }
  return folder;
}

// ===== SETTINGS STORAGE =====
function getConfigRow_() {
  const json = readConfigJson_();
  if (!json) return null;
  return {
    index_: -1,
    row: ['__TRACKER_CONFIG__', json],
    syncError: PropertiesService.getScriptProperties().getProperty(SYNC_ERROR_KEY) || ''
  };
}

function saveConfigRow_(row) {
  const json = String((row && row[1]) || '');
  if (json.length > CONFIG_MAX) throw new Error('Settings are too large to save. Remove some categories, statuses or checklist items.');
  const oldJson = readConfigJson_();
  writeConfigJson_(json);
  runSettingsSync_(oldJson);   // any failure is stored and shown in the app
}

function isConfigRow_(row) {
  return Array.isArray(row) && String(row[0]).trim() === '__TRACKER_CONFIG__';
}

// ===== DATA API =====
function getData() {
  const sh = getSheet_();
  const lastRow = lastCaseRow_(sh);
  const tz = getSpreadsheet_().getSpreadsheetTimeZone();
  const out = [];
  if (sh.getLastRow() > 0) {
    const cols = Math.max(CASE_FOLDER_COL, sh.getLastColumn());
    const values = sh.getRange(1, 1, lastRow, cols).getValues();
    values.forEach((row, i) => out.push({
      index_: i,
      row: row.map(c => (c instanceof Date ? Utilities.formatDate(c, tz, 'yyyy-MM-dd') : c))
    }));
  }
  const cfg = getConfigRow_();
  if (cfg) out.push(cfg);
  return out;
}

function updateRow(index, row) {
  if (isConfigRow_(row)) {
    requireOwner_();                      // Settings: owner/editors only
    saveConfigRow_(row);
    return getData();
  }
  requireEditorOrDemo_();                 // demo visitors may edit cases
  if (index < 1) throw new Error('The header row cannot be edited.');
  const visitor = isDemoVisitor_();
  withLock_(() => {
    const sh = getSheet_();
    const oldIds = readPhotoIds_(sh.getRange(index + 1, PHOTO_META_COL).getValue());
    writeRow_(sh, index, row);
    writeExtras_(sh, index, row);
    const newIds = readPhotoIds_(row && row[PHOTO_META_COL - 1]);
    if (visitor) return;                  // demo visitors: no Drive changes
    trashFiles_(oldIds.filter(id => newIds.indexOf(id) === -1));
    ensureCaseFolder_(sh, index, newIds.some(id => oldIds.indexOf(id) === -1));
  });
  return getData();
}

function insertRow(index, row) {
  if (isConfigRow_(row)) {
    requireOwner_();                      // Settings: owner/editors only
    saveConfigRow_(row);
    return getData();
  }
  requireEditorOrDemo_();                 // demo visitors may add cases
  const visitor = isDemoVisitor_();
  withLock_(() => {
    const sh = getSheet_();
    const target = lastCaseRow_(sh);   // 0-based index of the row right after the last case
    if (visitor && target > 300) throw new Error('The demo is full for today. It resets every night.');
    writeRow_(sh, target, row);
    writeExtras_(sh, target, row);
    if (!visitor) ensureCaseFolder_(sh, target, true);
  });
  return getData();
}

function deleteCase(index) {
  requireOwner_();
  if (index < 1) throw new Error('The header row cannot be deleted.');
  withLock_(() => {
    const sh = getSheet_();
    const row = sh.getRange(index + 1, 1, 1, CASE_FOLDER_COL).getValues()[0];
    const folderId = parseFolderId_(row[CASE_FOLDER_COL - 1]);
    let trashed = false;
    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        if (isInTracker_(folder)) {
          folder.setTrashed(true);
          trashed = true;
        }
      } catch (e) { /* already gone */ }
    }
    if (!trashed) trashFiles_(readPhotoIds_(row[PHOTO_META_COL - 1]));
    sh.deleteRow(index + 1);
  });
  return getData();
}

// ===== ACCESS: EDITORS vs VIEW-ONLY =====
const VIEWERS_KEY = 'VIEW_ONLY_USERS';

function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

function isEditor_() {
  const email = currentEmail_();
  if (!email) return false;
  const cache = CacheService.getUserCache();
  const cached = cache.get('role:' + email);
  if (cached) return cached === 'editor';
  let editor = false;
  try {
    const ss = getSpreadsheet_();
    const owner = ss.getOwner();
    const emails = ss.getEditors().map(u => String(u.getEmail()).toLowerCase());
    if (owner) emails.push(String(owner.getEmail()).toLowerCase());
    editor = emails.indexOf(email) !== -1;
  } catch (e) {
    editor = false;
  }
  cache.put('role:' + email, editor ? 'editor' : 'viewer', 300);
  return editor;
}

function requireEditor_() {
  if (!isEditor_()) throw new Error('You have view-only access to this tracker.');
}

function listViewers_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(VIEWERS_KEY) || '[]');
  } catch (e) {
    return [];
  }
}

function saveViewers_(list) {
  PropertiesService.getScriptProperties().setProperty(VIEWERS_KEY, JSON.stringify(list));
}

function getAccessInfo() {
  const editor = isEditor_();
  const demo = typeof DEMO !== 'undefined' && DEMO.ENABLED;
  return {
    // Demo visitors get the editing screens (sandbox); the server still locks deleting, photos, Settings and sharing.
    role: (editor || demo) ? 'editor' : 'viewer',
    demo: demo && !editor,
    email: editor ? currentEmail_() : '',
    webAppUrl: ScriptApp.getService().getUrl() || '',
    viewers: editor ? listViewers_() : []
  };
}

function addViewer(email) {
  requireOwner_();
  email = String(email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Please enter a valid email address.');
  const ss = getSpreadsheet_();
  const editors = ss.getEditors().map(u => String(u.getEmail()).toLowerCase());
  if (editors.indexOf(email) !== -1) throw new Error(email + ' already has edit access. Remove them as an editor in the spreadsheet\'s Share settings first.');
  ss.addViewer(email);
  getRootFolder_().addViewer(email);
  const list = listViewers_().filter(v => v.email !== email);
  list.push({ email: email, added: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') });
  saveViewers_(list);
  return list;
}

function removeViewer(email) {
  requireOwner_();
  email = String(email || '').trim().toLowerCase();
  try { getSpreadsheet_().removeViewer(email); } catch (e) { /* already removed */ }
  try { getRootFolder_().removeViewer(email); } catch (e) { /* already removed */ }
  const list = listViewers_().filter(v => v.email !== email);
  saveViewers_(list);
  return list;
}

// ===== PHOTOS =====
function uploadPhoto(payload) {
  requireOwner_();
  const folder = withLock_(() => {
    let f = null;
    if (typeof payload.caseIndex === 'number') f = ensureCaseFolder_(getSheet_(), payload.caseIndex, false);
    return f || getPendingFolder_();
  });
  const m = /^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/.exec(payload.dataUrl || '');
  if (!m) throw new Error('Invalid image data.');
  const name = String(payload.name || 'photo.jpg').replace(/[\\\/:*?"<>|]/g, '-').slice(0, 120);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  return {
    id: file.getId(),
    driveId: file.getId(),
    name: name,
    size: Number(payload.size) || file.getSize(),
    type: m[1],
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function getPhotoData(fileId, thumb) {
  const file = DriveApp.getFileById(fileId);
  if (!isInTracker_(file)) throw new Error('Not a tracker photo.');
  let blob = null;
  if (thumb) {
    try { blob = file.getThumbnail(); } catch (e) { blob = null; }
  }
  if (!blob) blob = file.getBlob();
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// ===== OPTIONAL: import cases (A–O) from the old tracker =====
function importOldCases() {
  if (!OLD_TRACKER_URL) throw new Error('Fill in OLD_TRACKER_URL at the top of Code.gs first.');
  const old = SpreadsheetApp.openByUrl(OLD_TRACKER_URL).getSheetByName(OLD_TRACKER_TAB);
  if (!old) throw new Error('Tab "' + OLD_TRACKER_TAB + '" not found in the old spreadsheet.');
  const vals = old.getRange(1, 1, old.getLastRow(), 15).getValues();

  let header = -1, legend = vals.length;
  for (let i = 0; i < vals.length; i++) {
    if (header < 0) {
      if (vals[i].some(c => typeof c === 'string' && c.toLowerCase().indexOf('student name') !== -1)) header = i;
    } else if (vals[i].some(c => typeof c === 'string' && c.trim() === 'Legends')) {
      legend = i;
      break;
    }
  }

  const sh = getSheet_();
  const lastCase = lastCaseRow_(sh);
  const existing = lastCase > 1
    ? sh.getRange(2, 2, lastCase - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase())
    : [];
  const rows = vals.slice(header + 1, legend)
    .filter(r => isCaseRow_(r) && existing.indexOf(String(r[1]).trim().toLowerCase()) === -1)
    .map(r => r.concat(['']));

  if (rows.length) sh.getRange(lastCase + 1, 1, rows.length, APP_COLS).setValues(rows);
  Logger.log('Imported ' + rows.length + ' case(s).');
  syncListColumns_('');   // converts imported "CC - …" style values to the current labels
  syncCaseFolders();
}