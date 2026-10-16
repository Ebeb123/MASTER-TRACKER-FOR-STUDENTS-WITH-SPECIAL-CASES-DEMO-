/**
 * Demo.gs — PUBLIC PORTFOLIO DEMO of the Master Tracker
 * Add this as a NEW script file in a COPY of the project.
 * ----------------------------------------------------------------
 * - Every student, adviser and case is FICTIONAL ("Sample, Juan D.").
 * - Visitors (anyone not signed in as the owner/editor) can browse,
 *   add and edit cases, change statuses and tick checklists.
 * - Locked for visitors: deleting cases, photo uploads, Settings,
 *   and sharing / view-only access management.
 * - The owner (and spreadsheet editors) keep full access.
 * - Everything resets to the sample cases every night.
 *
 * ONE-TIME SETUP (in the demo copy only):
 *  1. Run setupDemo() from the editor and approve the permissions.
 *  2. Deploy → New deployment → Web app → Execute as: Me →
 *     Who has access: Anyone → Deploy. Share the /exec link.
 *
 * Other functions you can run from the editor:
 *    resetDemoData()     put the sample cases back right now
 *    installDemoReset()  (re)create the nightly reset trigger
 */

const DEMO = Object.freeze({
  ENABLED: true,
  RESET_HOUR: 2,                       // nightly reset hour (script time zone)
  PHOTO_FOLDER_NAME: 'Master Tracker DEMO – Case Photos',
  MAX_WRITES_PER_HOUR: 400             // all visitors together (abuse guard)
});

const DEMO_LOCKED_MSG = 'Locked in the demo: deleting cases, photo uploads, Settings and sharing are turned off. You can still add and edit cases — everything resets every night.';

/* ========== WHO IS A DEMO VISITOR? ========== */

/** True for anyone who is not the owner or a spreadsheet editor. */
function isDemoVisitor_() {
  return DEMO.ENABLED && !isEditor_();
}

/** Visitors may add/edit cases (with a global hourly cap). Owners/editors always may. */
function requireEditorOrDemo_() {
  if (isEditor_()) return;
  if (!DEMO.ENABLED) throw new Error('You have view-only access to this tracker.');
  const cache = CacheService.getScriptCache();
  const key = 'demo_writes_' + Math.floor(Date.now() / 3600000);
  const n = Number(cache.get(key) || 0);
  if (n >= DEMO.MAX_WRITES_PER_HOUR) throw new Error('The demo is busy right now. Please try again later.');
  cache.put(key, String(n + 1), 3700);
}

/** Owner/editor only (deleting, photos, settings, sharing). */
function requireOwner_() {
  if (isEditor_()) return;
  throw new Error(DEMO.ENABLED ? DEMO_LOCKED_MSG : 'You have view-only access to this tracker.');
}

/** Demo photo folder (created by setupDemo, ID kept in Script Properties). */
function demoRootFolderId_() {
  return PropertiesService.getScriptProperties().getProperty('DEMO_ROOT_FOLDER_ID') || '';
}

/* ========== SETUP + NIGHTLY RESET (run from the editor) ========== */

function setupDemo() {
  demoRequireOwnerRun_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!demoRootFolderId_()) {
    const it = DriveApp.getFoldersByName(DEMO.PHOTO_FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(DEMO.PHOTO_FOLDER_NAME);
    props.setProperty('DEMO_ROOT_FOLDER_ID', folder.getId());
  }
  formatSheet_(getSheet_());
  const msg = resetDemoData();
  installDemoReset();
  Logger.log(msg);
  Logger.log('Nightly reset installed at ' + DEMO.RESET_HOUR + ':00.');
  Logger.log('NEXT: Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone → Deploy.');
}

/** Clears everything visitors changed and writes the sample cases again. Runs nightly. */
function resetDemoData(e) {
  if (!(e && e.triggerUid)) demoRequireOwnerRun_();
  const sh = getSheet_();
  withLock_(() => {
    const last = sh.getLastRow();
    if (last > 1) {
      const range = sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), CASE_FOLDER_COL));
      range.clearContent();
    }
    // Back to the default categories, statuses, checklist and titles.
    writeConfigJson_('');
    runSettingsSync_('');
    const rows = demoCases_();
    sh.getRange(2, 1, rows.length, APP_COLS).setValues(rows);
  });
  demoEmptyPhotoFolder_();
  PropertiesService.getScriptProperties().setProperty('DEMO_LAST_RESET', new Date().toISOString());
  return 'Demo reset: ' + demoCases_().length + ' fictional cases.';
}

function installDemoReset() {
  demoRequireOwnerRun_();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'resetDemoData') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('resetDemoData').timeBased().everyDays(1).atHour(DEMO.RESET_HOUR).create();
}

function demoRequireOwnerRun_() {
  const active = Session.getActiveUser().getEmail();
  const owner = Session.getEffectiveUser().getEmail();
  if (!active || active !== owner) throw new Error('Run this from the Apps Script editor while signed in as the owner.');
}

/** Trashes leftover files/folders inside the demo photo folder only. */
function demoEmptyPhotoFolder_() {
  const id = demoRootFolderId_();
  if (!id) return;
  try {
    const root = DriveApp.getFolderById(id);
    const files = root.getFiles();
    while (files.hasNext()) files.next().setTrashed(true);
    const folders = root.getFolders();
    while (folders.hasNext()) folders.next().setTrashed(true);
  } catch (err) {
    Logger.log('Demo photo folder cleanup skipped: ' + err);
  }
}

/* ========== SAMPLE CASES (all fictional) ========== */

function demoCases_() {
  const tz = Session.getScriptTimeZone();
  const day = (offset) => Utilities.formatDate(new Date(Date.now() + offset * 86400000), tz, 'yyyy-MM-dd');
  const S = {
    green: 'Green – Cleared for Face-to-Face',
    yellow: 'Yellow – Under Monitoring',
    blue: 'Blue – Alternative Learning Modality',
    red: 'Red – Restricted or Pending Documents',
    purple: 'Purple – WITHDRAWN / Case Closed'
  };
  const M = { f2f: 'F2F – Face-to-Face', alm: 'ALM – Alternative Learning Modality', hyb: 'HYB – Hybrid', tfa: 'TFA – Temporarily Not Attending' };
  // [name, section, adviser, categories, tagged (days ago), modality, approved?, notified, health decl, med clearance, status, last update (days ago), next review (days ahead), remarks]
  const list = [
    ['Sample, Juan D.', '11 - STEM Demo A', 'Ms. Adviser Uno', 'AC – Accident/Injury Recovery', 40, M.hyb, true, true, true, true, S.green, 3, 25, 'Cleared by the school clinic; back to full face-to-face classes.'],
    ['Testcase, Maria R.', '12 - ABM Demo B', 'Mr. Adviser Dos', 'MC – Medical Condition', 35, M.hyb, true, true, true, false, S.yellow, 5, 9, 'Waiting for the updated medical clearance from the family.'],
    ['Placeholder, Ken L.', '11 - HUMSS Demo C', 'Mrs. Adviser Tres', 'TP – Temporary Physical Limitation', 28, M.f2f, true, true, true, true, S.green, 2, 20, 'Excused from PE activities for four weeks.'],
    ['Example, Andrea P.', '12 - STEM Demo A', 'Ms. Adviser Uno', 'FC – Family Crisis', 26, M.alm, true, true, false, false, S.blue, 4, 7, 'Modules sent weekly; adviser checks in every Friday.'],
    ['Mockup, Paolo S.', '11 - TVL Demo D', 'Mr. Adviser Cuatro', 'MH – Mental Health', 22, M.hyb, true, true, true, false, S.yellow, 6, 8, 'Regular check-ins with the guidance office (details kept by Guidance).'],
    ['Demo, Bea T.', '12 - HUMSS Demo C', 'Mrs. Adviser Tres', 'LD – Learning/Developmental Condition', 60, M.f2f, true, true, true, true, S.green, 10, 30, 'Extended time on quizzes agreed with subject teachers.'],
    ['Sandbox, Carlo M.', '11 - ABM Demo B', 'Mr. Adviser Dos', 'MC – Medical Condition', 15, M.tfa, true, true, false, false, S.red, 2, 3, 'Pending hospital documents before returning to class.'],
    ['Lorem, Trisha N.', '12 - GAS Demo E', 'Ms. Adviser Cinco', 'AC – Accident/Injury Recovery, TP – Temporary Physical Limitation', 18, M.alm, true, true, true, false, S.blue, 3, 11, 'Arm cast; online submissions accepted until cleared.'],
    ['Ipsum, Miguel A.', '11 - STEM Demo A', 'Ms. Adviser Uno', 'FC – Family Crisis, MH – Mental Health', 12, M.hyb, true, true, false, false, S.yellow, 1, 6, 'Family meeting scheduled with the adviser and guidance office.'],
    ['Fakerson, Nicole E.', '12 - TVL Demo D', 'Mr. Adviser Cuatro', 'OT – Others', 50, M.f2f, true, true, true, true, S.purple, 20, 0, 'Case closed: transferred to another school (sample record).'],
    ['Prototype, Daniel G.', '11 - HUMSS Demo C', 'Mrs. Adviser Tres', 'MC – Medical Condition', 9, M.hyb, false, true, false, false, S.red, 1, 4, 'Waiting for the health declaration form.'],
    ['Testcase, Angela F.', '12 - ABM Demo B', 'Mr. Adviser Dos', 'TP – Temporary Physical Limitation', 7, M.f2f, true, true, true, false, S.yellow, 2, 12, 'Using the ground-floor classroom for now.'],
    ['Sample, Kevin O.', '11 - TVL Demo D', 'Mr. Adviser Cuatro', 'LD – Learning/Developmental Condition', 75, M.f2f, true, true, true, true, S.green, 14, 40, 'Seat near the board; check progress each grading period.'],
    ['Example, Sofia J.', '12 - GAS Demo E', 'Ms. Adviser Cinco', 'MH – Mental Health', 5, M.alm, false, true, false, false, S.blue, 1, 5, 'Newly tagged; adviser notified, waiting for approval.'],
    ['Mockup, Gabriel V.', '11 - STEM Demo A', 'Ms. Adviser Uno', 'AC – Accident/Injury Recovery', 3, M.tfa, false, false, false, false, S.red, 0, 2, 'Reported today by the class adviser (sample record).'],
    ['Placeholder, Erika C.', '12 - HUMSS Demo C', 'Mrs. Adviser Tres', 'OT – Others', 45, M.hyb, true, true, true, true, S.green, 8, 21, 'Temporary hybrid setup approved for a family relocation.']
  ];
  return list.map((c, i) => [
    i + 1, c[0], c[1], c[2], c[3], day(-c[4]), c[5], c[6] ? day(-c[4] + 2) : '',
    c[7], c[8], c[9], c[10], day(-c[11]), c[12] ? day(c[12]) : '', c[13], ''
  ]);
}