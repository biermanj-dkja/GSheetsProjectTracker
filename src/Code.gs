/**
 * =============================================================================
 *  GSheetsProjectTracker
 *  Google Sheets Lightweight Project Tracking AppScript
 * =============================================================================
 *
 *  A container-bound Google Apps Script that turns a plain Google Sheet into a
 *  lightweight personal project tracker. The goal is to add automation WITHOUT
 *  adding anything new for the human to learn: you keep typing into the same
 *  cells you always did, and the script handles the housekeeping.
 *
 *  HOW THE SHEET IS ORGANISED (tab names are configurable in CFG below)
 *  ---------------------------------------------------------------------------
 *   STATUS      The front page. One row per OPEN project.
 *                 A = project name (you type it)
 *                 B = priority, 1 = most important (you type it)
 *                 C:F = newest log entry for that project (formula, automatic)
 *
 *   In Process  The running log. One row per update you make.
 *                 D = project (dropdown), E = item / note, F = action / status
 *                 word, G = date (auto-filled), H = link (optional)
 *                 Columns A:C are left free for your own use.
 *
 *   COMPLETED   The archive for the current year.
 *                 A = project name, B = priority (written by the script)
 *                 C:F = final log entry for that project (formula)
 *                 H:L = every log entry of every closed project (moved here)
 *
 *   Stats       Created by setup. Summary numbers plus a per-project table,
 *               all live formulas. Hidden column M feeds the dropdown.
 *
 *   Grad year List   OPTIONAL. If a tab with this name exists, setup turns it
 *                    into a self-updating "graduation year -> grade" table
 *                    (handy for K-12 schools; harmless if you don't have one).
 *
 *  DAILY USE
 *  ---------------------------------------------------------------------------
 *   1. Add a project on STATUS (name in A, priority in B). The list re-sorts by
 *      priority automatically.
 *   2. Log updates on In Process. Pick the project from the dropdown; the date
 *      fills itself in.
 *   3. When a project is finished, log an entry whose Action is "complete",
 *      then choose  Project Tools > Close completed projects.
 *
 *  INSTALL
 *  ---------------------------------------------------------------------------
 *   Extensions > Apps Script, paste this file into Code.gs, Save, reload the
 *   spreadsheet, then run  Project Tools > One-time setup / repair.
 *   See README.md for full instructions.
 *
 *  License: MIT (see LICENSE).
 * =============================================================================
 */


/* =============================================================================
 *  CONFIGURATION
 *  Everything you might reasonably want to change lives in this one object.
 * ========================================================================== */

const CFG = {
  // ---- Tab names ------------------------------------------------------------
  STATUS: 'STATUS',          // front page of open projects
  LOG: 'In Process',         // running log of updates
  DONE: 'COMPLETED',         // archive of closed projects
  STATS: 'Stats',            // auto-built statistics tab
  GRAD: 'Grad year List',    // optional graduation-year helper tab

  // ---- Layout ---------------------------------------------------------------
  // First project row on STATUS. Row 1 is the header row.
  STATUS_FIRST_ROW: 2,

  // ---- Staleness colours on STATUS column E (the "last update" date) --------
  YELLOW_DAYS: 7,            // no update for this many days  -> yellow
  RED_DAYS: 14,              // no update for this many days  -> red
  HOLD_WORD: 'hold',         // Action containing this word   -> grey, never yellow/red

  // ---- Behaviour toggles ----------------------------------------------------
  AUTO_DATE: true,           // fill in today's date when a new log row is started
  AUTO_SORT_STATUS: true,    // re-sort STATUS by priority when a priority is edited

  // Month the optional grad-year list rolls forward (7 = July).
  NEW_YEAR_MONTH: 7,

  // An Action that STARTS with one of these words marks a project as finished.
  // Case-insensitive. "cancel+ed" matches both "canceled" and "cancelled".
  CLOSE_WORDS: /^(complete|completed|done|closed|cancel+ed)\b/i,
};


/* =============================================================================
 *  COLUMN MAP
 *  Sheet column numbers (A = 1, B = 2, ...). Change these only if you move
 *  columns around; the formulas built by setup assume this layout.
 * ========================================================================== */

// In Process: the log occupies D..H (5 columns).
const L_PROJECT = 4;   // D  project name
const L_ITEM    = 5;   // E  item / note
const L_ACTION  = 6;   // F  action / status word
const L_DATE    = 7;   // G  date
const L_LINK    = 8;   // H  link
const L_WIDTH   = 5;   // number of log columns (D..H)

// COMPLETED: project list in A:B, archived log rows in H..L.
const D_NAME = 1;      // A  completed project name (B = its priority)
const D_LOG  = 8;      // H  first column of the archived log block


/* =============================================================================
 *  MENU
 *  onOpen() is a "simple trigger": Google runs it automatically every time the
 *  spreadsheet is opened, so the custom menu is always there.
 * ========================================================================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Project Tools')
    .addItem('Close completed projects', 'closeCompletedProjects')
    .addItem('Sort log by project & date', 'sortLog')
    .addItem('Sort open projects by priority', 'sortStatus')
    .addSeparator()
    .addItem('Start new year…', 'startNewYear')
    .addSeparator()
    .addItem('One-time setup / repair', 'setupOrRepair')
    .addToUi();
}


/* =============================================================================
 *  EDIT HANDLER
 *  onEdit(e) is another simple trigger: it runs after every manual edit.
 *  It must be quick and must not show dialogs. It does two small jobs:
 *    1. On STATUS, when a priority (column B) changes, re-sort by priority.
 *    2. On In Process, when you start a new log row, fill in today's date.
 * ========================================================================== */

function onEdit(e) {
  // Guard: e is missing if someone runs onEdit by hand from the editor.
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();

  // ---- Job 1: keep STATUS in priority order --------------------------------
  if (sh.getName() === CFG.STATUS) {
    const touchedPriority = firstCol <= 2 && lastCol >= 2;          // edit includes column B
    const inDataRows = e.range.getLastRow() >= CFG.STATUS_FIRST_ROW; // not just the header
    if (CFG.AUTO_SORT_STATUS && touchedPriority && inDataRows) sortStatus();
    return;
  }

  // ---- Job 2: auto-date new log rows ---------------------------------------
  if (!CFG.AUTO_DATE || sh.getName() !== CFG.LOG) return;

  // Only react to edits in the Project / Item / Action columns (D..F).
  if (lastCol < L_PROJECT || firstCol > L_ACTION) return;

  // Row range of the edit, never including the header row.
  const r1 = Math.max(2, e.range.getRow());
  const r2 = e.range.getLastRow();
  // Skip giant pastes; nobody starts 200 new entries by hand.
  if (r2 < r1 || r2 - r1 > 200) return;

  // Read Project..Date (D..G) for the edited rows in one call.
  const width = L_DATE - L_PROJECT + 1;          // 4 columns: D, E, F, G
  const dateIdx = L_DATE - L_PROJECT;            // index of G inside that block
  const block = sh.getRange(r1, L_PROJECT, r2 - r1 + 1, width).getValues();

  // Today's date in the SPREADSHEET's time zone (File > Settings), formatted as
  // an ISO string. Sheets parses it as a real date value when written.
  const today = Utilities.formatDate(new Date(), e.source.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

  // For each row: if it has a project but no date yet, stamp today.
  let changed = false;
  const dates = block.map(row => {
    const hasProject = row[0] !== '';
    const hasDate = row[dateIdx] !== '';
    if (hasProject && !hasDate) { changed = true; return [today]; }
    return [row[dateIdx]];                       // leave existing dates alone
  });

  // Write the whole date column back in one call, only if something changed.
  if (changed) sh.getRange(r1, L_DATE, dates.length, 1).setValues(dates);
}


/* =============================================================================
 *  CLOSE COMPLETED PROJECTS
 *  Replaces the manual routine of: sort the log, cut a project's rows, paste
 *  them into COMPLETED, remove the project from STATUS, delete blank cells.
 * ========================================================================== */

/** Menu entry point. Tells you if there is nothing to close. */
function closeCompletedProjects() {
  closeCompleted_(true);
}

/**
 * Finds every open project whose NEWEST log entry has a closing Action
 * (see CFG.CLOSE_WORDS), asks for confirmation, then archives it.
 *
 * @param {boolean} announceNothing  Show a message when nothing qualifies.
 * @return {number} How many projects were closed.
 */
function closeCompleted_(announceNothing) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const status = ss.getSheetByName(CFG.STATUS);
  const log = ss.getSheetByName(CFG.LOG);
  const done = ss.getSheetByName(CFG.DONE);

  // ---- Read open projects from STATUS as [name, priority] pairs -------------
  const sLast = lastDataRow_(status, 1);
  const sCount = Math.max(0, sLast - CFG.STATUS_FIRST_ROW + 1);
  const sRange = sCount ? status.getRange(CFG.STATUS_FIRST_ROW, 1, sCount, 2) : null;
  const open = sRange ? sRange.getValues().filter(r => String(r[0]).trim() !== '') : [];

  // ---- Sort the log so each project's rows sit together, oldest first ------
  // This makes every project a contiguous block, which lets us copy and delete
  // whole blocks at once instead of row by row.
  sortLog();
  const lLast = lastDataRow_(log, L_PROJECT);
  const logVals = lLast >= 2 ? log.getRange(2, L_PROJECT, lLast - 1, L_WIDTH).getValues() : [];
  // logVals[i] = [project, item, action, date, link] for sheet row i + 2.

  // ---- Find the newest entry for each project -------------------------------
  // "Newest" means latest date; if two entries share a date, the lower row
  // (entered later) wins because of the >= comparison.
  const latest = {};                                   // key -> { t, action }
  logVals.forEach(r => {
    const k = key_(r[0]);
    if (!k) return;                                    // blank project cell
    const t = r[3] instanceof Date ? r[3].getTime() : 0;
    if (!latest[k] || t >= latest[k].t) latest[k] = { t: t, action: String(r[2]).trim() };
  });

  // ---- Which open projects are ready to close? ------------------------------
  const toClose = open.filter(r => {
    const last = latest[key_(r[0])];
    return last && CFG.CLOSE_WORDS.test(last.action);
  });

  if (!toClose.length) {
    if (announceNothing) {
      ui.alert('Nothing to close.\n\nA project closes when its newest entry on "' +
        CFG.LOG + '" has an Action of "complete".');
    }
    return 0;
  }

  // ---- Confirm with the user before moving anything -------------------------
  const answer = ui.alert('Close these projects?',
    toClose.map(r => '• ' + r[0]).join('\n') +
    '\n\nTheir log entries move to ' + CFG.DONE + ' and they come off ' + CFG.STATUS + '.',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return 0;

  const closeKeys = new Set(toClose.map(r => key_(r[0])));

  // ---- Collect the sheet rows to move, grouped into contiguous runs ---------
  // e.g. rows [5,6,7,20,21] become runs [[5,3],[20,2]]  ([startRow, count]).
  const rows = [];
  logVals.forEach((r, i) => { if (closeKeys.has(key_(r[0]))) rows.push(i + 2); });
  const runs = toRuns_(rows);

  // ---- Step 1: copy the log rows into the COMPLETED archive (H..L) ----------
  // copyTo() keeps formatting and hyperlinks, unlike getValues()/setValues().
  let dest = Math.max(2, lastDataRow_(done, D_LOG) + 1);
  ensureRows_(done, dest + rows.length);
  runs.forEach(([start, n]) => {
    log.getRange(start, L_PROJECT, n, L_WIDTH).copyTo(done.getRange(dest, D_LOG, n, L_WIDTH));
    dest += n;
  });

  // ---- Step 2: add the project names + priorities to COMPLETED A:B ----------
  const nameDest = Math.max(2, lastDataRow_(done, D_NAME) + 1);
  ensureRows_(done, nameDest + toClose.length);
  done.getRange(nameDest, D_NAME, toClose.length, 2).setValues(toClose.map(r => [r[0], r[1]]));

  // ---- Step 3: remove the moved rows from In Process ------------------------
  // Delete bottom-up so earlier row numbers stay valid. Only columns D..H are
  // shifted up; anything you keep in columns A:C is untouched.
  runs.slice().reverse().forEach(([start, n]) =>
    log.getRange(start, L_PROJECT, n, L_WIDTH).deleteCells(SpreadsheetApp.Dimension.ROWS));

  // ---- Step 4: rewrite STATUS A:B without the closed projects ---------------
  // Blank rows disappear too, and the list is re-sorted by priority. This is
  // safe because STATUS C:F is ONE formula keyed on the names, not per-row.
  const remaining = open.filter(r => !closeKeys.has(key_(r[0]))).sort(byPriority_);
  if (sRange) sRange.clearContent();
  if (remaining.length) {
    status.getRange(CFG.STATUS_FIRST_ROW, 1, remaining.length, 2).setValues(remaining);
  }

  // ---- Step 5: report, and flag log entries that match no project -----------
  // These "orphans" are usually typos or renamed projects. They never show on
  // STATUS, and would be left behind when the real project closes.
  const known = new Set(remaining.map(r => key_(r[0])));
  const orphans = {};                                  // name -> entry count
  logVals.forEach(r => {
    const k = key_(r[0]);
    if (k && !known.has(k) && !closeKeys.has(k)) orphans[r[0]] = (orphans[r[0]] || 0) + 1;
  });

  let msg = 'Closed ' + toClose.length + ' project' + (toClose.length > 1 ? 's' : '') + '.';
  const names = Object.keys(orphans);
  if (names.length) {
    msg += '\n\nHeads up: these entries on "' + CFG.LOG + '" don\'t match any open project ' +
      '(typo or renamed?):\n' + names.map(n => '• ' + n + ' (' + orphans[n] + ')').join('\n');
  }
  ui.alert(msg);
  return toClose.length;
}


/* =============================================================================
 *  SORTING
 * ========================================================================== */

/**
 * Sorts STATUS A:B by priority (1 first, blanks last, ties by name) and removes
 * blank rows. Only names and priorities move; columns C:F follow automatically
 * because they are a single lookup formula keyed on the name.
 */
function sortStatus() {
  const status = SpreadsheetApp.getActive().getSheetByName(CFG.STATUS);
  const fr = CFG.STATUS_FIRST_ROW;
  const last = lastDataRow_(status, 1);
  if (last < fr) return;                                     // no projects

  const range = status.getRange(fr, 1, last - fr + 1, 2);
  const current = range.getValues();
  const sorted = current.filter(r => String(r[0]).trim() !== '').sort(byPriority_);

  // Skip the write if nothing would change (avoids needless edits/flicker).
  if (JSON.stringify(sorted) === JSON.stringify(current)) return;

  range.clearContent();
  if (sorted.length) status.getRange(fr, 1, sorted.length, 2).setValues(sorted);
}

/**
 * Comparator for [name, priority] rows: lowest priority number first, blank
 * priority last, then alphabetical by name (case-insensitive).
 */
function byPriority_(a, b) {
  const pa = a[1] === '' ? Infinity : Number(a[1]);
  const pb = b[1] === '' ? Infinity : Number(b[1]);
  if (pa !== pb) return pa - pb;
  return String(a[0]).localeCompare(String(b[0]), undefined, { sensitivity: 'base' });
}

/** Sorts the In Process log (D..H) by project, then by date (oldest first). */
function sortLog() {
  const log = SpreadsheetApp.getActive().getSheetByName(CFG.LOG);
  const last = lastDataRow_(log, L_PROJECT);
  if (last > 2) {                                            // need 2+ data rows
    log.getRange(2, L_PROJECT, last - 1, L_WIDTH)
      .sort([{ column: L_PROJECT, ascending: true }, { column: L_DATE, ascending: true }]);
  }
}


/* =============================================================================
 *  NEW YEAR ROLLOVER
 *  Makes a fresh copy of the tracker for the next year. Open projects (and
 *  their full history) carry over; the COMPLETED tab starts empty. The current
 *  file is left alone as the archive of the year just finished.
 * ========================================================================== */

function startNewYear() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // ---- Work out the new file name -------------------------------------------
  // If the name ends in a "YY-YY" pair (e.g. "Tracker 26-27"), bump both
  // numbers ("Tracker 27-28"). Otherwise just append "(new year)".
  const name = ss.getName();
  const m = name.match(/(\d{2})\s*-\s*(\d{2})(?!.*\d{2}\s*-\s*\d{2})/);   // LAST YY-YY pair
  const pad = n => ('0' + (n % 100)).slice(-2);                         // 27 -> "27", 100 -> "00"
  const newName = m ? name.replace(m[0], pad(+m[2]) + '-' + pad(+m[2] + 1)) : name + ' (new year)';

  const ok = ui.alert('Start a new year?',
    'This makes a copy called "' + newName + '" that carries over your open projects ' +
    '(with their history) and starts with an empty ' + CFG.DONE + ' tab.\n\n' +
    'This file stays as-is as the archive.', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  // Offer to close any finished projects first, so they land in THIS year.
  closeCompleted_(false);

  // ---- Copy the file and empty the archive in the copy ----------------------
  // Copying a spreadsheet also copies its bound script, so the menu, triggers
  // and formulas all work in the new file straight away.
  const copy = ss.copy(newName);
  const done = copy.getSheetByName(CFG.DONE);
  done.getRange(2, D_NAME, done.getMaxRows() - 1, 2).clearContent();       // A:B list
  done.getRange(2, D_LOG, done.getMaxRows() - 1, L_WIDTH).clearContent();  // H:L archive
  // (COMPLETED C:F is a formula, so it empties itself.)

  // ---- Show a clickable link to the new file --------------------------------
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial">New file created:</p>' +
    '<p style="font-family:Arial"><a href="' + copy.getUrl() + '" target="_blank">' +
    escapeHtml_(newName) + '</a></p>' +
    '<p style="font-family:Arial;color:#666">It was saved to your My Drive; move it wherever you like.</p>'
  ).setWidth(360).setHeight(150);
  ui.showModalDialog(html, 'New year');
}


/* =============================================================================
 *  ONE-TIME SETUP / REPAIR
 *  Safe to run as often as you like. It never touches the data you typed
 *  (STATUS A:B, In Process D:H, COMPLETED A:B and H:L); it only (re)builds the
 *  formulas, colours, dropdown and the Stats tab.
 * ========================================================================== */

function setupOrRepair() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const ok = ui.alert('Set up / repair',
    'This creates any missing tabs, replaces the formulas in ' + CFG.STATUS + ' C:F and ' +
    CFG.DONE + ' C:F with single self-expanding formulas, rebuilds the date colours on ' +
    CFG.STATUS + ', sorts open projects by priority, adds an A–Z project dropdown on ' +
    CFG.LOG + ', and builds the ' + CFG.STATS + ' tab.\n\nYour typed data is not touched.',
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  // ---- Make sure the three core tabs exist (with headers) -------------------
  // On a brand-new spreadsheet this builds the whole layout from scratch.
  const status = ensureSheet_(ss, CFG.STATUS, 1,
    ['Project', 'Priority', 'Latest item', 'Action', 'Date', 'Link']);
  const log = ensureSheet_(ss, CFG.LOG, L_PROJECT,
    ['Project', 'Item', 'Action', 'Date', 'Link']);
  const done = ensureSheet_(ss, CFG.DONE, 1,
    ['Completed project', 'Priority', 'Final item', 'Action', 'Date', 'Link']);
  if (done.getRange(1, D_LOG).getValue() === '') {
    done.getRange(1, D_LOG, 1, L_WIDTH)
      .setValues([['Project', 'Item', 'Action', 'Date', 'Link']]).setFontWeight('bold');
  }

  const fr = CFG.STATUS_FIRST_ROW;
  const L = "'" + CFG.LOG + "'";       // quoted tab name, required when it has spaces

  // ---- STATUS C:F: newest log entry per project, as ONE formula -------------
  // How the formula works, from the inside out:
  //   FILTER(log D:H, project not blank)        -> all log rows
  //   HSTACK(..., ROW numbers)                  -> add a 6th column: sheet row
  //   SORT(..., col 4 date DESC, col 6 row DESC)-> newest first; same-day ties
  //                                                go to the row entered later
  //   VLOOKUP(name, ..., {2,3,4,5}, FALSE)      -> first (= newest) match,
  //                                                returning Item/Action/Date/Link
  //   ARRAYFORMULA(IF(name="",,...))            -> do it for every STATUS row
  // SEQUENCE(1,4,2) is {2,3,4,5} written in a locale-independent way.
  status.getRange(2, 3, status.getMaxRows() - 1, 4).clearContent();   // clear old per-row formulas
  status.getRange(fr, 3).setFormula(
    `=ARRAYFORMULA(IF(A${fr}:A="",,IFERROR(VLOOKUP(A${fr}:A,` +
    `SORT(HSTACK(FILTER(${L}!D2:H,${L}!D2:D<>""),FILTER(ROW(${L}!D2:D),${L}!D2:D<>"")),4,FALSE,6,FALSE),` +
    `SEQUENCE(1,4,2),FALSE))))`);
  status.getRange(`E${fr}:E`).setNumberFormat('M/d/yyyy');

  // ---- STATUS date colours ---------------------------------------------------
  // Google Sheets applies the FIRST matching rule, so order matters:
  //   grey   = latest Action mentions CFG.HOLD_WORD (on hold never goes stale)
  //   red    = no update for RED_DAYS or more
  //   yellow = no update for YELLOW_DAYS or more
  //   green  = updated recently
  const eRange = status.getRange(`E${fr}:E`);

  // Keep any of your own rules that don't touch column E; replace the rest.
  const keep = status.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getColumn() <= 5 && r.getLastColumn() >= 5));

  // $D is the Action column on STATUS. &"" turns blanks/numbers into text so
  // LOWER/REGEXMATCH never error.
  const isHold = `REGEXMATCH(LOWER($D${fr}&""),"${CFG.HOLD_WORD}")`;
  const olderThan = n => `=AND(ISNUMBER($E${fr}),$E${fr}<=TODAY()-${n},NOT(${isHold}))`;

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${fr}<>"",${isHold})`)
      .setBackground('#D9D9D9').setRanges([eRange]).build(),               // grey
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(olderThan(CFG.RED_DAYS))
      .setBackground('#EA9999').setRanges([eRange]).build(),               // red
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(olderThan(CFG.YELLOW_DAYS))
      .setBackground('#FFFF00').setRanges([eRange]).build(),               // yellow
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(ISNUMBER($E${fr}),$E${fr}>TODAY()-${CFG.YELLOW_DAYS})`)
      .setBackground('#B7E1CD').setRanges([eRange]).build(),               // green
  ];
  status.setConditionalFormatRules(rules.concat(keep));

  // ---- Optional: show the completed list on STATUS H:J ----------------------
  // Only added if that area is empty, so an existing layout is never overwritten.
  if (status.getRange('H1').getValue() === '' && status.getRange('H1').getFormula() === '') {
    status.getRange('H1').setFormula(
      `=QUERY('${CFG.DONE}'!A1:E,"select A, B, E where A is not null",1)`);
  }

  // ---- COMPLETED C:F: final entry per completed project ---------------------
  // Same technique as STATUS C:F, but looking at the archive in H:L.
  done.getRange(2, 3, done.getMaxRows() - 1, 4).clearContent();
  done.getRange(2, 3).setFormula(
    '=ARRAYFORMULA(IF(A2:A="",,IFERROR(VLOOKUP(A2:A,' +
    'SORT(HSTACK(FILTER(H2:L,H2:H<>""),FILTER(ROW(H2:H),H2:H<>"")),4,FALSE,6,FALSE),' +
    'SEQUENCE(1,4,2),FALSE))))');
  done.getRange('E2:E').setNumberFormat('M/d/yyyy');

  // ---- Stats tab (also holds the hidden A–Z list for the dropdown) ----------
  buildStats_(ss);

  // ---- In Process dropdown ----------------------------------------------------
  // Points at the alphabetical list in Stats!M. setAllowInvalid(true) means a
  // typo shows a warning triangle instead of being rejected outright.
  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(CFG.STATS).getRange('M4:M'), true)
    .setAllowInvalid(true)
    .build();
  log.getRange(2, L_PROJECT, log.getMaxRows() - 1, 1).setDataValidation(dv);
  log.getRange(2, L_DATE, log.getMaxRows() - 1, 1).setNumberFormat('M/d/yyyy');

  // ---- Optional grad-year tab, then tidy STATUS -------------------------------
  buildGradList_(ss);
  sortStatus();

  ui.alert('Done. Check the ' + CFG.STATUS + ' and ' + CFG.STATS + ' tabs.');
}


/* =============================================================================
 *  STATS TAB
 *  Everything on this tab is a live formula; no trigger or refresh needed.
 *
 *   A:B   summary numbers
 *   D:K   one row per project (open first, then closed), built by one formula
 *   M     hidden: open project names A–Z, used by the In Process dropdown
 * ========================================================================== */

function buildStats_(ss) {
  // Create the tab if needed (placed before the last tab), then wipe it.
  let sh = ss.getSheetByName(CFG.STATS);
  if (!sh) sh = ss.insertSheet(CFG.STATS, Math.max(0, ss.getSheets().length - 1));
  sh.clear();

  const L = "'" + CFG.LOG + "'";
  const S = "'" + CFG.STATUS + "'";
  const D = "'" + CFG.DONE + "'";
  const H = CFG.HOLD_WORD;
  const fr = CFG.STATUS_FIRST_ROW;

  sh.getRange('A1').setValue('Project Stats').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setValue('Everything here updates on its own.').setFontColor('#666666');

  // ---- Summary block (A4:B18) -------------------------------------------------
  // The "Avg ..." rows read from the per-project table in D:K below, where
  // column E = Status, H = Last update, I = Days active, J = Updates,
  // K = Updates per week.
  const summary = [
    ['Open projects',
      `=COUNTA(${S}!A${fr}:A)`],
    ['Closed projects (this year)',
      `=COUNTA(${D}!A2:A)`],
    [`Open, no update in ${CFG.YELLOW_DAYS}+ days (not on hold)`,
      `=COUNTIFS(${S}!A${fr}:A,"<>",${S}!E${fr}:E,"<="&(TODAY()-${CFG.YELLOW_DAYS}),${S}!D${fr}:D,"<>*${H}*")`],
    [`Open, no update in ${CFG.RED_DAYS}+ days (not on hold)`,
      `=COUNTIFS(${S}!A${fr}:A,"<>",${S}!E${fr}:E,"<="&(TODAY()-${CFG.RED_DAYS}),${S}!D${fr}:D,"<>*${H}*")`],
    ['Open, on hold',
      `=COUNTIFS(${S}!A${fr}:A,"<>",${S}!D${fr}:D,"*${H}*")`],
    ['', ''],
    ['Avg days to complete a project',
      '=IFERROR(AVERAGEIFS(I4:I,E4:E,"Closed"),"—")'],
    ['Avg age of open projects (days)',
      '=IFERROR(AVERAGEIFS(I4:I,E4:E,"Open"),"—")'],
    ['Avg updates per project',
      '=IFERROR(AVERAGE(J4:J),"—")'],
    ['Avg updates per project per week',
      '=IFERROR(AVERAGE(K4:K),"—")'],
    ['', ''],
    ['Total updates logged',
      `=COUNTA(${L}!D2:D)+COUNTA(${D}!H2:H)`],
    ['Updates in the last 7 days',
      `=COUNTIFS(${L}!G2:G,">"&(TODAY()-7))+COUNTIFS(${D}!K2:K,">"&(TODAY()-7))`],
    ['Avg updates per week (last 4 weeks)',
      `=(COUNTIFS(${L}!G2:G,">"&(TODAY()-28))+COUNTIFS(${D}!K2:K,">"&(TODAY()-28)))/4`],
    ['Projects closed in the last 30 days',
      '=COUNTIFS(E4:E,"Closed",H4:H,">"&(TODAY()-30))'],
  ];
  sh.getRange(4, 1, summary.length, 2).setValues(summary);
  sh.getRange(4, 2, summary.length, 1).setNumberFormat('0.#').setHorizontalAlignment('right');

  // ---- Per-project table (D3:K) -----------------------------------------------
  sh.getRange('D3:K3').setValues([['Project', 'Status', 'Priority', 'Started', 'Last update',
    'Days active', 'Updates', 'Updates/wk']]).setFontWeight('bold');

  // One LET formula builds the whole table. Variable names start with "x" so
  // they can never collide with Sheets function names (e.g. DAYS, RATE, OR).
  //   xLogName / xLogDate   every log entry, open and archived, stacked
  //   xOpen* / xDone*       project names + priorities from STATUS / COMPLETED
  //                         (IFERROR(...,"") keeps an empty list from erroring)
  //   xStat                 "Open" or "Closed" for each project
  //   xCnt                  number of log entries for the project
  //   xStart / xEnd         first and last log dates
  //   xDur                  days active: today - start (open) or end - start
  //                         (closed); MAX(0,...) guards against future dates
  //   xPace                 updates per week; projects under a week count as
  //                         one week so a single busy day doesn't skew it
  //   final SORT            Open before Closed, then by priority, then by name
  sh.getRange('D4').setFormula(
    `=LET(` +
    `xLogName,VSTACK(${L}!D2:D,${D}!H2:H),` +
    `xLogDate,VSTACK(${L}!G2:G,${D}!K2:K),` +
    `xOpenName,IFERROR(FILTER(${S}!A${fr}:A,${S}!A${fr}:A<>""),""),` +
    `xOpenPri,IFERROR(FILTER(${S}!B${fr}:B,${S}!A${fr}:A<>""),""),` +
    `xDoneName,IFERROR(FILTER(${D}!A2:A,${D}!A2:A<>""),""),` +
    `xDonePri,IFERROR(FILTER(${D}!B2:B,${D}!A2:A<>""),""),` +
    `xProj,VSTACK(xOpenName,xDoneName),` +
    `xStat,VSTACK(MAP(xOpenName,LAMBDA(p,"Open")),MAP(xDoneName,LAMBDA(p,"Closed"))),` +
    `xPri,VSTACK(xOpenPri,xDonePri),` +
    `xCnt,MAP(xProj,LAMBDA(p,IF(p="",0,SUMPRODUCT(--(xLogName=p))))),` +
    `xStart,MAP(xProj,xCnt,LAMBDA(p,n,IF(n=0,"",MIN(FILTER(xLogDate,xLogName=p))))),` +
    `xEnd,MAP(xProj,xCnt,LAMBDA(p,n,IF(n=0,"",MAX(FILTER(xLogDate,xLogName=p))))),` +
    `xDur,MAP(xStat,xStart,xEnd,LAMBDA(s,a,b,IF(a="","",MAX(0,IF(s="Open",TODAY(),b)-a)))),` +
    `xPace,MAP(xCnt,xDur,LAMBDA(n,d,IF(OR(n=0,d=""),"",n/MAX(1,(d+1)/7)))),` +
    `xTable,HSTACK(xProj,xStat,xPri,xStart,xEnd,xDur,xCnt,xPace),` +
    `IFNA(SORT(FILTER(xTable,xProj<>""),2,FALSE,3,TRUE,1,TRUE),"No projects yet"))`);

  sh.getRange('G4:H').setNumberFormat('M/d/yyyy');   // Started, Last update
  sh.getRange('I4:J').setNumberFormat('0');          // Days active, Updates
  sh.getRange('K4:K').setNumberFormat('0.0');        // Updates/wk
  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(4, 220);
  sh.setFrozenRows(3);

  // ---- Hidden helper column M: open projects A–Z -----------------------------
  // A dropdown always lists items in sheet order. STATUS is in priority order,
  // so the dropdown points here instead to get an alphabetical list.
  sh.getRange('M3').setValue('Dropdown list (A–Z)').setFontWeight('bold');
  sh.getRange('M4').setFormula(`=IFERROR(SORT(FILTER(${S}!A${fr}:A,${S}!A${fr}:A<>"")),"")`);
  sh.hideColumns(13);
}


/* =============================================================================
 *  OPTIONAL GRAD-YEAR TAB
 *  Only runs if a tab named CFG.GRAD exists. Replaces its A:B with formulas so
 *  it never needs updating by hand:
 *    A = graduation year, starting with the current senior class
 *    B = grade (12, 11, ... 0 = K, negatives = not yet in school)
 *  The senior class year becomes "next calendar year" once NEW_YEAR_MONTH
 *  arrives (e.g. from July 2026 onward, seniors are the class of 2027).
 * ========================================================================== */

function buildGradList_(ss) {
  const sh = ss.getSheetByName(CFG.GRAD);
  if (!sh) return;                                          // feature not used
  sh.getRange('A1:B').clearContent();
  sh.getRange('A1').setFormula(
    `=SEQUENCE(23,1,YEAR(TODAY())+IF(MONTH(TODAY())>=${CFG.NEW_YEAR_MONTH},1,0))`);
  sh.getRange('B1').setFormula('=SEQUENCE(23,1,12,-1)');
}


/* =============================================================================
 *  HELPERS
 * ========================================================================== */

/** Normalises a project name for comparison: trimmed and lower-case. */
function key_(v) {
  return String(v).trim().toLowerCase();
}

/**
 * Last row that has a value in the given column (0 if the column is empty).
 * Needed because sheet.getLastRow() looks at ALL columns, and formulas that
 * spill down (like QUERY) would throw it off.
 */
function lastDataRow_(sheet, col) {
  const vals = sheet.getRange(1, col, sheet.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') return i + 1;
  }
  return 0;
}

/** Adds rows to the bottom of a sheet if it is shorter than neededLastRow. */
function ensureRows_(sheet, neededLastRow) {
  const extra = neededLastRow - sheet.getMaxRows();
  if (extra > 0) sheet.insertRowsAfter(sheet.getMaxRows(), extra + 50);
}

/**
 * Groups a sorted list of row numbers into contiguous runs.
 * [5, 6, 7, 20, 21]  ->  [[5, 3], [20, 2]]   (each run = [startRow, count])
 */
function toRuns_(rows) {
  const runs = [];
  rows.forEach(r => {
    const last = runs[runs.length - 1];
    if (last && last[0] + last[1] === r) last[1]++;
    else runs.push([r, 1]);
  });
  return runs;
}

/**
 * Returns the named tab, creating it if missing. When created (or when its
 * header cell is empty), writes bold headers into row 1 starting at startCol.
 */
function ensureSheet_(ss, name, startCol, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getRange(1, startCol).getValue() === '' && sh.getRange(1, startCol).getFormula() === '') {
    sh.getRange(1, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Escapes text for safe use inside the HTML dialog. */
function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}
