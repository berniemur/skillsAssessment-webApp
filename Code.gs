/**
 * BASKETBALL SKILLS ASSESSMENT — SCORING APP
 * Google Apps Script, bound to a Google Sheet.
 *
 * Sheet tabs (created by Draft Tools > Set up sheets):
 *   Roster     – Player #, First Name, Last Name, Division
 *   Assessors  – Assessor Name
 *   Scores     – one row per assessor + player (written by the app)
 *   Summary    – built by Draft Tools > Build / refresh Summary
 */

// ======================= SETTINGS =======================
const CONFIG = {
  APP_TITLE: 'Skills Assessment Scoring',
  // Assessors type this once on their phone. Set to '' to turn it off.
  ACCESS_CODE: 'hoops2026',
  // true  = show "#12  Jake M." on the scoring screen
  // false = show only the player number
  SHOW_NAMES: true,
  // In the Summary, flag a player when assessors' scores differ by this much or more
  SPREAD_FLAG: 2,
};
// ========================================================

const SHEETS = {
  ROSTER: 'Roster',
  ASSESSORS: 'Assessors',
  SCORES: 'Scores',
  SUMMARY: 'Summary',
};

/* ---------------- Sheet menu & setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Draft Tools')
    .addItem('Set up sheets', 'setupSheets')
    .addItem('Build / refresh Summary', 'buildSummary')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roster = ensureSheet_(ss, SHEETS.ROSTER, ['Player #', 'First Name', 'Last Name', 'Division']);
  roster.getRange('A:A').setNumberFormat('@'); // keep numbers like "007" as typed
  ensureSheet_(ss, SHEETS.ASSESSORS, ['Assessor Name']);
  const scores = ensureSheet_(ss, SHEETS.SCORES, ['Updated', 'Assessor', 'Player #', 'Score']);
  scores.getRange('C:C').setNumberFormat('@');
  ensureSheet_(ss, SHEETS.SUMMARY, ['Use Draft Tools > Build / refresh Summary']);
  try {
    SpreadsheetApp.getUi().alert('Sheets are ready. Paste your roster into "Roster" and assessor names into "Assessors".');
  } catch (e) { /* no UI when run from the editor */ }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- Web app ---------------- */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appTitle = CONFIG.APP_TITLE;
  return t.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    // Needed so the page can be embedded in Google Sites
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called by the page on load. Returns assessor names and the roster. */
function getSetup(code) {
  checkCode_(code);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    title: CONFIG.APP_TITLE,
    assessors: getAssessors_(ss),
    players: getPlayers_(ss).map(function (p) {
      return {
        id: p.id,
        name: CONFIG.SHOW_NAMES ? shortName_(p.first, p.last) : '',
        division: p.division,
      };
    }),
  };
}

/** Returns { playerId: score } for one assessor, so they can pick up where they left off. */
function getMyScores(code, assessor) {
  checkCode_(code);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = {};
  readRows_(ss, SHEETS.SCORES).forEach(function (r) {
    if (String(r[1]).trim() === assessor && r[3] !== '' && r[3] !== null) {
      out[String(r[2]).trim()] = Number(r[3]);
    }
  });
  return out;
}

/**
 * Saves a batch of scores for one assessor.
 * entries: [{ id: '12', score: 4 }, ...]   score of 0 or '' clears it.
 * Each assessor has at most one row per player; re-scoring overwrites it.
 */
function saveScores(code, assessor, entries) {
  checkCode_(code);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  assessor = String(assessor || '').trim();
  if (getAssessors_(ss).indexOf(assessor) === -1) throw new Error('Unknown assessor: ' + assessor);
  if (!entries || !entries.length) return { saved: 0 };

  const validIds = {};
  getPlayers_(ss).forEach(function (p) { validIds[p.id] = true; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(SHEETS.SCORES);
    const last = sh.getLastRow();
    const index = {}; // "assessor|id" -> sheet row number
    if (last > 1) {
      sh.getRange(2, 2, last - 1, 2).getValues().forEach(function (r, i) {
        index[String(r[0]).trim() + '|' + String(r[1]).trim()] = i + 2;
      });
    }

    const now = new Date();
    const newRows = [];
    let saved = 0;
    entries.forEach(function (e) {
      const id = String(e.id).trim();
      if (!validIds[id]) return;
      let score = e.score === '' || e.score === null ? 0 : Number(e.score);
      if (!(score >= 0 && score <= 5) || Math.floor(score) !== score) return;
      const value = score === 0 ? '' : score;
      const key = assessor + '|' + id;
      if (index[key]) {
        sh.getRange(index[key], 1, 1, 4).setValues([[now, assessor, id, value]]);
      } else if (value !== '') {
        newRows.push([now, assessor, id, value]);
        index[key] = -1; // guard against duplicates within the same batch
      }
      saved++;
    });
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);
    }
    return { saved: saved };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Summary for the draft ---------------- */

function buildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const players = getPlayers_(ss);
  const assessors = getAssessors_(ss);

  // scores[id][assessor] = score
  const scores = {};
  readRows_(ss, SHEETS.SCORES).forEach(function (r) {
    const a = String(r[1]).trim(), id = String(r[2]).trim(), s = r[3];
    if (!a || !id || s === '' || s === null) return;
    if (assessors.indexOf(a) === -1) assessors.push(a); // keep scores from renamed/removed assessors
    (scores[id] = scores[id] || {})[a] = Number(s);
  });

  const header = ['Player #', 'First Name', 'Last Name', 'Division']
    .concat(assessors)
    .concat(['# Scores', 'Average', 'Low', 'High', 'Spread', 'Check']);

  const rows = players.map(function (p) {
    const ps = scores[p.id] || {};
    const vals = assessors.map(function (a) { return ps[a] !== undefined ? ps[a] : ''; });
    const nums = vals.filter(function (v) { return v !== ''; });
    const n = nums.length;
    const avg = n ? nums.reduce(function (s, v) { return s + v; }, 0) / n : '';
    const lo = n ? Math.min.apply(null, nums) : '';
    const hi = n ? Math.max.apply(null, nums) : '';
    const spread = n ? hi - lo : '';
    let check = '';
    if (n === 0) check = 'No scores';
    else if (n < assessors.length) check = 'Missing ' + (assessors.length - n);
    if (n && spread >= CONFIG.SPREAD_FLAG) check = (check ? check + '; ' : '') + 'Big spread';
    return [p.id, p.first, p.last, p.division].concat(vals).concat([n, avg, lo, hi, spread, check]);
  });

  // Sort: division A→Z, then average high→low (unscored last)
  rows.sort(function (a, b) {
    const d = String(a[3]).localeCompare(String(b[3]));
    if (d) return d;
    const avgCol = 4 + assessors.length + 1;
    const x = a[avgCol] === '' ? -1 : a[avgCol];
    const y = b[avgCol] === '' ? -1 : b[avgCol];
    return y - x;
  });

  let sh = ss.getSheetByName(SHEETS.SUMMARY) || ss.insertSheet(SHEETS.SUMMARY);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e8eef7');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    const avgColNum = 4 + assessors.length + 2; // 1-based
    const avgRange = sh.getRange(2, avgColNum, rows.length, 1);
    avgRange.setNumberFormat('0.00');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#f4c7c3', SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMidpointWithValue('#fff2cc', SpreadsheetApp.InterpolationType.NUMBER, '3')
        .setGradientMaxpointWithValue('#b7e1cd', SpreadsheetApp.InterpolationType.NUMBER, '5')
        .setRanges([avgRange])
        .build(),
    ]);
  }
  sh.autoResizeColumns(1, header.length);
  sh.activate();
}

/* ---------------- Helpers ---------------- */

function checkCode_(code) {
  if (CONFIG.ACCESS_CODE && String(code || '').trim() !== CONFIG.ACCESS_CODE) {
    throw new Error('ACCESS_CODE');
  }
}

function readRows_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 1)).getValues();
}

function getAssessors_(ss) {
  const seen = {};
  return readRows_(ss, SHEETS.ASSESSORS)
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n && !seen[n] && (seen[n] = true); });
}

function getPlayers_(ss) {
  const seen = {};
  return readRows_(ss, SHEETS.ROSTER)
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        first: String(r[1] || '').trim(),
        last: String(r[2] || '').trim(),
        division: String(r[3] || '').trim(),
      };
    })
    .filter(function (p) { return p.id && !seen[p.id] && (seen[p.id] = true); });
}

function shortName_(first, last) {
  return (first + (last ? ' ' + last.charAt(0) + '.' : '')).trim();
}
