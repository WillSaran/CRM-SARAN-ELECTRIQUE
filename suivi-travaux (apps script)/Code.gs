/*******************************************************
 * SUIVI TRAVAUX — Google Calendar -> Google Sheets
 * VERSION CORRIGÉE V3 — juillet 2026
 *
 * V3 — NOUVEAU (par rapport a V2) :
 * 14. Rapport sans facture : le courriel est maintenant separe en
 *     DEUX sections. En haut : "FACTURES MANQUANTES" (aucun appel
 *     de service enregistre). En bas, dans sa propre section separee :
 *     "EN TRAITEMENT" (appel de service deja fait, facture a venir).
 *     Avant, les deux etaient melanges ensemble sous chaque feuille.
 *
 * (V1 et V2 : voir commentaires plus bas / versions precedentes —
 *  protection K/L/M, retry Calendar, isolation par feuille, ecriture
 *  en place au lieu de clear-puis-write, alertes courriel automatiques.)
 *******************************************************/

/** ================== CONFIG ================== */
const CONFIG = {
  // NOUVEAU — ID fixe de la feuille 2026-2027 (plus besoin de compter sur "feuille active")
  SPREADSHEET_ID: "1uHgs4eiD0t25hYHp69q27tESsuVB4HraCud3c04vLwY",
  CALENDAR_ID: "saranelectrique@gmail.com",

  // NOUVEAU — nouvelle annee fiscale : on part le 1er juillet 2026, plus le 1er janvier
  START_DATE_ISO: "2026-07-01",
  FUTURE_DAYS: 365,

  HEADER_ROWS: 2,
  HEADER_BG: "#111111",
  HEADER_FG: "#ffffff",

  COL_DATE: 1,         // A
  COL_TITLE: 2,        // B ← couleur ici
  COL_ADDRESS: 3,      // C ← bidirectionnel
  COL_DESCRIPTION: 4,  // D ← bidirectionnel
  COL_EVENT_ID: 14,    // N

  MANUAL_COLS: [11, 12, 13], // K, L, M
  MANUAL_LABELS: {
    11: "Appel de service",
    12: "Ordre de travail",
    13: "Facture",
  },

  MANUAL_BLOCK_PREFIX: "── Suivi ──",
  MANUAL_BLOCK_SUFFIX: "── ↑ ──",

  WRITE_COLS: 14,

  ROW_HEIGHT_NORMAL: 21,
  ROW_HEIGHT_MONTH: 24,
  ROW_HEIGHT_WEEK: 6,

  MONTH_ROW_BG: "#111111",
  MONTH_ROW_FG: "#ffffff",
  MONTH_ROW_TEXT_COL: 2,

  WEEK_ROW_BG: "#111111",

  BACKUP_SHEET_NAME: "BACKUP_MANUAL",

  ROUTING_RULES: [
    { sheet: "ASC",          test: /(^|\s)(ASC)(\s|$)/i },
    { sheet: "CRS",          test: /(^|\s)(CRS)(\s|$)/i },
    { sheet: "PAY",          test: /(^|\s)(PAY)(\s|$)/i },
    { sheet: "JOY",          test: /(^|\s)(JOY)(\s|$)/i },
    { sheet: "SAR",          test: /(^|\s)(SAR)(\s|$)/i },
    { sheet: "GGP",          test: /(^|\s)(GGP)(\s|$)/i },
    { sheet: "Préparations", test: /(^|\s)(PREP)(\s|$)/i },
    { sheet: "BOM",          test: /(^|\s)(BOM)(\s|$)/i },
    { sheet: "5PIL",         test: /(^|\s)(5PIL)(\s|$)/i },
    { sheet: "RROY",         test: /(^|\s)(RROY)(\s|$)/i },
    { sheet: "GOU",          test: /(^|\s)(GOU)(\s|$)/i },
  ],
  DEFAULT_SHEET: "Particuliers",
};

// ================== CONFIG RAPPORT ==================
const RAPPORT_CONFIG = {
  EMAIL_MOI:    "info@saranelectrique.com",
  EMAIL_EQUIPE: "saranelectrique@gmail.com",

  COL_DATE:    1,   // A
  COL_PROJET:  2,   // B
  COL_APPEL:   11,  // K
  COL_OT:      12,  // L
  COL_FACTURE: 13,  // M

  LIGNE_DEBUT: 3,
};

// ================== CONFIG ALERTES ==================
const ALERT_CONFIG = {
  ADMIN_EMAIL: "info@saranelectrique.com",
  EXEC_BUDGET_MS: 4.5 * 60 * 1000,
};
// =====================================================

/** ================== MENUS ================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SUIVI TRAVAUX")
    .addItem("Sync (récent)", "sync_recent")
    .addItem("Backfill (2026 -> futur)", "backfill_2026_to_future")
    .addSeparator()
    .addItem("Restaurer K/L/M depuis backup", "restoreFromBackup")
    .addSeparator()
    .addItem("Rapport sans facture", "envoyerRapportSansFacture")
    .addSeparator()
    .addItem("RESET (toutes feuilles, garde headers)", "reset_all_sheets_keep_headers")
    .addItem("Normaliser hauteur des lignes (toutes feuilles)", "normalize_all_row_heights")
    .addSeparator()
    .addItem("Installer / réinstaller trigger nightly", "_setupSyncTrigger_")
    .addToUi();
}

/** ================== ALERTES SÛRES ================== */
function _alertSafe_(title, message) {
  Logger.log(`[ALERTE] ${title} — ${message}`);
  try {
    MailApp.sendEmail(ALERT_CONFIG.ADMIN_EMAIL, "SUIVI TRAVAUX — " + title, message);
  } catch (e) {
    Logger.log("Echec envoi courriel d'alerte : " + e.message);
  }
  try {
    SpreadsheetApp.getUi().alert(title + "\n\n" + message);
  } catch (e) {}
}

/** ================== RETRY ================== */
function _withRetry_(fn, attempts, label) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      Logger.log(`[RETRY] ${label} — tentative ${i}/${attempts} echouee : ${e.message}`);
      if (i < attempts) Utilities.sleep(1000 * i);
    }
  }
  throw lastErr;
}

/** ================== ACTIONS PRINCIPALES ================== */
function sync_recent() {
  try {
    const tz = getTimeZone_();
    const start = isoToDate_(CONFIG.START_DATE_ISO, tz);
    const end = new Date(new Date().getTime() + CONFIG.FUTURE_DAYS * 24 * 3600 * 1000);
    syncRangeAllSheets_(start, end);
  } catch (e) {
    _alertSafe_("ECHEC sync_recent", "Le sync a plante completement :\n\n" + e.message + "\n\n" + (e.stack || ""));
    throw e;
  }
}

function backfill_2026_to_future() {
  try {
    const tz = getTimeZone_();
    const start = isoToDate_(CONFIG.START_DATE_ISO, tz);
    const end = new Date(new Date().getTime() + CONFIG.FUTURE_DAYS * 24 * 3600 * 1000);
    syncRangeAllSheets_(start, end);
  } catch (e) {
    _alertSafe_("ECHEC backfill", "Le backfill a plante completement :\n\n" + e.message + "\n\n" + (e.stack || ""));
    throw e;
  }
}

function reset_all_sheets_keep_headers() {
  const ss = getSpreadsheet_();
  ss.getSheets().forEach(sh => {
    if (!sh) return;
    resetSheetKeepHeader_(sh, CONFIG.HEADER_ROWS, CONFIG.WRITE_COLS);
    normalizeRowHeights_(sh, CONFIG.HEADER_ROWS + 1, CONFIG.ROW_HEIGHT_NORMAL);
  });
}

function normalize_all_row_heights() {
  const ss = getSpreadsheet_();
  ss.getSheets().forEach(sh => {
    normalizeRowHeights_(sh, CONFIG.HEADER_ROWS + 1, CONFIG.ROW_HEIGHT_NORMAL);
  });
}

/** ================== TRIGGER SETUP ================== */
function _setupSyncTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sync_recent") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sync_recent")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .nearMinute(0)
    .create();
  const msg = "Trigger nightly installe : sync_recent tous les jours vers 3h du matin.";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

/** ================== BACKUP SHEET ================== */
function getOrCreateBackupSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.BACKUP_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.BACKUP_SHEET_NAME);
    sh.hideSheet();
    sh.getRange(1, 1, 1, 6).setValues([["CLE", "FEUILLE", "K", "L", "M", "TIMESTAMP"]]);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#222222").setFontColor("#ffffff");
  }
  return sh;
}

function buildBackupKey_(sheetName, eventId, dateVal) {
  let dateKey = "";
  if (dateVal) {
    const parts = String(dateVal).split("/");
    if (parts.length === 3) {
      dateKey = `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
  }
  const idPart = dateKey ? `${eventId}__${dateKey}` : eventId;
  return `${sheetName}__${idPart}`;
}

function backupAllManualData_() {
  const ss = getSpreadsheet_();
  const backupSh = getOrCreateBackupSheet_();
  const timestamp = new Date().toISOString();

  const newRows = [];
  const sheetNames = CONFIG.ROUTING_RULES.map(r => r.sheet).concat([CONFIG.DEFAULT_SHEET]);

  for (const sheetName of sheetNames) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    const dataStart = CONFIG.HEADER_ROWS + 1;
    if (lastRow < dataStart) continue;

    const numRows = lastRow - dataStart + 1;
    const allData = sh.getRange(dataStart, 1, numRows, CONFIG.WRITE_COLS).getValues();

    for (const row of allData) {
      const eventId = String(row[CONFIG.COL_EVENT_ID - 1] || "").trim();
      if (!eventId) continue;

      const dateVal = row[CONFIG.COL_DATE - 1];
      const key = buildBackupKey_(sheetName, eventId, dateVal);

      const kVal = row[10]; // K
      const lVal = row[11]; // L
      const mVal = row[12]; // M

      const hasData = [kVal, lVal, mVal].some(v => v !== "" && v !== null && v !== undefined);
      if (!hasData) continue;

      newRows.push([key, sheetName, kVal, lVal, mVal, timestamp]);
    }
  }

  const lastBackup = backupSh.getLastRow();
  if (lastBackup > 1) {
    backupSh.getRange(2, 1, lastBackup - 1, 6).clearContent();
  }

  if (newRows.length > 0) {
    backupSh.getRange(2, 1, newRows.length, 6).setValues(newRows);
    backupSh.getRange(2, 3, newRows.length, 3).setNumberFormat("@");
  }

  Logger.log(`[BACKUP] ${newRows.length} entrees sauvegardees dans le backup.`);
}

function readBackupMap_(backupSh) {
  const lastRow = backupSh.getLastRow();
  if (lastRow < 2) return new Map();

  const data = backupSh.getRange(2, 1, lastRow - 1, 5).getValues();
  const map = new Map();
  for (const row of data) {
    const key = String(row[0] || "").trim();
    if (!key) continue;
    map.set(key, {
      11: row[2], // K
      12: row[3], // L
      13: row[4], // M
    });
  }
  return map;
}

function readBackupMapForSheet_(backupSh, sheetName) {
  const lastRow = backupSh.getLastRow();
  if (lastRow < 2) return new Map();

  const data = backupSh.getRange(2, 1, lastRow - 1, 6).getValues();
  const map = new Map();
  for (const row of data) {
    const key      = String(row[0] || "").trim();
    const rowSheet = String(row[1] || "").trim();
    if (!key) continue;
    if (rowSheet !== sheetName) continue;
    map.set(key, {
      11: row[2], // K
      12: row[3], // L
      13: row[4], // M
    });
  }
  return map;
}

function restoreFromBackup() {
  const ss = getSpreadsheet_();
  const backupSh = getOrCreateBackupSheet_();

  if (backupSh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Le backup est vide — rien a restaurer.");
    return;
  }

  const sheetNames = CONFIG.ROUTING_RULES.map(r => r.sheet).concat([CONFIG.DEFAULT_SHEET]);
  let totalRestored = 0;

  for (const sheetName of sheetNames) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) continue;

    const backupMap = readBackupMapForSheet_(backupSh, sheetName);
    if (backupMap.size === 0) continue;

    const lastRow = sh.getLastRow();
    const dataStart = CONFIG.HEADER_ROWS + 1;
    if (lastRow < dataStart) continue;

    const numRows = lastRow - dataStart + 1;
    const idCol   = sh.getRange(dataStart, CONFIG.COL_EVENT_ID, numRows, 1).getValues();
    const dateCol = sh.getRange(dataStart, CONFIG.COL_DATE, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      const eventId = String(idCol[i][0] || "").trim();
      if (!eventId) continue;

      const key   = buildBackupKey_(sheetName, eventId, dateCol[i][0]);
      const saved = backupMap.get(key);
      if (!saved) continue;

      const rowIndex = dataStart + i;
      for (const col of CONFIG.MANUAL_COLS) {
        const val = saved[col];
        if (val !== "" && val !== null && val !== undefined) {
          sh.getRange(rowIndex, col).setValue(val);
          totalRestored++;
        }
      }
    }

    if (lastRow >= dataStart) {
      for (const col of CONFIG.MANUAL_COLS) {
        sh.getRange(dataStart, col, lastRow - dataStart + 1, 1).setNumberFormat("@");
      }
    }
  }

  SpreadsheetApp.getUi().alert("Restauration terminee — " + totalRestored + " cellules restaurees depuis le backup.");
}

/** ================== SYNC COEUR ================== */
function syncRangeAllSheets_(start, end) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("[LOCK] Une synchro est deja en cours — on saute ce cycle.");
    return;
  }

  try {
    const ss = getSpreadsheet_();
    const tz = getTimeZone_();
    const execStart = Date.now();

    backupAllManualData_();

    const { items: events, colorMap, defaultCalColorHex } = fetchCalendarEventsWithColors_(CONFIG.CALENDAR_ID, start, end);
    const originalDataMap = buildOriginalDataMap_(events);
    const records = buildRecords_(events, tz, colorMap, defaultCalColorHex);
    const grouped = groupBySheet_(records);

    const sheetNames = Object.keys(grouped);
    const skipped = [];
    const errored = [];

    for (const sheetName of sheetNames) {
      if (Date.now() - execStart > ALERT_CONFIG.EXEC_BUDGET_MS) {
        skipped.push(sheetName);
        continue;
      }

      try {
        const sh = getOrCreateWorkSheet_(ss, sheetName);
        writeSheet_(sh, grouped[sheetName], tz, originalDataMap);
      } catch (e) {
        errored.push(`${sheetName} : ${e.message}`);
        Logger.log(`[ERREUR] Feuille ${sheetName} : ${e.message}`);
      }
    }

    if (skipped.length > 0) {
      _alertSafe_(
        "Sync incomplet (budget de temps)",
        "Ces feuilles n'ont pas ete traitees cette fois (budget de temps atteint), elles seront a jour au prochain sync :\n\n" +
        skipped.join(", ")
      );
    }
    if (errored.length > 0) {
      _alertSafe_(
        "Erreurs pendant le sync",
        "Ces feuilles ont echoue pendant le sync (les autres ont ete traitees normalement) :\n\n" +
        errored.join("\n")
      );
    }
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateWorkSheet_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (sh) return sh;

  Logger.log("[" + sheetName + "] Onglet introuvable -> creation automatique.");
  sh = ss.insertSheet(sheetName);

  const sheetNames = CONFIG.ROUTING_RULES.map(r => r.sheet).concat([CONFIG.DEFAULT_SHEET]);
  let modele = null;
  for (const nom of sheetNames) {
    const candidat = ss.getSheetByName(nom);
    if (candidat && candidat.getName() !== sheetName) {
      modele = candidat;
      break;
    }
  }

  if (modele) {
    for (let col = 1; col <= CONFIG.WRITE_COLS; col++) {
      try {
        sh.setColumnWidth(col, modele.getColumnWidth(col));
      } catch (e) {}
    }
  }

  sh.getRange(1, 1, 1, CONFIG.WRITE_COLS)
    .setBackground(CONFIG.HEADER_BG)
    .setFontColor(CONFIG.HEADER_FG);

  const titres = new Array(CONFIG.WRITE_COLS).fill("");
  titres[CONFIG.COL_DATE - 1]        = "Date";
  titres[CONFIG.COL_TITLE - 1]       = "Projet / Evenement";
  titres[CONFIG.COL_ADDRESS - 1]     = "Adresse";
  titres[CONFIG.COL_DESCRIPTION - 1] = "Description";
  titres[10] = "Appel de service";
  titres[11] = "Ordre de travail";
  titres[12] = "Facture";
  titres[CONFIG.COL_EVENT_ID - 1]    = "Event ID";

  sh.getRange(2, 1, 1, CONFIG.WRITE_COLS)
    .setValues([titres])
    .setBackground(CONFIG.HEADER_BG)
    .setFontColor(CONFIG.HEADER_FG)
    .setFontWeight("bold");

  sh.setRowHeight(1, CONFIG.ROW_HEIGHT_MONTH);
  sh.setRowHeight(2, CONFIG.ROW_HEIGHT_MONTH);

  try { sh.setFrozenRows(CONFIG.HEADER_ROWS); } catch (e) {}

  return sh;
}

/** ================== FETCH EVENTS ================== */
function fetchCalendarEventsWithColors_(calendarId, start, end) {
  const colors = Calendar.Colors.get();
  const eventColors = (colors && colors.event) ? colors.event : {};
  const calColors = (colors && colors.calendar) ? colors.calendar : {};

  let defaultCalColorHex = "#bdbdbd";
  try {
    const calEntry = Calendar.CalendarList.get(calendarId);
    const calColorId = calEntry && calEntry.colorId ? String(calEntry.colorId) : null;
    if (calColorId && calColors[calColorId] && calColors[calColorId].background) {
      defaultCalColorHex = calColors[calColorId].background;
    }
  } catch (e) {}

  const timeMin = start.toISOString();
  const timeMax = end.toISOString();
  let pageToken = null;
  const all = [];
  do {
    const resp = _withRetry_(
      () => Calendar.Events.list(calendarId, {
        timeMin, timeMax,
        singleEvents: true,
        orderBy: "startTime",
        showDeleted: false,
        maxResults: 2500,
        pageToken,
      }),
      3,
      "Calendar.Events.list"
    );
    if (resp && resp.items && resp.items.length) all.push(...resp.items);
    pageToken = resp.nextPageToken;
  } while (pageToken);

  const colorMap = {};
  Object.keys(eventColors).forEach(id => {
    const c = eventColors[id];
    if (c && c.background) colorMap[String(id)] = c.background;
  });

  return { items: all, colorMap, defaultCalColorHex };
}

function buildOriginalDataMap_(events) {
  const map = new Map();
  for (const ev of events) {
    if (!ev || !ev.id) continue;
    map.set(ev.id, {
      originalDesc: stripSuiviBlock_(ev.description || ""),
      originalLocation: ev.location || "",
    });
  }
  return map;
}

function stripSuiviBlock_(desc) {
  if (!desc) return "";
  const prefix = CONFIG.MANUAL_BLOCK_PREFIX;
  const suffix = CONFIG.MANUAL_BLOCK_SUFFIX;
  const start = desc.indexOf(prefix);
  const end = desc.indexOf(suffix);
  if (start === -1) return desc;
  if (end === -1) return desc.substring(0, start).trim();
  const before = desc.substring(0, start).trim();
  const after = desc.substring(end + suffix.length).trim();
  return [before, after].filter(s => s).join("\n");
}

/** ================== RECORDS ================== */
function buildRecords_(events, tz, colorMap, defaultCalColorHex) {
  const minStartMs = isoToDate_(CONFIG.START_DATE_ISO, tz).getTime();
  const seen = new Set();
  const out = [];

  for (const ev of events) {
    if (!ev || !ev.start) continue;
    const id = ev.id || "";
    const title = (ev.summary || "").trim();
    const location = (ev.location || "").trim();
    const description = stripSuiviBlock_(ev.description || "").trim();
    const colorHex = resolveEventColorHex_(ev, colorMap, defaultCalColorHex);
    const dayStarts = expandEventToDailyStartDates_(ev, tz);
    if (!dayStarts.length) continue;

    for (const dayStart of dayStarts) {
      if (dayStart.getTime() < minStartMs) continue;
      const key = `${id}__${Utilities.formatDate(dayStart, tz, "yyyy-MM-dd")}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ eventId: id, startDate: dayStart, isAllDay: !!(ev.start && ev.start.date), title, location, description, colorHex });
    }
  }
  return out;
}

function expandEventToDailyStartDates_(ev, tz) {
  let start, end;
  if (ev.start && ev.start.date) {
    start = isoToDate_(ev.start.date, tz);
    end = isoToDate_(ev.end.date, tz);
    end.setDate(end.getDate() - 1);
  } else if (ev.start && ev.start.dateTime) {
    start = new Date(ev.start.dateTime);
    end = ev.end && ev.end.dateTime ? new Date(ev.end.dateTime) : new Date(start);
  } else {
    return [];
  }
  const cur = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 12, 0, 0);
  const last = new Date(end.getFullYear(), end.getMonth(), end.getDate(), 12, 0, 0);
  const days = [];
  while (cur.getTime() <= last.getTime()) {
    days.push(new Date(cur));
    cur.setDate(cur.getDate() + 1);
  }
  return days;
}

function resolveEventColorHex_(ev, colorMap, defaultCalColorHex) {
  const colorId = ev.colorId ? String(ev.colorId) : null;
  if (colorId && colorMap[colorId]) return colorMap[colorId];
  return defaultCalColorHex;
}

/** ================== ROUTING ================== */
function groupBySheet_(records) {
  const grouped = {};
  for (const r of records) {
    const target = routeSheet_(r.title);
    if (!grouped[target]) grouped[target] = [];
    grouped[target].push(r);
  }
  return grouped;
}

function routeSheet_(title) {
  for (const rule of CONFIG.ROUTING_RULES) {
    if (rule.test.test(title)) return rule.sheet;
  }
  return CONFIG.DEFAULT_SHEET;
}

/** ================== WEEK KEY ================== */
function getWeekKey_(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day === 0) ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}

/** ================== SAUVEGARDE COLONNES MANUELLES ================== */
function saveManualData_(sheet, sheetName) {
  const lastRow = sheet.getLastRow();
  const dataStart = CONFIG.HEADER_ROWS + 1;
  if (lastRow < dataStart) return new Map();

  const numRows = lastRow - dataStart + 1;
  const allData = sheet.getRange(dataStart, 1, numRows, CONFIG.WRITE_COLS).getValues();
  const saved = new Map();

  for (const row of allData) {
    const eventId = String(row[CONFIG.COL_EVENT_ID - 1] || "").trim();
    if (!eventId) continue;

    const dateVal = row[CONFIG.COL_DATE - 1];
    const key = buildBackupKey_(sheetName, eventId, dateVal);

    const manualValues = {};
    for (const col of CONFIG.MANUAL_COLS) {
      manualValues[col] = row[col - 1];
    }
    manualValues[CONFIG.COL_ADDRESS]     = row[CONFIG.COL_ADDRESS - 1];
    manualValues[CONFIG.COL_DESCRIPTION] = row[CONFIG.COL_DESCRIPTION - 1];

    const hasData =
      CONFIG.MANUAL_COLS.some(col => {
        const v = manualValues[col];
        return v !== "" && v !== null && v !== undefined;
      }) ||
      manualValues[CONFIG.COL_ADDRESS] !== "" ||
      manualValues[CONFIG.COL_DESCRIPTION] !== "";

    if (hasData) saved.set(key, manualValues);
  }

  return saved;
}

function restoreManualData_(sheet, sheetName, saved, backupMap) {
  const merged = new Map(backupMap);
  for (const [key, val] of saved) {
    merged.set(key, val);
  }

  if (merged.size === 0) {
    Logger.log(`[${sheetName}] restoreManualData_ : rien a restaurer (merged vide).`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const dataStart = CONFIG.HEADER_ROWS + 1;
  if (lastRow < dataStart) return;

  const numRows = lastRow - dataStart + 1;
  const idCol   = sheet.getRange(dataStart, CONFIG.COL_EVENT_ID, numRows, 1).getValues();
  const dateCol = sheet.getRange(dataStart, CONFIG.COL_DATE, numRows, 1).getValues();

  let restoredCount = 0;

  for (let i = 0; i < numRows; i++) {
    const eventId = String(idCol[i][0] || "").trim();
    if (!eventId) continue;

    const key = buildBackupKey_(sheetName, eventId, dateCol[i][0]);
    const manualValues = merged.get(key);
    if (!manualValues) continue;

    const rowIndex = dataStart + i;

    for (const col of CONFIG.MANUAL_COLS) {
      const val = manualValues[col];
      if (val !== "" && val !== null && val !== undefined) {
        sheet.getRange(rowIndex, col).setValue(val);
        restoredCount++;
      }
    }

    const savedAddr = manualValues[CONFIG.COL_ADDRESS];
    if (savedAddr !== "" && savedAddr !== null && savedAddr !== undefined) {
      sheet.getRange(rowIndex, CONFIG.COL_ADDRESS).setValue(savedAddr);
    }

    const savedDesc = manualValues[CONFIG.COL_DESCRIPTION];
    if (savedDesc !== "" && savedDesc !== null && savedDesc !== undefined) {
      sheet.getRange(rowIndex, CONFIG.COL_DESCRIPTION).setValue(savedDesc);
    }
  }

  Logger.log(`[${sheetName}] restoreManualData_ : ${restoredCount} cellules K/L/M restaurees.`);
}

/** ================== FORMAT TEXTE COLONNES MANUELLES ================== */
function setManualColsAsText_(sheet) {
  const lastRow = sheet.getLastRow();
  const dataStart = CONFIG.HEADER_ROWS + 1;
  if (lastRow < dataStart) return;
  const numRows = lastRow - dataStart + 1;
  for (const col of CONFIG.MANUAL_COLS) {
    sheet.getRange(dataStart, col, numRows, 1).setNumberFormat("@");
  }
}

/** ================== PUSH VERS AGENDA ================== */
function pushToCalendar_(sheet, originalDataMap) {
  const lastRow = sheet.getLastRow();
  const dataStart = CONFIG.HEADER_ROWS + 1;
  if (lastRow < dataStart) return;

  const numRows = lastRow - dataStart + 1;
  const allData = sheet.getRange(dataStart, 1, numRows, CONFIG.WRITE_COLS).getValues();
  const alreadyPatched = new Set();

  for (const row of allData) {
    const eventId = String(row[CONFIG.COL_EVENT_ID - 1] || "").trim();
    if (!eventId) continue;
    if (alreadyPatched.has(eventId)) continue;

    const original = originalDataMap.get(eventId);
    if (!original) continue;

    const sheetAddress = String(row[CONFIG.COL_ADDRESS - 1] || "").trim();
    const sheetDesc    = String(row[CONFIG.COL_DESCRIPTION - 1] || "").trim();

    const addressChanged = sheetAddress !== original.originalLocation.trim();
    const descChanged    = sheetDesc !== original.originalDesc.trim();

    const hasManual = CONFIG.MANUAL_COLS.some(col => {
      const v = row[col - 1];
      return v !== "" && v !== null && v !== undefined;
    });

    if (!addressChanged && !descChanged && !hasManual) continue;

    const patch = {};

    if (addressChanged) patch.location = sheetAddress;

    const baseDesc = descChanged ? sheetDesc : original.originalDesc;

    if (hasManual) {
      const lines = [];
      for (const col of CONFIG.MANUAL_COLS) {
        const val = row[col - 1];
        if (val !== "" && val !== null && val !== undefined) {
          const label = CONFIG.MANUAL_LABELS[col] || `Col ${col}`;
          lines.push(`${label} : ${val}`);
        }
      }
      const suiviBlock = `${CONFIG.MANUAL_BLOCK_PREFIX}\n${lines.join("\n")}\n${CONFIG.MANUAL_BLOCK_SUFFIX}`;
      patch.description = baseDesc ? `${suiviBlock}\n\n${baseDesc}` : suiviBlock;
    } else if (descChanged) {
      patch.description = baseDesc;
    }

    try {
      Calendar.Events.patch(patch, CONFIG.CALENDAR_ID, eventId);
      alreadyPatched.add(eventId);
    } catch (e) {
      Logger.log(`Erreur patch eventId ${eventId} : ${e.message}`);
    }
  }
}

/** ================== WRITE SHEET ================== */
function writeSheet_(sheet, records, tz, originalDataMap) {
  const sheetName = sheet.getName();

  const backupSh  = getOrCreateBackupSheet_();
  const backupMap = readBackupMapForSheet_(backupSh, sheetName);
  const savedManual = saveManualData_(sheet, sheetName);

  Logger.log(`[${sheetName}] savedManual.size = ${savedManual.size}, backupMap.size = ${backupMap.size}`);

  const lastRowBefore = sheet.getLastRow();
  const dataStartCheck = CONFIG.HEADER_ROWS + 1;
  const oldDataRows = Math.max(0, lastRowBefore - CONFIG.HEADER_ROWS);

  if (lastRowBefore >= dataStartCheck && savedManual.size === 0 && backupMap.size === 0) {
    const numRowsCheck = lastRowBefore - dataStartCheck + 1;
    const kRange = sheet.getRange(dataStartCheck, 11, numRowsCheck, 3).getValues();
    const hadManualData = kRange.some(row => row.some(v => v !== "" && v !== null && v !== undefined));
    if (hadManualData) {
      Logger.log(`[${sheetName}] ABORT RESET : donnees K/L/M detectees mais aucune cle sauvegardee.`);
      _alertSafe_(
        "Feuille \"" + sheetName + "\" protegee",
        "Des donnees K/L/M ont ete detectees mais les cles de correspondance sont introuvables.\n" +
        "Le reset de cette feuille a ete ANNULE pour proteger vos donnees.\n\n" +
        "-> Verifiez que la colonne N (Event ID) est bien remplie et relancez la sync."
      );
      return;
    }
  }

  records.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

  const rows = [];
  const rowMeta = [];
  let lastMonthKey = null;
  let lastWeekKey = null;

  for (const r of records) {
    const monthKey = Utilities.formatDate(r.startDate, tz, "yyyy-MM");
    const weekKey = getWeekKey_(r.startDate);

    if (monthKey !== lastMonthKey) {
      lastMonthKey = monthKey;
      lastWeekKey = null;
      const monthLabel = monthLabelFR_(r.startDate, tz);
      const arr = new Array(CONFIG.WRITE_COLS).fill("");
      arr[CONFIG.MONTH_ROW_TEXT_COL - 1] = monthLabel;
      rows.push(arr);
      rowMeta.push({ type: "month" });
    }

    if (weekKey !== lastWeekKey) {
      lastWeekKey = weekKey;
      if (rowMeta.length > 0 && rowMeta[rowMeta.length - 1].type !== "month") {
        rows.push(new Array(CONFIG.WRITE_COLS).fill(""));
        rowMeta.push({ type: "week" });
      }
    }

    const key = buildBackupKey_(sheetName, r.eventId, Utilities.formatDate(r.startDate, tz, "dd/MM/yyyy"));
    const saved = savedManual.get(key) || backupMap.get(key);

    const address = (saved && saved[CONFIG.COL_ADDRESS] !== "" && saved[CONFIG.COL_ADDRESS] != null)
      ? saved[CONFIG.COL_ADDRESS]
      : r.location;

    const description = (saved && saved[CONFIG.COL_DESCRIPTION] !== "" && saved[CONFIG.COL_DESCRIPTION] != null)
      ? saved[CONFIG.COL_DESCRIPTION]
      : r.description;

    const arr = new Array(CONFIG.WRITE_COLS).fill("");
    arr[CONFIG.COL_DATE - 1]        = Utilities.formatDate(r.startDate, tz, "dd/MM/yyyy");
    arr[CONFIG.COL_TITLE - 1]       = r.title;
    arr[CONFIG.COL_ADDRESS - 1]     = address;
    arr[CONFIG.COL_DESCRIPTION - 1] = description;
    arr[CONFIG.COL_EVENT_ID - 1]    = r.eventId;
    rows.push(arr);
    rowMeta.push({ type: "event", colorHex: r.colorHex });
  }

  if (rows.length === 0 && oldDataRows > 0) {
    Logger.log(`[${sheetName}] ABORT : 0 evenement recu du Calendar mais la feuille avait ${oldDataRows} lignes. On ne vide pas.`);
    _alertSafe_(
      "Feuille \"" + sheetName + "\" protegee (0 evenement)",
      "Le sync a recu 0 evenement pour cette feuille alors qu'elle contenait deja " + oldDataRows + " ligne(s).\n" +
      "Pour eviter de tout vider a cause d'un probleme cote Google, cette feuille n'a PAS ete touchee.\n" +
      "Verifiez le Calendar et relancez la sync."
    );
    return;
  }

  const startRow = CONFIG.HEADER_ROWS + 1;

  if (rows.length > 0) {
    const dataRange = sheet.getRange(startRow, 1, rows.length, CONFIG.WRITE_COLS);
    dataRange.setValues(rows);
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    dataRange.clearFormat();
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  if (oldDataRows > rows.length) {
    const extraStart = startRow + rows.length;
    const extraCount = oldDataRows - rows.length;
    const extraRange = sheet.getRange(extraStart, 1, extraCount, CONFIG.WRITE_COLS);
    extraRange.clearContent();
    extraRange.clearFormat();
    extraRange.setBackground(null);
    sheet.setRowHeights(extraStart, extraCount, CONFIG.ROW_HEIGHT_NORMAL);
  }

  if (rows.length === 0) {
    normalizeRowHeights_(sheet, startRow, CONFIG.ROW_HEIGHT_NORMAL);
    return;
  }

  for (let i = 0; i < rowMeta.length; i++) {
    const meta = rowMeta[i];
    const rowIndex = startRow + i;

    if (meta.type === "month") {
      sheet.setRowHeight(rowIndex, CONFIG.ROW_HEIGHT_MONTH);
      sheet.getRange(rowIndex, 1, 1, CONFIG.WRITE_COLS)
        .setBackground(CONFIG.MONTH_ROW_BG)
        .setFontColor(CONFIG.MONTH_ROW_FG)
        .setFontWeight("bold");
    } else if (meta.type === "week") {
      sheet.setRowHeight(rowIndex, CONFIG.ROW_HEIGHT_WEEK);
      sheet.getRange(rowIndex, 1, 1, CONFIG.WRITE_COLS)
        .setBackground(CONFIG.WEEK_ROW_BG);
    } else {
      sheet.setRowHeight(rowIndex, CONFIG.ROW_HEIGHT_NORMAL);
      sheet.getRange(rowIndex, CONFIG.COL_TITLE)
        .setBackground(meta.colorHex || "#bdbdbd");
    }
  }

  restoreManualData_(sheet, sheetName, savedManual, backupMap);
  setManualColsAsText_(sheet);
  pushToCalendar_(sheet, originalDataMap);

  try { sheet.setFrozenRows(CONFIG.HEADER_ROWS); } catch (e) {}
}

/** ================== RESET / NORMALIZE ================== */
function resetSheetKeepHeader_(sheet, headerRows, writeCols) {
  const lastRow = sheet.getLastRow();
  const dataStart = headerRows + 1;
  if (lastRow < dataStart) return;
  const numRows = lastRow - headerRows;
  const clearRange = sheet.getRange(dataStart, 1, numRows, writeCols);
  clearRange.clearContent();
  clearRange.clearFormat();
  clearRange.setBackground(null);
  sheet.setRowHeights(dataStart, numRows, CONFIG.ROW_HEIGHT_NORMAL);
}

function normalizeRowHeights_(sheet, startRow, height) {
  const last = sheet.getLastRow();
  if (last < startRow) return;
  sheet.setRowHeights(startRow, last - startRow + 1, height);
}

/** ================== DATES / LABELS ================== */
function isoToDate_(isoDate, tz) {
  const [y, m, d] = isoDate.split("-").map(n => parseInt(n, 10));
  return new Date(y, m - 1, d, 12, 0, 0);
}

function monthLabelFR_(dateObj, tz) {
  const m = parseInt(Utilities.formatDate(dateObj, tz, "M"), 10);
  const y = Utilities.formatDate(dateObj, tz, "yyyy");
  const months = ["janvier","fevrier","mars","avril","mai","juin","juillet","aout","septembre","octobre","novembre","decembre"];
  return `${months[m - 1]} ${y}`;
}

function getTimeZone_() {
  return Session.getScriptTimeZone() || "America/Toronto";
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID && String(CONFIG.SPREADSHEET_ID).trim()) {
    return SpreadsheetApp.openById(String(CONFIG.SPREADSHEET_ID).trim());
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** ================== RAPPORT SANS FACTURE ================== */

function envoyerRapportSansFacture() {
  const ui = SpreadsheetApp.getUi();

  const choix = ui.alert(
    "Envoyer le rapport a qui?",
    "OUI  = Juste moi (" + RAPPORT_CONFIG.EMAIL_MOI + ")\n" +
    "NON  = Toute l'equipe (" + RAPPORT_CONFIG.EMAIL_EQUIPE + ")",
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (choix === ui.Button.CANCEL) return;

  const destinataire = (choix === ui.Button.YES)
    ? RAPPORT_CONFIG.EMAIL_MOI
    : RAPPORT_CONFIG.EMAIL_EQUIPE;

  const travaux = _rapport_getTravausSansFacture_();

  if (travaux.length === 0) {
    ui.alert("Aucun travail sans facture trouve! Tout est a jour.");
    return;
  }

  // NOUVEAU V3 — split pour le sujet du courriel aussi
  const manquants    = travaux.filter(t => t.appel === "");
  const enTraitement = travaux.filter(t => t.appel !== "");

  const html   = _rapport_construireCourriel_(travaux);
  const sujet  = "SARAN ELECTRIQUE — " + manquants.length + " facture(s) manquante(s)"
               + (enTraitement.length > 0 ? " / " + enTraitement.length + " en traitement" : "");

  GmailApp.sendEmail(destinataire, sujet, "", { htmlBody: html });
  ui.alert("Rapport envoye a " + destinataire + "\n(" + manquants.length + " manquantes, " + enTraitement.length + " en traitement)");
}

function _rapport_getTravausSansFacture_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTimeZone_();
  const aujourd_hui = new Date();
  aujourd_hui.setHours(0, 0, 0, 0);

  const nomsFeuilles = CONFIG.ROUTING_RULES.map(r => r.sheet).concat([CONFIG.DEFAULT_SHEET]);
  const resultats = [];

  for (const nomFeuille of nomsFeuilles) {
    const feuille = ss.getSheetByName(nomFeuille);
    if (!feuille) continue;

    const derniereLigne = feuille.getLastRow();
    const ligneDebut    = RAPPORT_CONFIG.LIGNE_DEBUT;
    if (derniereLigne < ligneDebut) continue;

    const nbLignes = derniereLigne - ligneDebut + 1;

    const datesRange    = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_DATE,    nbLignes, 1).getValues();
    const projetsRange  = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_PROJET,  nbLignes, 1).getValues();
    const appelsRange   = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_APPEL,   nbLignes, 1).getValues();
    const otRange       = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_OT,      nbLignes, 1).getValues();
    const facturesRange = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_FACTURE, nbLignes, 1).getValues();
    const couleursRange = feuille.getRange(ligneDebut, RAPPORT_CONFIG.COL_PROJET,  nbLignes, 1).getBackgrounds();

    for (let i = 0; i < nbLignes; i++) {
      const dateVal = datesRange[i][0];
      const projet  = String(projetsRange[i][0]).trim();
      const appel   = String(appelsRange[i][0]).trim();
      const ot      = String(otRange[i][0]).trim();
      const facture = String(facturesRange[i][0]).trim();
      const couleur = couleursRange[i][0];

      if (!projet) continue;
      if (facture !== "" && facture !== "0") continue;

      if (dateVal instanceof Date) {
        const d = new Date(dateVal);
        d.setHours(0, 0, 0, 0);
        if (d >= aujourd_hui) continue;
      } else {
        const parts = String(dateVal).split("/");
        if (parts.length === 3) {
          const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          if (d >= aujourd_hui) continue;
        } else {
          continue;
        }
      }

      resultats.push({
        feuille: nomFeuille,
        date:    dateVal,
        projet:  projet,
        appel:   appel,
        ot:      ot,
        couleur: (couleur && couleur !== "#ffffff" && couleur !== "#000000" && couleur !== "white")
                   ? couleur : "#039be5",
      });
    }
  }

  resultats.sort((a, b) => {
    if (a.feuille < b.feuille) return -1;
    if (a.feuille > b.feuille) return 1;
    return _rapport_parseDate_(a.date) - _rapport_parseDate_(b.date);
  });

  return resultats;
}

function _rapport_parseDate_(val) {
  if (val instanceof Date) return val.getTime();
  const parts = String(val).split("/");
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0])).getTime();
  }
  return 0;
}

function _rapport_formatDate_(val) {
  const tz = getTimeZone_();
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "EEE d MMM yyyy");
  }
  const parts = String(val).split("/");
  if (parts.length === 3) {
    const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    return Utilities.formatDate(d, tz, "EEE d MMM yyyy");
  }
  return String(val);
}

// NOUVEAU V3 — construit les lignes de tableau pour UNE liste (deja filtree par statut),
// groupees par feuille, exactement comme avant.
function _rapport_blocParFeuille_(liste) {
  const parFeuille = {};
  for (const t of liste) {
    if (!parFeuille[t.feuille]) parFeuille[t.feuille] = [];
    parFeuille[t.feuille].push(t);
  }

  let out = "";
  for (const nomFeuille of Object.keys(parFeuille)) {
    const items = parFeuille[nomFeuille];

    out += `
      <tr>
        <td colspan="4" style="background:#1a1a2e;color:#ffffff;font-weight:bold;
          font-size:13px;padding:10px 14px;letter-spacing:0.5px;">
          ${nomFeuille} &mdash; ${items.length} travaux
        </td>
      </tr>`;

    for (const t of items) {
      const dateStr = _rapport_formatDate_(t.date);

      const pastille = '<span style="display:inline-block;width:12px;height:12px;' +
        'border-radius:50%;background:' + t.couleur + ';margin-right:7px;' +
        'vertical-align:middle;border:1px solid rgba(0,0,0,0.15);"></span>';

      const celluleAppel = t.appel !== ""
        ? t.appel
        : '<span style="color:#bbb;font-style:italic;">Aucun appel</span>';

      const celluleFacture = t.appel !== ""
        ? '<span style="background:#fff8e1;color:#e65100;border:1px solid #ffe082;' +
          'border-radius:4px;padding:2px 8px;font-size:11px;font-weight:bold;">En traitement</span>'
        : '<span style="background:#fff0f0;color:#c0392b;border:1px solid #f5c6cb;' +
          'border-radius:4px;padding:2px 8px;font-size:11px;font-weight:bold;">Manquante</span>';

      out += `
        <tr style="border-bottom:1px solid #e8e8e8;">
          <td style="padding:9px 14px;vertical-align:middle;width:140px;color:#555;font-size:12px;">
            ${dateStr}
          </td>
          <td style="padding:9px 14px;vertical-align:middle;font-weight:600;font-size:13px;">
            ${pastille}${t.projet}
          </td>
          <td style="padding:9px 14px;vertical-align:middle;color:#333;font-size:12px;">
            ${celluleAppel}
          </td>
          <td style="padding:9px 14px;vertical-align:middle;text-align:center;">
            ${celluleFacture}
          </td>
        </tr>`;
    }
  }
  return out;
}

function _rapport_construireCourriel_(travaux) {
  const tz = getTimeZone_();
  const maintenant = Utilities.formatDate(new Date(), tz, "d MMMM yyyy 'a' HH'h'mm");

  // NOUVEAU V3 — separation en deux groupes : manquantes (vraiment rien) vs en traitement
  const manquants    = travaux.filter(t => t.appel === "");
  const enTraitement = travaux.filter(t => t.appel !== "");

  const enteteSection = (titre, couleurFond, nb) => `
    <tr>
      <td colspan="4" style="background:${couleurFond};color:#ffffff;font-weight:bold;
        font-size:14px;padding:14px 16px;letter-spacing:0.5px;">
        ${titre} &mdash; ${nb}
      </td>
    </tr>`;

  let corpsTableau = "";

  // SECTION 1 (en haut) — vraiment manquantes, aucun appel de service
  corpsTableau += enteteSection("FACTURES MANQUANTES — aucun appel de service", "#c0392b", manquants.length);
  corpsTableau += manquants.length > 0
    ? _rapport_blocParFeuille_(manquants)
    : `<tr><td colspan="4" style="padding:14px 16px;color:#888;font-style:italic;">Aucune — tout est a jour de ce cote.</td></tr>`;

  // SECTION 2 (a la fin, separee) — appel de service deja fait, facture a venir
  corpsTableau += enteteSection("EN TRAITEMENT — appel de service deja fait, facture a venir", "#e67e22", enTraitement.length);
  corpsTableau += enTraitement.length > 0
    ? _rapport_blocParFeuille_(enTraitement)
    : `<tr><td colspan="4" style="padding:14px 16px;color:#888;font-style:italic;">Aucun.</td></tr>`;

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">' +
    '<div style="max-width:700px;margin:30px auto;background:#fff;border-radius:8px;' +
      'overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">' +

    '<div style="background:#c0392b;padding:24px 28px;">' +
      '<div style="color:#fff;font-size:20px;font-weight:bold;">Saran Electrique</div>' +
      '<div style="color:#f5b7b1;font-size:13px;margin-top:4px;">Rapport des travaux sans facture</div>' +
    '</div>' +

    '<div style="background:#fff8f8;padding:16px 28px;border-bottom:2px solid #f5c6cb;">' +
      '<span style="font-size:15px;font-weight:bold;color:#c0392b;">' +
        manquants.length + ' facture(s) manquante(s)' +
      '</span>' +
      '<span style="font-size:13px;color:#e67e22;margin-left:12px;">' +
        '+ ' + enTraitement.length + ' en traitement' +
      '</span>' +
      '<span style="color:#888;font-size:12px;margin-left:16px;">Genere le ' + maintenant + '</span>' +
    '</div>' +

    '<div style="padding:20px 28px;">' +
      '<table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;">' +
        '<thead>' +
          '<tr style="background:#f8f8f8;border-bottom:2px solid #ddd;">' +
            '<th style="padding:10px 14px;text-align:left;font-size:12px;color:#666;font-weight:600;width:140px;">DATE</th>' +
            '<th style="padding:10px 14px;text-align:left;font-size:12px;color:#666;font-weight:600;">PROJET / EVENEMENT</th>' +
            '<th style="padding:10px 14px;text-align:left;font-size:12px;color:#666;font-weight:600;">APPEL DE SERVICE (K)</th>' +
            '<th style="padding:10px 14px;text-align:center;font-size:12px;color:#666;font-weight:600;">FACTURE (M)</th>' +
          '</tr>' +
        '</thead>' +
        '<tbody>' +
          corpsTableau +
        '</tbody>' +
      '</table>' +
    '</div>' +

    '<div style="background:#f8f8f8;padding:14px 28px;border-top:1px solid #eee;text-align:center;">' +
      '<span style="color:#aaa;font-size:11px;">Saran Electrique &middot; info@saranelectrique.com &middot; Saint-Jean-sur-Richelieu</span>' +
    '</div>' +

    '</div></body></html>';
}
