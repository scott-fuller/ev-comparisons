/**
 * Backend for the EV comparison page (ev-comparisons repo): serves the
 * editable primary table data (Trims + Used tabs) AND the shared "our
 * notes" overlay (Overlay tab) from one Google Sheet + one Apps Script
 * web app deployment.
 *
 * Deploy this bound to a Google Sheet with three tabs:
 *
 * 1. "Trims" — one row per trim. Header row (row 1, columns A-Q, exact spelling):
 *    modelId | make | model | year | warrantyBasic | warrantyBattery | trimName |
 *    msrp | drivetrain | range | mpge | batteryKwh | dcfcKw | zeroToSixty |
 *    cargo | comfort | priceFlag
 *
 *    - Group all trims for the same model on consecutive rows (same modelId) —
 *      the first row for a modelId supplies make/model/year/warranty for the
 *      whole model, so only its first trim row strictly needs those filled in
 *      (later rows for the same modelId may repeat them or leave them blank).
 *    - Leave a cell blank for "not confirmed" (shows as "—" on the page,
 *      matching this page's own sourcing convention — never guess a value).
 *    - comfort: multiple short feature notes separated by " | " (pipe),
 *      e.g. "Heated front seats standard | 10-way power driver seat"
 *    - priceFlag: TRUE to show the † data-conflict flag next to that trim.
 *    - drivetrain: FWD / RWD / AWD. year: a 4-digit number (e.g. 2026).
 *
 * 2. "Used" — one row per model on the Used '25 tab. Header row (columns A-F):
 *    modelId | usedPrice | sales2025 | lowAvail | availNote | trimNameOverride
 *
 *    - modelId must match a modelId used in the Trims tab.
 *    - lowAvail: TRUE/FALSE. sales2025 / availNote / trimNameOverride may be
 *      left blank.
 *
 * 3. "Overlay" — unchanged from before: key | eliminated | rating | note | updatedAt
 *    (the page manages this tab entirely itself via star ratings / notes /
 *    eliminate buttons — you don't need to type into it).
 *
 * If the Trims or Used tabs are empty (or unreachable), the live page falls
 * back to its own built-in seed data automatically — editing the Sheet is
 * optional, not required for the page to work.
 *
 * Setup (one-time, in your own Google account):
 * 1. Create a new Google Sheet (or reuse the one already backing "our notes").
 * 2. Add tabs named exactly "Trims", "Used", and "Overlay" (case-sensitive).
 *    Missing tabs are auto-created with the correct header row the first
 *    time this script's web app receives a GET request, but the sheet
 *    itself must already exist for that to happen — visiting the /exec URL
 *    once after deploying is enough to trigger it.
 * 3. Paste your trim/used data below each header row (see the repo's
 *    sheet-data/trims.csv and sheet-data/used.csv for the current page
 *    data, ready to paste in).
 * 4. Extensions -> Apps Script. Delete the placeholder code and paste this
 *    whole file in.
 * 5. Deploy -> New deployment -> type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 6. Authorize when prompted (it's your own script touching your own
 *    sheet — the scary-looking consent screen is normal for a
 *    self-authored Apps Script project).
 * 7. Copy the resulting web app URL (ends in /exec) and paste it into the
 *    DATA_ENDPOINT constant in index.html (it already doubles as the
 *    overlay-notes endpoint, so there's only one URL to update).
 *
 * Re-deploying after an edit to this file: Deploy -> Manage deployments
 * -> pencil icon -> New version -> Deploy. Editing the script without
 * creating a new version will NOT update the live /exec URL.
 *
 * Editing the Trims/Used tab contents (prices, specs, etc.) does NOT need
 * a redeploy — that data is read live on every page load.
 */

const OVERLAY_SHEET = 'Overlay';
const OVERLAY_HEADERS = ['key', 'eliminated', 'rating', 'note', 'updatedAt'];

const TRIMS_SHEET = 'Trims';
const TRIMS_HEADERS = ['modelId', 'make', 'model', 'year', 'warrantyBasic', 'warrantyBattery',
  'trimName', 'msrp', 'drivetrain', 'range', 'mpge', 'batteryKwh', 'dcfcKw', 'zeroToSixty',
  'cargo', 'comfort', 'priceFlag'];

const USED_SHEET = 'Used';
const USED_HEADERS = ['modelId', 'usedPrice', 'sales2025', 'lowAvail', 'availNote', 'trimNameOverride'];

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function readRows_(name, headers) {
  const sheet = getOrCreateSheet_(name, headers);
  const values = sheet.getDataRange().getValues();
  const sheetHeaders = values.shift();
  return values
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      sheetHeaders.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const rows = readRows_(OVERLAY_SHEET, OVERLAY_HEADERS);
  const trims = readRows_(TRIMS_SHEET, TRIMS_HEADERS);
  const used = readRows_(USED_SHEET, USED_HEADERS);
  return json_({ ok: true, rows, trims, used });
}

// POST only ever writes to the Overlay tab (star ratings, notes, eliminations
// from the page). Trims/Used are edited directly in the Sheet by hand.
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad JSON body' });
  }
  if (!payload || typeof payload.key !== 'string' || !payload.key) {
    return json_({ ok: false, error: 'Missing key' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getOrCreateSheet_(OVERLAY_SHEET, OVERLAY_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const keyCol = headers.indexOf('key');

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][keyCol] === payload.key) { rowIndex = i; break; }
    }

    const existing = rowIndex === -1 ? {} : headers.reduce((acc, h, i) => {
      acc[h] = values[rowIndex][i];
      return acc;
    }, {});

    const merged = Object.assign({}, existing, payload, {
      updatedAt: new Date().toISOString(),
    });
    const newRow = headers.map(h => (merged[h] !== undefined ? merged[h] : ''));

    if (rowIndex === -1) {
      sheet.appendRow(newRow);
    } else {
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([newRow]);
    }

    return json_({ ok: true, row: merged });
  } finally {
    lock.releaseLock();
  }
}
