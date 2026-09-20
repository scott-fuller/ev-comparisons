/**
 * Backend for the EV comparison page (ev-comparisons repo): serves the
 * editable primary table data (Trims + Used tabs), the dealer/test-drive
 * planning data (Dealers tab), AND the shared "our notes" + test-drive
 * status overlay (Overlay tab) from one Google Sheet + one Apps Script
 * web app deployment.
 *
 * Deploy this bound to a Google Sheet with four tabs:
 *
 * 1. "Trims" — one row per trim. Header row (row 1, columns A-S, exact spelling):
 *    modelId | make | model | year | warrantyBasic | warrantyBattery | trimName |
 *    msrp | drivetrain | range | mpge | batteryKwh | dcfcKw | zeroToSixty |
 *    cargo | length | wheelbase | comfort | priceFlag
 *
 *    - Group all trims for the same model on consecutive rows (same modelId) —
 *      the first row for a modelId supplies make/model/year/warranty/length/
 *      wheelbase for the whole model, so only its first trim row strictly
 *      needs those filled in (later rows for the same modelId may repeat
 *      them or leave them blank).
 *    - length / wheelbase: inches, model-generation-level (not per-trim —
 *      these don't vary by trim the way range or MPGe do).
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
 * 3. "Dealers" — one row per dealership, for the "Plan your test drives"
 *    section. Header row (columns A-H):
 *    brand | dealerName | address | phone | rank | distanceNote | specialNote | inventoryUrl
 *
 *    - brand must exactly match a "make" value used in the Trims tab
 *      (e.g. "Kia", "Toyota", "Subaru", "Hyundai", "Volvo", "Ford", "Tesla",
 *      "Nissan", "VinFast", "Chevrolet") — one dealer serves every model
 *      under that brand, so this is brand-level, not per-model.
 *    - rank: 1 = closest, 2 = second-closest. You can add more than 2 rows
 *      per brand (e.g. a 3rd option) — the page just shows all rows for
 *      the selected brand, sorted by rank.
 *    - phone / distanceNote / specialNote may be left blank. specialNote is
 *      for caveats like "Book via the Tesla app, not a phone call" or
 *      "Nearest of only a few US locations — call ahead to confirm stock."
 *    - inventoryUrl: that brand's own official live-inventory-search page
 *      (same value repeated on every row for that brand, same repetition
 *      pattern as make/model/year on the Trims tab). Leave blank if the
 *      brand has no public inventory search.
 *
 * 4. "Overlay" — the page manages this tab entirely itself; you don't need
 *    to type into it. Header row (columns A-G):
 *    key | eliminated | rating | note | status | scheduledDate | updatedAt
 *
 *    Two different kinds of rows share this tab by key prefix:
 *    - "new:<modelId>" / "used:<modelId>" — star rating, note, eliminated
 *      flag from the main table's "Our notes" column (eliminated/rating/note).
 *    - "td:<modelId>" — test-drive status from the "Plan your test drives"
 *      section (status/scheduledDate/note). status is one of "none",
 *      "scheduled", "driven".
 *
 * If the Trims, Used, or Dealers tabs are empty (or unreachable), the live
 * page falls back to its own built-in seed data automatically — editing the
 * Sheet is optional, not required for the page to work.
 *
 * Setup (one-time, in your own Google account):
 * 1. Create a new Google Sheet (or reuse the one already backing "our notes").
 * 2. Add tabs named exactly "Trims", "Used", "Dealers", and "Overlay"
 *    (case-sensitive). Missing tabs are auto-created with the correct header
 *    row the first time this script's web app receives a GET request, but
 *    the sheet itself must already exist for that to happen — visiting the
 *    /exec URL once after deploying is enough to trigger it.
 * 3. Paste your data below each header row (see the repo's sheet-data/*.csv
 *    files for the current page data, ready to paste in).
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
 * Editing the Trims/Used/Dealers tab contents (prices, specs, dealer info,
 * etc.) does NOT need a redeploy — that data is read live on every page load.
 */

const OVERLAY_SHEET = 'Overlay';
const OVERLAY_HEADERS = ['key', 'eliminated', 'rating', 'note', 'status', 'scheduledDate', 'updatedAt'];

const TRIMS_SHEET = 'Trims';
const TRIMS_HEADERS = ['modelId', 'make', 'model', 'year', 'warrantyBasic', 'warrantyBattery',
  'trimName', 'msrp', 'drivetrain', 'range', 'mpge', 'batteryKwh', 'dcfcKw', 'zeroToSixty',
  'cargo', 'length', 'wheelbase', 'comfort', 'priceFlag'];

const USED_SHEET = 'Used';
const USED_HEADERS = ['modelId', 'usedPrice', 'sales2025', 'lowAvail', 'availNote', 'trimNameOverride'];

const DEALERS_SHEET = 'Dealers';
const DEALERS_HEADERS = ['brand', 'dealerName', 'address', 'phone', 'rank', 'distanceNote', 'specialNote', 'inventoryUrl'];

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
  const dealers = readRows_(DEALERS_SHEET, DEALERS_HEADERS);
  return json_({ ok: true, rows, trims, used, dealers });
}

// POST only ever writes to the Overlay tab (star ratings/notes/eliminations
// from the main table, and test-drive status from the planning section).
// Trims/Used/Dealers are edited directly in the Sheet by hand.
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
