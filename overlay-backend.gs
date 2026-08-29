/**
 * Backend for the "our notes" overlay on the EV comparison page.
 *
 * Deploy this bound to a Google Sheet with a tab named "Overlay" whose
 * first row is exactly:
 *   key | eliminated | rating | note | updatedAt
 *
 * Setup (one-time, in your own Google account):
 * 1. Create a new Google Sheet. Rename the first tab to "Overlay".
 * 2. Add the header row above (row 1, columns A-E, exact spelling).
 * 3. Extensions -> Apps Script. Delete the placeholder code and paste
 *    this whole file in.
 * 4. Deploy -> New deployment -> type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Authorize when prompted (it's your own script touching your own
 *    sheet - the scary-looking consent screen is normal for a
 *    self-authored Apps Script project).
 * 6. Copy the resulting web app URL (ends in /exec) and paste it into
 *    the EDITOR_ENDPOINT constant in ev-crossovers.html.
 *
 * Re-deploying after an edit to this file: Deploy -> Manage deployments
 * -> pencil icon -> New version -> Deploy. Editing the script without
 * creating a new version will NOT update the live /exec URL.
 */

const SHEET_NAME = 'Overlay';
const HEADERS = ['key', 'eliminated', 'rating', 'note', 'updatedAt'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const rows = values
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  return json_({ ok: true, rows });
}

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
    const sheet = getSheet_();
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
