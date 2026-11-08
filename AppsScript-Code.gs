/**
 * fieldlog — Google Apps Script backend (v2: food + body)
 * Bridges the fieldlog app <-> your Google Sheet.
 *
 * v2 adds waist/height ("Body") tracking in a separate tab.
 *
 * If you already deployed v1: paste this over the old code, Save, then
 * Deploy > Manage deployments > (edit, pencil icon) > Version: New version > Deploy.
 * Your Web app URL stays the same.
 */

const SHEET_NAME = 'Log';
const HEADER = ['date', 'name', 'cal', 'protein', 'fiber', 'iron', 'plant', 'id', 'logged_at'];

const BODY_NAME = 'Body';
const BODY_HEADER = ['date', 'waist_in', 'height_in', 'id', 'logged_at'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  const firstRow = sheet.getRange(1, 1, 1, HEADER.length).getValues()[0];
  if (firstRow.join('') !== HEADER.join('')) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  return sheet;
}

function getBodySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BODY_NAME);
  if (!sheet) sheet = ss.insertSheet(BODY_NAME);
  const firstRow = sheet.getRange(1, 1, 1, BODY_HEADER.length).getValues()[0];
  if (firstRow.join('') !== BODY_HEADER.join('')) {
    sheet.getRange(1, 1, 1, BODY_HEADER.length).setValues([BODY_HEADER]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fmtDate(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function doGet() {
  try {
    // --- food entries ---
    const sheet = getSheet();
    const entries = [];
    const last = sheet.getLastRow();
    if (last > 1) {
      const values = sheet.getRange(2, 1, last - 1, HEADER.length).getValues();
      values.forEach(function (r) {
        if (!r[7]) return;
        entries.push({
          date: fmtDate(r[0]), name: String(r[1]),
          cal: Number(r[2]) || 0, protein: Number(r[3]) || 0,
          fiber: Number(r[4]) || 0, iron: Number(r[5]) || 0,
          plant: (r[6] === true || String(r[6]).toLowerCase() === 'true'),
          id: String(r[7])
        });
      });
    }
    // --- body entries ---
    const bsheet = getBodySheet();
    const body = [];
    const blast = bsheet.getLastRow();
    if (blast > 1) {
      const bvals = bsheet.getRange(2, 1, blast - 1, BODY_HEADER.length).getValues();
      bvals.forEach(function (r) {
        if (!r[3]) return;
        body.push({
          date: fmtDate(r[0]),
          waist: Number(r[1]) || 0,
          height: Number(r[2]) || 0,
          id: String(r[3])
        });
      });
    }
    return jsonOut({ ok: true, entries: entries, body: body });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // delete a food entry
    if (body.action === 'delete') {
      const sheet = getSheet();
      const last = sheet.getLastRow();
      if (last > 1) {
        const ids = sheet.getRange(2, 8, last - 1, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(body.id)) sheet.deleteRow(i + 2);
        }
      }
      return jsonOut({ ok: true });
    }

    // add a body (waist/height) measurement
    if (body.action === 'body') {
      const bsheet = getBodySheet();
      bsheet.appendRow([
        String(body.date), Number(body.waist) || 0, Number(body.height) || 0,
        String(body.id), new Date()
      ]);
      return jsonOut({ ok: true });
    }

    // default: add a food entry
    const sheet = getSheet();
    sheet.appendRow([
      String(body.date), String(body.name),
      Number(body.cal) || 0, Number(body.protein) || 0,
      Number(body.fiber) || 0, Number(body.iron) || 0,
      body.plant ? true : false, String(body.id), new Date()
    ]);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
