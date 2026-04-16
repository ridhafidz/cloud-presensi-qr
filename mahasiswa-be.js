// ============================================================
// GAS MAHASISWA - gas-mahasiswa.js
// Deploy ini sebagai Google Apps Script TERPISAH (Web App #2)
// Spreadsheet: khusus mahasiswa, harus ada sheet:
//   - "presence"
//   - "accel"
//   - "gps"
// ============================================================

const SPREADSHEET_MAHASISWA = "1A-PRMkgt7YbXCelkrrfSLMBHh1fdzACXiM4gy4trONM";

function doGet(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";

  if (path === "presence/status")      return handleGetStatus(e);
  if (path === "sensor/accel/latest")  return handleGetAccelLatest(e);
  if (path === "sensor/gps/marker")    return handleGetGPSMarker(e);

  return sendError("Route GET tidak ditemukan di GAS Mahasiswa");
}

function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  let body = {};

  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return sendError("Invalid JSON Body");
  }

  if (path === "presence/checkin")   return handleCheckin(body);
  if (path === "sensor/accel/batch") return handleAccelBatch(body);
  if (path === "sensor/gps")         return handleGPS(body);

  return sendError("Route POST tidak ditemukan di GAS Mahasiswa");
}

// --- Checkin Mahasiswa ---
// body sudah berisi hasil validasi token dari Bridge (token sudah dicek ke GAS Dosen duluan)
function handleCheckin(body) {
  if (!body.user_id || !body.device_id || !body.course_id || !body.session_id || !body.token) {
    return sendError("missing_field: user_id, device_id, course_id, session_id, token wajib ada");
  }

  const sheetPresence = SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("presence");

  const dataPresence = sheetPresence.getDataRange().getValues();

  // Cek duplikat checkin
  for (let i = dataPresence.length - 1; i > 0; i--) {
    if (
      dataPresence[i][1] === body.user_id &&
      dataPresence[i][3] === body.course_id &&
      dataPresence[i][4] === body.session_id
    ) {
      return sendError("user_already_checked_in");
    }
  }

  const presenceId = "PR-" + Utilities.getUuid().substring(0, 6).toUpperCase();
  const now = new Date().toISOString();

  sheetPresence.appendRow([
    presenceId,
    body.user_id,
    body.device_id,
    body.course_id,
    body.session_id,
    body.token,
    now,
    now
  ]);

  return sendSuccess({ presence_id: presenceId, status: "checked_in" });
}

// --- Cek Status Presensi ---
function handleGetStatus(e) {
  const { user_id, course_id, session_id } = e.parameter;
  const data = SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("presence")
    .getDataRange()
    .getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (
      data[i][1] === user_id &&
      data[i][3] === course_id &&
      data[i][4] === session_id
    ) {
      return sendSuccess({
        user_id,
        course_id,
        session_id,
        status: "checked_in",
        last_ts: data[i][7]
      });
    }
  }

  return sendError("not_checked_in");
}

// --- Simpan Data Accelerometer (batch) ---
function handleAccelBatch(body) {
  if (!body.device_id || !Array.isArray(body.data)) {
    return sendError("invalid_batch_data");
  }

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("accel");

  const now = new Date().toISOString();

  const rows = body.data.map(d => [
    body.device_id,
    d.x,
    d.y,
    d.z,
    d.ts,
    now,
    now
  ]);

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  return sendSuccess({ saved_records: rows.length });
}

// --- Simpan Data GPS ---
function handleGPS(body) {
  if (!body.device_id || !body.lat || !body.lng) {
    return sendError("missing_gps_data");
  }

  const now = new Date().toISOString();

  SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("gps")
    .appendRow([
      body.device_id,
      body.lat,
      body.lng,
      body.accuracy || 0,
      body.altitude || 0,
      body.ts || now,
      now
    ]);

  return sendSuccess({ status: "recorded" });
}

// --- Ambil Data Accel Terbaru (untuk dosen pantau) ---
function handleGetAccelLatest(e) {
  if (!e.parameter.nim) return sendError("missing_nim");

  const data = SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("accel")
    .getDataRange()
    .getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim)) {
      return sendSuccess({
        x: data[i][1],
        y: data[i][2],
        z: data[i][3],
        ts: data[i][6]
      });
    }
  }

  return sendError("data_not_found");
}

// --- Ambil Data GPS Terbaru (untuk dosen pantau) ---
function handleGetGPSMarker(e) {
  if (!e.parameter.nim) return sendError("missing_nim");

  const data = SpreadsheetApp
    .openById(SPREADSHEET_MAHASISWA)
    .getSheetByName("gps")
    .getDataRange()
    .getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim)) {
      return sendSuccess({
        lat: data[i][1],
        lng: data[i][2],
        ts: data[i][6]
      });
    }
  }

  return sendError("data_not_found");
}

// --- Helper Response ---
function sendSuccess(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendError(error) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error }))
    .setMimeType(ContentService.MimeType.JSON);
}