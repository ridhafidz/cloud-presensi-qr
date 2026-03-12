const SPREADSHEET_ID = "1pp0VRB_VVo4SZFCRvNHEJwX0BOtiy5P3-MY7tpqDYo0";

function doGet(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";

  if (path === "presence/status") return handleGetStatus(e);
  if (path === "sensor/accel/latest") return handleGetAccelLatest(e);
  if (path === "sensor/gps/marker") return handleGetGPSMarker(e);
  if (path === "sensor/gps/polyline") return handleGetGPSPolyline(e);

  return sendError("Route GET not found");
}

function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  let body = {};
  
  try {
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return sendError("Invalid JSON Body");
  }

  if (path === "presence/qr/generate") return handleGenerateQR(body);
  if (path === "presence/checkin") return handleCheckin(body);
  if (path === "sensor/accel/batch") return handleAccelBatch(body);
  if (path === "sensor/gps") return handleGPS(body);

  return sendError("Route POST not found");
}

function sendSuccess(data) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data })).setMimeType(ContentService.MimeType.JSON);
}

function sendError(error) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error })).setMimeType(ContentService.MimeType.JSON);
}

// --- FUNGSI MODUL PRESENSI ---
function handleGenerateQR(body) {
  if (!body.course_id || !body.session_id) return sendError("missing_field");
  const token = "TKN-" + Utilities.getUuid().substring(0,6).toUpperCase();
  const now = new Date();
  const expires = new Date(now.getTime() + 120000); 

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("tokens");
  sheet.appendRow([token, body.course_id, body.session_id, now.toISOString(), expires.toISOString(), false]);
  return sendSuccess({ qr_token: token, expires_at: expires.toISOString() });
}

function handleCheckin(body) {
  if (!body.user_id || !body.qr_token || !body.course_id || !body.session_id) return sendError("missing_field");

  const sheetTokens = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("tokens");
  const dataTokens = sheetTokens.getDataRange().getValues();
  let tokenValid = false, tokenExists = false;

  for (let i = dataTokens.length - 1; i > 0; i--) {
    if (dataTokens[i][0] === body.qr_token) {
      tokenExists = true;
      if (dataTokens[i][1] === body.course_id && dataTokens[i][2] === body.session_id && new Date() <= new Date(dataTokens[i][4])) {
        tokenValid = true;
      }
      break; 
    }
  }

  if (!tokenExists) return sendError("token_invalid");
  if (!tokenValid) return sendError("token_expired");

  const sheetPresence = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("presence");
  const presenceId = "PR-" + Utilities.getUuid().substring(0, 4).toUpperCase();
  const recordedAt = new Date().toISOString();

  sheetPresence.appendRow([presenceId, body.user_id, body.device_id || "UNKNOWN", body.course_id, body.session_id, body.qr_token, body.ts || recordedAt, recordedAt]);
  return sendSuccess({ presence_id: presenceId, status: "checked_in" });
}

function handleGetStatus(e) {
  const { user_id, course_id, session_id } = e.parameter;
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("presence").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === user_id && data[i][3] === course_id && data[i][4] === session_id) {
      return sendSuccess({ user_id, course_id, session_id, status: "checked_in", last_ts: data[i][7] });
    }
  }
  return sendError("not_checked_in");
}

// --- FUNGSI MODUL SENSOR ---
function handleAccelBatch(body) {
  if (!body.device_id || !body.data || !Array.isArray(body.data)) return sendError("invalid_batch_data");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("accel");
  const batch_ts = new Date().toISOString();
  const rows = body.data.map(d => [body.device_id, d.x, d.y, d.z, d.ts, batch_ts, batch_ts]);
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return sendSuccess({ saved_records: rows.length });
}

function handleGPS(body) {
  if (!body.device_id || !body.lat || !body.lng) return sendError("missing_gps_data");
  const ts = new Date().toISOString();
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("gps").appendRow([body.device_id, body.lat, body.lng, body.accuracy || 0, body.altitude || 0, body.ts || ts, ts]);
  return sendSuccess({ status: "recorded", timestamp: ts });
}

function handleGetAccelLatest(e) {
  if (!e.parameter.nim) return sendError("missing_nim");
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("accel").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim)) {
      return sendSuccess({ 
        x: data[i][1], 
        y: data[i][2], 
        z: data[i][3], 
        ts: data[i][6] || data[i][5] 
      });
    }
  }
  return sendError("data_not_found");
}

function handleGetGPSMarker(e) {
  if (!e.parameter.nim) return sendError("missing_nim");
  const data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("gps").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim)) return sendSuccess({ lat: data[i][1], lng: data[i][2], ts: data[i][6] });
  }
  return sendError("data_not_found");
}