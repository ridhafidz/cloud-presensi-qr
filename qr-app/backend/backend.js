const SPREADSHEET_ID = "1pp0VRB_VVo4SZFCRvNHEJwX0BOtiy5P3-MY7tpqDYo0";

function doGet(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";

  if (path === "") {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dashboard Dosen - Presensi QR')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (path === "presence/status") return handleGetStatus(e);

  if (path === "sensor/gps/marker") return handleGetGPSMarker(e);
  if (path === "sensor/gps/polyline") return handleGetGPSPolyline(e);

  return sendError("Route not found");
}

function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  const body = JSON.parse(e.postData.contents);

  if (path === "presence/qr/generate") return handleGenerateQR(body);
  if (path === "presence/checkin") return handleCheckin(body);

  if (path === "sensor/accel/batch") return handleAccelBatch(body);
  if (path === "sensor/gps") return handleGPS(body);

  return sendError("Route not found");
}

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

function handleGenerateQR(body) {
  if (!body.course_id || !body.session_id) return sendError("missing_field");

  const token = "TKN-" + Utilities.getUuid().substring(0,6).toUpperCase();
  const now = new Date();
  const expires = new Date(now.getTime() + 120000); // 120 detik

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("tokens");
  // Kolom: qr_token, course_id, session_id, created_at, expires_at, used
  sheet.appendRow([token, body.course_id, body.session_id, now.toISOString(), expires.toISOString(), false]);

  return sendSuccess({ qr_token: token, expires_at: expires.toISOString() });
}

function processGenerateQR(body) {
  if (!body.course_id || !body.session_id) return { ok: false, error: "missing_field" };

  const token = "TKN-" + Utilities.getUuid().substring(0,6).toUpperCase();
  const now = new Date();
  const expires = new Date(now.getTime() + 120000); 

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("tokens");
  sheet.appendRow([token, body.course_id, body.session_id, now.toISOString(), expires.toISOString(), false]);

  return { ok: true, data: { qr_token: token, expires_at: expires.toISOString() } };
}

function handleCheckin(body) {
  if (!body.user_id || !body.qr_token || !body.course_id || !body.session_id) return sendError("missing_field");

  const sheetTokens = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("tokens");
  const dataTokens = sheetTokens.getDataRange().getValues();
  
  let tokenValid = false;
  let tokenExists = false;

  for (let i = dataTokens.length - 1; i > 0; i--) {
    if (dataTokens[i][0] === body.qr_token) {
      tokenExists = true;
      const expiresAt = new Date(dataTokens[i][4]);
      if (dataTokens[i][1] === body.course_id && dataTokens[i][2] === body.session_id && new Date() <= expiresAt) {
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

  sheetPresence.appendRow([
    presenceId, body.user_id, body.device_id || "UNKNOWN", body.course_id, 
    body.session_id, body.qr_token, body.ts || recordedAt, recordedAt
  ]);

  return sendSuccess({ presence_id: presenceId, status: "checked_in" });
}

function handleGetStatus(e) {
  const { user_id, course_id, session_id } = e.parameter;
  if (!user_id || !course_id || !session_id) return sendError("missing_parameter");

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("presence");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === user_id && data[i][3] === course_id && data[i][4] === session_id) {
      return sendSuccess({
        user_id: user_id,
        course_id: course_id,
        session_id: session_id,
        status: "checked_in",
        last_ts: data[i][7] 
      });
    }
  }
  return sendError("not_checked_in");
}

function handleAccelBatch(body) {
  if (!body.device_id || !body.data || !Array.isArray(body.data)) return sendError("invalid_batch_data");

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("accel");
  const batch_ts = new Date().toISOString();
  
  const rows = body.data.map(d => [
    body.device_id, d.x, d.y, d.z, d.ts, batch_ts, batch_ts
  ]);

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return sendSuccess({ saved_records: rows.length });
}

function handleGPS(body) {
  if (!body.device_id || !body.lat || !body.lng) return sendError("missing_gps_data");

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("gps");
  const recordedAt = new Date().toISOString();
  
  sheet.appendRow([
    body.device_id, body.lat, body.lng, 
    body.accuracy || 0, body.altitude || 0, 
    body.ts || recordedAt, recordedAt
  ]);

  return sendSuccess({ status: "recorded", timestamp: recordedAt });
}

function handleGetGPSMarker(e) {
  const deviceId = e.parameter.device_id;
  if (!deviceId) return sendError("missing_device_id");

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("gps");
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][0] === deviceId) {
      return sendSuccess({
        device_id: deviceId,
        lat: data[i][1],
        lng: data[i][2],
        ts: data[i][6]
      });
    }
  }
  return sendError("data_not_found");
}

function handleGetGPSPolyline(e) {
  const deviceId = e.parameter.device_id;
  const fromDate = e.parameter.from ? new Date(e.parameter.from) : null;
  const toDate = e.parameter.to ? new Date(e.parameter.to) : null;
  
  if (!deviceId) return sendError("missing_device_id");

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("gps");
  const data = sheet.getDataRange().getValues();
  
  let route = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === deviceId) {
      const recordTime = new Date(data[i][6]);
      
      if (fromDate && recordTime < fromDate) continue;
      if (toDate && recordTime > toDate) continue;

      route.push({
        lat: data[i][1],
        lng: data[i][2],
        ts: data[i][6]
      });
    }
  }

  return sendSuccess({
    device_id: deviceId,
    total_points: route.length,
    polyline: route
  });
}