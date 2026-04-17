const SPREADSHEET_ID = "1A-PRMkgt7YbXCelkrrfSLMBHh1fdzACXiM4gy4trONM";

function doGet(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  if (path === "presence/status")     return handleGetStatus(e);
  if (path === "sensor/accel/latest") return handleGetAccelLatest(e);
  if (path === "sensor/gps/marker")   return handleGetGPSMarker(e);
  return sendError("Route GET tidak ditemukan");
}

function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  let body = {};
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return sendError("Invalid JSON Body"); }

  // Route Dosen
  if (path === "presence/qr/generate")  return handleGenerateQR(body);
  if (path === "presence/token/validate") return handleValidateToken(body);

  // Route Mahasiswa
  if (path === "presence/checkin")   return handleCheckin(body);
  if (path === "sensor/accel/batch") return handleAccelBatch(body);
  if (path === "sensor/gps")         return handleGPS(body);

  return sendError("Route POST tidak ditemukan");
}

// ─── DOSEN: Generate QR ───────────────────────────────────────
function handleGenerateQR(body) {
  if (!body.course_id || !body.session_id)
    return sendError("missing_field: course_id dan session_id wajib ada");

  const token = Utilities.getUuid();
  const now = new Date();
  const expires = new Date(now.getTime() + 120000);

  const qrPayload = {
    token, course_id: body.course_id,
    session_id: body.session_id,
    expires_at: expires.toISOString()
  };

  getSheet("tokens").appendRow([
    token, body.course_id, body.session_id,
    now.toISOString(), expires.toISOString()
  ]);

  return sendSuccess({ qr_data: qrPayload });
}

// ─── DOSEN: Validasi Token ────────────────────────────────────
function handleValidateToken(body) {
  const { token, course_id, session_id } = body;
  if (!token || !course_id || !session_id)
    return sendError("missing_field: token, course_id, session_id wajib ada");

  const data = getSheet("tokens").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (
      data[i][0] === token &&
      String(data[i][1]).toLowerCase() === String(course_id).toLowerCase() &&
      String(data[i][2]).toLowerCase() === String(session_id).toLowerCase()
    ) {
      if (new Date() > new Date(data[i][4])) return sendError("token_expired");
      return sendSuccess({ valid: true, course_id, session_id, token });
    }
  }
  return sendError("token_invalid");
}

// ─── MAHASISWA: Checkin ───────────────────────────────────────
function handleCheckin(body) {
  if (!body.user_id || !body.device_id || !body.course_id || !body.session_id || !body.token)
    return sendError("missing_field: user_id, device_id, course_id, session_id, token wajib ada");

  const sheet = getSheet("presence");
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === body.user_id && data[i][3] === body.course_id && data[i][4] === body.session_id)
      return sendError("user_already_checked_in");
  }

  const presenceId = "PR-" + Utilities.getUuid().substring(0, 6).toUpperCase();
  const now = new Date().toISOString();
  sheet.appendRow([presenceId, body.user_id, body.device_id, body.course_id, body.session_id, body.token, now, now]);
  return sendSuccess({ presence_id: presenceId, status: "checked_in" });
}

// ─── MAHASISWA: Status ────────────────────────────────────────
function handleGetStatus(e) {
  const { user_id, course_id, session_id } = e.parameter;
  const data = getSheet("presence").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1] === user_id && data[i][3] === course_id && data[i][4] === session_id)
      return sendSuccess({ user_id, course_id, session_id, status: "checked_in", last_ts: data[i][7] });
  }
  return sendError("not_checked_in");
}

// ─── SENSOR: Accel Batch ──────────────────────────────────────
function handleAccelBatch(body) {
  if (!body.device_id || !Array.isArray(body.data))
    return sendError("invalid_batch_data");

  const sheet = getSheet("accel");
  const now = new Date().toISOString();
  const rows = body.data.map(d => [body.device_id, d.x, d.y, d.z, d.ts, now, now]);
  if (rows.length > 0)
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return sendSuccess({ saved_records: rows.length });
}

// ─── SENSOR: GPS ──────────────────────────────────────────────
function handleGPS(body) {
  if (!body.device_id || !body.lat || !body.lng)
    return sendError("missing_gps_data");

  const now = new Date().toISOString();
  getSheet("gps").appendRow([body.device_id, body.lat, body.lng, body.accuracy || 0, body.altitude || 0, body.ts || now, now]);
  return sendSuccess({ status: "recorded" });
}

// ─── SENSOR: Accel Latest ─────────────────────────────────────
function handleGetAccelLatest(e) {
  if (!e.parameter.nim) return sendError("missing_nim");
  const data = getSheet("accel").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim))
      return sendSuccess({ x: data[i][1], y: data[i][2], z: data[i][3], ts: data[i][6] });
  }
  return sendError("data_not_found");
}

// ─── SENSOR: GPS Marker ───────────────────────────────────────
function handleGetGPSMarker(e) {
  if (!e.parameter.nim) return sendError("missing_nim");
  const data = getSheet("gps").getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]).startsWith(e.parameter.nim))
      return sendSuccess({ lat: data[i][1], lng: data[i][2], ts: data[i][6] });
  }
  return sendError("data_not_found");
}

// ─── Helper ───────────────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}
function sendSuccess(data) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data })).setMimeType(ContentService.MimeType.JSON);
}
function sendError(error) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error })).setMimeType(ContentService.MimeType.JSON);
}