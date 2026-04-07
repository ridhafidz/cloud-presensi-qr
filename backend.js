// ============================================================
// GAS BRIDGE - gas-bridge.js
// Deploy ini sebagai Google Apps Script TERPISAH (Web App #3)
// TIDAK perlu spreadsheet sendiri.
// Semua request dari frontend (dosen.html & mahasiswa.html) 
// masuk ke sini, lalu diteruskan ke GAS yang tepat.
// ============================================================

// URL kedua GAS yang sudah di-deploy sebagai Web App
const URL_GAS_DOSEN     = "https://script.google.com/macros/s/AKfycbwkzAgfyr1ArXGgag5DdI-3MrCmjEHZT-zmO8IBoO_tMbxpIbjqhGGARHNHjkTkZl-fCA/exec";
const URL_GAS_MAHASISWA = "https://script.google.com/macros/s/AKfycbwkzAgfyr1ArXGgag5DdI-3MrCmjEHZT-zmO8IBoO_tMbxpIbjqhGGARHNHjkTkZl-fCA/exec";

// ============================================================
// ROUTING GET (dosen pantau sensor mahasiswa)
// ============================================================
function doGet(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";

  // Route GET ke GAS Mahasiswa
  if (path === "presence/status")     return forwardGet(URL_GAS_MAHASISWA, path, e.parameter);
  if (path === "sensor/accel/latest") return forwardGet(URL_GAS_MAHASISWA, path, e.parameter);
  if (path === "sensor/gps/marker")   return forwardGet(URL_GAS_MAHASISWA, path, e.parameter);

  return sendError("Route GET tidak dikenal di Bridge");
}

// ============================================================
// ROUTING POST
// ============================================================
function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  let body = {};

  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return sendError("Invalid JSON Body");
  }

  // Dosen generate QR → langsung ke GAS Dosen
  if (path === "presence/qr/generate") {
    return forwardPost(URL_GAS_DOSEN, path, body);
  }

  // Mahasiswa checkin → Bridge validasi token ke GAS Dosen dulu,
  // baru simpan presence ke GAS Mahasiswa
  if (path === "presence/checkin") {
    return handleCheckinBridge(body);
  }

  // Sensor mahasiswa → langsung ke GAS Mahasiswa
  if (path === "sensor/accel/batch") return forwardPost(URL_GAS_MAHASISWA, path, body);
  if (path === "sensor/gps")         return forwardPost(URL_GAS_MAHASISWA, path, body);

  return sendError("Route POST tidak dikenal di Bridge");
}

// ============================================================
// LOGIC CHECKIN: Validasi ke GAS Dosen, lalu simpan ke GAS Mahasiswa
// ============================================================
function handleCheckinBridge(body) {
  // Pastikan qr_data ada
  if (!body.user_id || !body.device_id || !body.qr_data) {
    return sendError("missing_field: user_id, device_id, qr_data wajib ada");
  }

  const { token, course_id, session_id, expires_at } = body.qr_data;

  // Cek expired di sisi Bridge dulu (tanpa request ke server)
  if (new Date() > new Date(expires_at)) {
    return sendError("qr_expired");
  }

  // Step 1: Validasi token ke GAS Dosen
  const validasiResult = forwardPostRaw(URL_GAS_DOSEN, "presence/token/validate", {
    token,
    course_id,
    session_id
  });

  if (!validasiResult.ok) {
    return sendRawResult(validasiResult); // token_invalid atau token_expired
  }

  // Step 2: Simpan checkin ke GAS Mahasiswa
  const checkinPayload = {
    user_id:    body.user_id,
    device_id:  body.device_id,
    course_id,
    session_id,
    token
  };

  return forwardPost(URL_GAS_MAHASISWA, "presence/checkin", checkinPayload);
}

// ============================================================
// HELPER: Forward GET request ke GAS lain
// ============================================================
function forwardGet(targetUrl, path, params) {
  try {
    // Bangun query string dari semua parameter (selain 'path' itu sendiri)
    const queryParts = [`path=${encodeURIComponent(path)}`];
    for (const key in params) {
      if (key !== "path") {
        queryParts.push(`${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`);
      }
    }

    const url = targetUrl + "?" + queryParts.join("&");
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());
    return sendRawResult(result);

  } catch (err) {
    return sendError("Bridge forwardGet error: " + err.toString());
  }
}

// ============================================================
// HELPER: Forward POST request ke GAS lain (return ContentService)
// ============================================================
function forwardPost(targetUrl, path, body) {
  try {
    const result = forwardPostRaw(targetUrl, path, body);
    return sendRawResult(result);
  } catch (err) {
    return sendError("Bridge forwardPost error: " + err.toString());
  }
}

// ============================================================
// HELPER: Forward POST dan return object (bukan ContentService)
// Dipakai untuk chaining (misalnya checkin yang butuh 2 step)
// ============================================================
function forwardPostRaw(targetUrl, path, body) {
  const response = UrlFetchApp.fetch(targetUrl + "?path=" + encodeURIComponent(path), {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return JSON.parse(response.getContentText());
}

// ============================================================
// HELPER: Kembalikan hasil dari GAS lain sebagai ContentService
// ============================================================
function sendRawResult(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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