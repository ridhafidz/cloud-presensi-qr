const SPREADSHEET_DOSEN = "1A-PRMkgt7YbXCelkrrfSLMBHh1fdzACXiM4gy4trONM";

function doGet(e) {
  return sendError("Route GET not found di GAS Dosen");
}

function doPost(e) {
  const path = (e.parameter && e.parameter.path) ? e.parameter.path : "";
  let body = {};

  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return sendError("Invalid JSON Body");
  }

  if (path === "presence/qr/generate") return handleGenerateQR(body);
  if (path === "presence/token/validate") return handleValidateToken(body);

  return sendError("Route POST tidak ditemukan di GAS Dosen");
}

// --- Generate QR ---
function handleGenerateQR(body) {
  if (!body.course_id || !body.session_id) {
    return sendError("missing_field: course_id dan session_id wajib ada");
  }

  const token = Utilities.getUuid();
  const now = new Date();
  const expires = new Date(now.getTime() + 120000); // 2 menit

  const qrPayload = {
    token: token,
    course_id: body.course_id,
    session_id: body.session_id,
    expires_at: expires.toISOString()
  };

  const sheet = SpreadsheetApp.openById(SPREADSHEET_DOSEN).getSheetByName("tokens");
  sheet.appendRow([
    token,
    body.course_id,
    body.session_id,
    now.toISOString(),
    expires.toISOString()
  ]);

  return sendSuccess({ qr_data: qrPayload });
}

// --- Validasi Token (dipanggil oleh Bridge saat mahasiswa checkin) ---
function handleValidateToken(body) {
  const { token, course_id, session_id } = body;

  if (!token || !course_id || !session_id) {
    return sendError("missing_field: token, course_id, session_id wajib ada");
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_DOSEN).getSheetByName("tokens");
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (
      data[i][0] === token &&
      String(data[i][1]).toLowerCase() === String(course_id).toLowerCase() &&
      String(data[i][2]).toLowerCase() === String(session_id).toLowerCase()
    ) {
      if (new Date() > new Date(data[i][4])) {
        return sendError("token_expired");
      }
      return sendSuccess({ valid: true, course_id, session_id, token });
    }
  }

  return sendError("token_invalid");
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