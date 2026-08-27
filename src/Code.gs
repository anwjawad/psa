/**
 * PSA Eligibility Checklist — Patient Web App
 * Augusta Victoria Hospital — Prostate Awareness Campaign
 *
 * Server-side controller. The browser NEVER talks to Google Sheets directly —
 * every write goes through submitPsaForm() below, which validates, locks,
 * de-duplicates and appends a single row per patient.
 *
 * Configuration (Spreadsheet ID / Sheet name) lives in Script Properties,
 * never in the HTML/JS sent to the browser. Run setupProject() once from the
 * Apps Script editor (or `clasp run setupProject`) to provision everything.
 */

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

var SHEET_NAME = 'PSA_RESPONSES';
var SPREADSHEET_NAME = 'PSA_RESPONSES_DB — Prostate Awareness Campaign';
var LOCK_WAIT_MS = 30000; // 30s max wait for the script lock

// Default admin passcode set the FIRST time setupProject() runs (only if
// ADMIN_PASSCODE isn't already set — never overwrites an existing one).
// Change it any time from Script Properties without touching code.
var DEFAULT_ADMIN_PASSCODE = 'A7GXL-DJLGD';

// Canonical column order for the Google Sheet. [internalKey, HeaderText]
// internalKey must match the keys sent from JavaScript.html (see buildRow_).
var FIELD_MAP = [
  ['submission_id', 'Submission_ID'],
  ['submitted_at', 'Submitted_At'],
  ['client_submission_id', 'Client_Submission_ID'],
  ['region', 'Region'],
  ['full_name', 'Full_Name'],
  ['national_id', 'National_ID'],
  ['address', 'Address'],
  ['age', 'Age'],
  ['phone', 'Phone'],
  ['smoker', 'Smoker'],
  ['previous_psa', 'Previous_PSA'],
  ['previous_psa_result', 'Previous_PSA_Result'],
  ['q1_age50_good_health', 'Q1_Age50_GoodHealth'],
  ['q3_age40_family_history_early', 'Q3_Age40_FamilyHistory_EarlyOnset'],
  ['q4_family_history_breast_ovarian_pancreatic_brca', 'Q4_FamilyHistory_Breast_Ovarian_Pancreatic_BRCA'],
  ['q5_urinary_symptoms', 'Q5_Urinary_Symptoms'],
  ['back_pain_pelvic_numbness', 'Back_Pain_Pelvic_Numbness'],
  ['unexplained_weight_loss', 'Unexplained_Weight_Loss'],
  ['q6_no_psa_last_2_years', 'Q6_No_PSA_Last_2Years'],
  ['q7_chronic_disease', 'Q7_Chronic_Disease'],
  ['chronic_bph', 'Chronic_BPH'],
  ['chronic_uti', 'Chronic_UTI'],
  ['chronic_hypertension', 'Chronic_Hypertension'],
  ['chronic_diabetes', 'Chronic_Diabetes'],
  ['chronic_other', 'Chronic_Other'],
  ['chronic_other_text', 'Chronic_Other_Text'],
  ['q8_no_previous_prostate_cancer', 'Q8_No_Previous_Prostate_Cancer'],
  ['previous_cancer_following_specialist', 'Previous_Cancer_Following_Specialist'],
  ['bph_medications', 'BPH_Medications'],
  ['q9_wants_psa', 'Q9_Wants_PSA'],
  ['age_above_75', 'Age_Above_75'],
  // Staff-only fields — never sent to or filled by the patient. Written only
  // by the Admin view (setEligibility below).
  ['staff_recommendation', 'Staff_Recommendation'],
  ['staff_notes', 'Staff_Notes'],
  ['staff_reviewed_by', 'Staff_Reviewed_By'],
  ['staff_reviewed_at', 'Staff_Reviewed_At']
];

var VALID_REGIONS = ['North', 'Central', 'South'];
var VALID_YES_NO = ['Yes', 'No'];

// ---------------------------------------------------------------------------
// WEB APP ENTRY POINT
// ---------------------------------------------------------------------------

function doGet(e) {
  var isAdmin = e && e.parameter && e.parameter.admin === '1';
  if (isAdmin) {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle('لوحة إدارة استمارات PSA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('قائمة تحقق أهلية فحص البروستاتا PSA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/** Allows Index.html/Admin.html to inline other HTML files at render time. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// PUBLIC API — called from the browser via google.script.run
// ---------------------------------------------------------------------------

/**
 * Validates and stores one patient submission. Idempotent on
 * clientSubmissionId: retrying the same submission (double click, network
 * retry) will never create a duplicate row.
 *
 * @param {Object} payload - see FIELD_MAP for the patient-fillable keys,
 *   plus a required `client_submission_id` (UUID generated once on the
 *   device when the patient reaches the final review step).
 * @return {Object} { success: true, submissionId, duplicate } or
 *                   { success: false, error: <safe, patient-facing message> }
 */
function submitPsaForm(payload) {
  var lock = LockService.getScriptLock();
  try {
    var validation = validatePayload_(payload);
    if (!validation.ok) {
      return { success: false, error: validation.error };
    }
    var clean = validation.data;

    var gotLock = lock.tryLock(LOCK_WAIT_MS);
    if (!gotLock) {
      return { success: false, error: 'الخادم مشغول حاليًا، يرجى المحاولة مرة أخرى خلال لحظات.' };
    }

    var sheet = getSheet_();

    // Idempotency check — must happen INSIDE the lock so two concurrent
    // requests with the same clientSubmissionId cannot both pass the check
    // before either has written its row.
    var existing = findRowByClientSubmissionId_(sheet, clean.client_submission_id);
    if (existing) {
      return { success: true, submissionId: existing.submissionId, duplicate: true };
    }

    var submissionId = Utilities.getUuid();
    var submittedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    clean.submission_id = submissionId;
    clean.submitted_at = submittedAt;
    clean.age_above_75 = (clean.age >= 76) ? 'Yes' : 'No';

    var row = buildRow_(clean);
    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return { success: true, submissionId: submissionId, submittedAt: submittedAt, duplicate: false };
  } catch (err) {
    // Detailed error stays server-side only — never leaks to the patient.
    console.error('submitPsaForm failed: ' + (err && err.stack ? err.stack : err));
    return { success: false, error: 'تعذر إرسال الاستمارة. يرجى المحاولة مرة أخرى.' };
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* lock was never acquired */ }
  }
}

// ---------------------------------------------------------------------------
// VALIDATION (server-side — never trust the client)
// ---------------------------------------------------------------------------

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    return { ok: false, error: 'بيانات غير صالحة.' };
  }

  var d = {};

  // clientSubmissionId — required, used for idempotency
  d.client_submission_id = sanitizeText_(payload.client_submission_id, 100);
  if (!d.client_submission_id) {
    return { ok: false, error: 'حدث خطأ تقني (missing submission id). يرجى إعادة تحميل الصفحة.' };
  }

  // Region
  if (VALID_REGIONS.indexOf(payload.region) === -1) {
    return { ok: false, error: 'يرجى اختيار المنطقة.' };
  }
  d.region = payload.region;

  // Full name / national ID / address — required free text
  d.full_name = sanitizeText_(payload.full_name, 200);
  if (!d.full_name) return { ok: false, error: 'يرجى إدخال الاسم الكامل.' };

  d.national_id = sanitizeNationalId_(payload.national_id);
  if (!d.national_id) return { ok: false, error: 'يرجى إدخال رقم هوية صحيح.' };

  d.address = sanitizeText_(payload.address, 200);
  if (!d.address) return { ok: false, error: 'يرجى إدخال العنوان.' };

  // Age — sane numeric bounds
  var age = Number(payload.age);
  if (!isFinite(age) || age < 1 || age > 120 || Math.floor(age) !== age) {
    return { ok: false, error: 'يرجى إدخال عمر صحيح.' };
  }
  d.age = age;

  // Phone — sanitize to digits/plus only, never used arithmetically
  d.phone = sanitizePhone_(payload.phone);
  if (!d.phone) return { ok: false, error: 'يرجى إدخال رقم هاتف صحيح.' };

  // Yes/No fields
  var yesNoFields = [
    'smoker', 'previous_psa', 'q1_age50_good_health', 'q3_age40_family_history_early',
    'q4_family_history_breast_ovarian_pancreatic_brca', 'q5_urinary_symptoms',
    'back_pain_pelvic_numbness', 'unexplained_weight_loss', 'q6_no_psa_last_2_years',
    'q7_chronic_disease', 'q8_no_previous_prostate_cancer', 'bph_medications', 'q9_wants_psa'
  ];
  for (var i = 0; i < yesNoFields.length; i++) {
    var key = yesNoFields[i];
    var val = payload[key];
    if (VALID_YES_NO.indexOf(val) === -1) {
      return { ok: false, error: 'يرجى الإجابة على جميع الأسئلة.' };
    }
    d[key] = val;
  }

  // Conditional: previous_psa_result only meaningful if previous_psa === 'Yes'
  d.previous_psa_result = d.previous_psa === 'Yes' ? sanitizeText_(payload.previous_psa_result, 50) : '';

  // Conditional: previous_cancer_following_specialist only if q8 === 'No'
  // (Q8 = "لا يوجد لدي تشخيص سابق" — answering No means a prior diagnosis exists)
  if (d.q8_no_previous_prostate_cancer === 'No') {
    if (VALID_YES_NO.indexOf(payload.previous_cancer_following_specialist) === -1) {
      return { ok: false, error: 'يرجى الإجابة على جميع الأسئلة.' };
    }
    d.previous_cancer_following_specialist = payload.previous_cancer_following_specialist;
  } else {
    d.previous_cancer_following_specialist = '';
  }

  // Conditional: chronic disease multi-select only if q7 === 'Yes'
  if (d.q7_chronic_disease === 'Yes') {
    var chronic = payload.chronic || {};
    d.chronic_bph = chronic.bph ? 'Yes' : 'No';
    d.chronic_uti = chronic.uti ? 'Yes' : 'No';
    d.chronic_hypertension = chronic.hypertension ? 'Yes' : 'No';
    d.chronic_diabetes = chronic.diabetes ? 'Yes' : 'No';
    d.chronic_other = chronic.other ? 'Yes' : 'No';
    d.chronic_other_text = chronic.other ? sanitizeText_(chronic.otherText, 200) : '';
  } else {
    d.chronic_bph = d.chronic_uti = d.chronic_hypertension = d.chronic_diabetes = d.chronic_other = 'No';
    d.chronic_other_text = '';
  }

  // Staff-only fields are never patient-supplied — always start empty.
  d.staff_recommendation = '';
  d.staff_notes = '';
  d.staff_reviewed_by = '';
  d.staff_reviewed_at = '';

  return { ok: true, data: d };
}

/**
 * Strips control characters (NUL, other C0 codes, DEL) and caps length.
 * Deliberately implemented with charCodeAt comparisons rather than a
 * \x-escaped regex character class, to avoid any ambiguity in how such
 * escapes get stored/interpreted — this way the intent is unambiguous.
 */
function sanitizeText_(value, maxLen) {
  if (value === undefined || value === null) return '';
  var s = String(value).trim().replace(/[\r\n\t]+/g, ' ');
  var out = '';
  for (var i = 0; i < s.length; i++) {
    var code = s.charCodeAt(i);
    if (code >= 32 && code !== 127) out += s.charAt(i);
  }
  if (out.length > maxLen) out = out.substring(0, maxLen);
  return out;
}

function sanitizeNationalId_(value) {
  if (value === undefined || value === null) return '';
  // Digits only — never used arithmetically, kept as a string throughout.
  var s = String(value).trim().replace(/\D/g, '');
  if (s.length < 4 || s.length > 15) return '';
  return s;
}

function sanitizePhone_(value) {
  if (value === undefined || value === null) return '';
  // Keep digits and a single leading + only — never treated as a number.
  var s = String(value).trim().replace(/[^\d+]/g, '');
  if (s.replace(/\D/g, '').length < 7) return '';
  return s;
}

// ---------------------------------------------------------------------------
// SHEET ACCESS
// ---------------------------------------------------------------------------

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) {
    throw new Error('SPREADSHEET_ID is not configured. Run setupProject() once from the Apps Script editor.');
  }
  var ss = SpreadsheetApp.openById(ssId);
  var sheetName = props.getProperty('SHEET_NAME') || SHEET_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found in spreadsheet ' + ssId);
  }
  return sheet;
}

// appendRow()/setValues() apply Sheets' own "as-typed" auto-detection to
// digit-only strings and silently convert them to numbers — which DROPS
// leading zeros in national IDs and phone numbers. A leading straight quote
// is Apps Script's documented way to force a value to be stored as literal
// text regardless of what it looks like; the quote itself is never stored
// or displayed. Pre-setting the column to Plain Text (setupProject) is not
// sufficient on its own — appendRow can still re-detect a numeric format
// for the newly written cell.
var FORCE_TEXT_KEYS = ['submission_id', 'client_submission_id', 'national_id', 'phone'];

function buildRow_(data) {
  return FIELD_MAP.map(function (pair) {
    var key = pair[0];
    var val = (data[key] !== undefined && data[key] !== null) ? data[key] : '';
    if (FORCE_TEXT_KEYS.indexOf(key) !== -1 && val !== '') {
      val = "'" + val;
    }
    return val;
  });
}

/**
 * Scans the Client_Submission_ID column for a match. Runs only while the
 * script lock is held, so this is race-free even under heavy concurrent load.
 */
function findRowByClientSubmissionId_(sheet, clientSubmissionId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var colIndex = FIELD_MAP.map(function (p) { return p[0]; }).indexOf('client_submission_id') + 1;
  var submissionColIndex = FIELD_MAP.map(function (p) { return p[0]; }).indexOf('submission_id') + 1;

  var values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === clientSubmissionId) {
      var submissionId = sheet.getRange(i + 2, submissionColIndex, 1, 1).getValue();
      return { rowIndex: i + 2, submissionId: submissionId };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// ADMIN VIEW — accessed at <web-app-url>?admin=1, gated by ADMIN_PASSCODE
// (Script Properties). This is a lightweight passcode gate, not a real
// login system — the admin link + passcode must be shared only with staff
// who should see patient data (names, national IDs, medical answers).
// ---------------------------------------------------------------------------

var VALID_STAFF_RECOMMENDATIONS = ['Eligible', 'Not Eligible', ''];

function assertAdminPasscode_(passcode) {
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
  if (!expected || !passcode || String(passcode) !== expected) {
    throw new Error('رمز الدخول غير صحيح.');
  }
}

/**
 * Returns every stored submission (all FIELD_MAP columns) for the admin
 * dashboard. The dataset size expected for an awareness-campaign sheet is
 * small, so a single full fetch (filtered/searched client-side) keeps this
 * simple — no separate list/detail endpoints to keep in sync.
 */
function getAdminData(passcode) {
  assertAdminPasscode_(passcode);
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var keys = FIELD_MAP.map(function (p) { return p[0]; });
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, keys.length).getValues();
  return values.map(function (row) {
    var obj = {};
    keys.forEach(function (k, idx) { obj[k] = row[idx]; });
    return obj;
  });
}

/**
 * Records (or clears) the staff eligibility call for one submission.
 * reviewerName is free text the admin UI asks for once per session purely
 * for an audit trail — it is never validated against an identity system.
 */
function setEligibility(passcode, submissionId, decision, reviewerName) {
  assertAdminPasscode_(passcode);
  if (VALID_STAFF_RECOMMENDATIONS.indexOf(decision) === -1) {
    throw new Error('قيمة غير صالحة.');
  }

  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(LOCK_WAIT_MS);
  if (!gotLock) {
    throw new Error('الخادم مشغول، حاول مرة أخرى.');
  }

  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('لا يوجد بيانات.');

    var keys = FIELD_MAP.map(function (p) { return p[0]; });
    var idCol = keys.indexOf('submission_id') + 1;
    var recCol = keys.indexOf('staff_recommendation') + 1;
    var byCol = keys.indexOf('staff_reviewed_by') + 1;
    var atCol = keys.indexOf('staff_reviewed_at') + 1;

    var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === submissionId) {
        var r = i + 2;
        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(r, recCol).setValue(decision);
        sheet.getRange(r, byCol).setValue(sanitizeText_(reviewerName, 100));
        sheet.getRange(r, atCol).setValue(decision ? now : '');
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    throw new Error('لم يتم العثور على هذا المريض.');
  } finally {
    try { lock.releaseLock(); } catch (e2) { /* not acquired */ }
  }
}

// ---------------------------------------------------------------------------
// ONE-TIME SETUP — run manually from the Apps Script editor (or `clasp run`)
// ---------------------------------------------------------------------------

/**
 * Creates the response spreadsheet (if not already configured) and writes
 * the header row. Safe to re-run: it will not duplicate the header, won't
 * overwrite an already-configured SPREADSHEET_ID, and won't overwrite an
 * already-set ADMIN_PASSCODE.
 */
function setupProject() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      throw new Error('SPREADSHEET_ID (' + ssId + ') is set but could not be opened: ' + e);
    }
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  props.setProperty('SHEET_NAME', SHEET_NAME);

  if (!props.getProperty('ADMIN_PASSCODE')) {
    props.setProperty('ADMIN_PASSCODE', DEFAULT_ADMIN_PASSCODE);
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }

  var headers = FIELD_MAP.map(function (p) { return p[1]; });
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var headersMatch = headers.every(function (h, i) { return firstRow[i] === h; });

  if (!headersMatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  // Force these columns to Plain Text formatting. Without this, Sheets
  // auto-detects digit-only strings (national ID, phone, UUIDs) as numbers
  // and SILENTLY DROPS leading zeros — a real corruption risk, not cosmetic.
  var textOnlyKeys = ['submission_id', 'client_submission_id', 'national_id', 'phone'];
  var keys = FIELD_MAP.map(function (p) { return p[0]; });
  textOnlyKeys.forEach(function (key) {
    var col = keys.indexOf(key) + 1;
    if (col > 0) {
      sheet.getRange(1, col, sheet.getMaxRows(), 1).setNumberFormat('@');
    }
  });

  Logger.log('Spreadsheet ready: ' + ss.getUrl());
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('Admin passcode: ' + props.getProperty('ADMIN_PASSCODE'));
  return ss.getUrl();
}
