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

// Canonical column order for the Google Sheet. [internalKey, HeaderText]
// internalKey must match the keys sent from JavaScript.html (see buildRow_).
var FIELD_MAP = [
  ['submission_id', 'Submission_ID'],
  ['submitted_at', 'Submitted_At'],
  ['client_submission_id', 'Client_Submission_ID'],
  ['region', 'Region'],
  ['full_name', 'Full_Name'],
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
  // Staff-only fields — never sent to or filled by the patient. Reserved so a
  // future Staff View can write into the same row without a schema change.
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
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('قائمة تحقق أهلية فحص البروستاتا PSA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/** Allows Index.html to inline Styles.html / JavaScript.html at render time. */
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

  // Full name / address — required free text
  d.full_name = sanitizeText_(payload.full_name, 200);
  if (!d.full_name) return { ok: false, error: 'يرجى إدخال الاسم الكامل.' };

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

function sanitizeText_(value, maxLen) {
  if (value === undefined || value === null) return '';
  var s = String(value).trim();
  // Strip control characters and cap length; keep Arabic/Latin text intact.
  s = s.replace(/[\r\n\t]+/g, ' ').replace(/[ --]/g, '');
  if (s.length > maxLen) s = s.substring(0, maxLen);
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

function buildRow_(data) {
  return FIELD_MAP.map(function (pair) {
    var key = pair[0];
    return (data[key] !== undefined && data[key] !== null) ? data[key] : '';
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
// ONE-TIME SETUP — run manually from the Apps Script editor (or `clasp run`)
// ---------------------------------------------------------------------------

/**
 * Creates the response spreadsheet (if not already configured) and writes
 * the header row. Safe to re-run: it will not duplicate the header or
 * overwrite an already-configured SPREADSHEET_ID.
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

  Logger.log('Spreadsheet ready: ' + ss.getUrl());
  Logger.log('Spreadsheet ID: ' + ss.getId());
  return ss.getUrl();
}
