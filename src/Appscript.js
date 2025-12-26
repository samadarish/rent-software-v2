/**
 * Apps Script entry point implementing normalized Sheets schema.
 *
 * The schema follows a database-like layout:
 * - Master data: TENANTS, UNITS, TENANCIES, FAMILY_MEMBERS, CLAUSES
 * - Monthly inputs: WING_MONTHLY_CONFIG, TENANT_MONTHLY_READINGS
 * - Outputs: BILL_LINES
 * - Events: PAYMENTS, ATTACHMENTS (for payment proofs)
 *
 * Actions exposed via doGet/doPost mirror the previous API surface while
 * internally writing to the normalized tables.
 */

/********* CONFIG *********/
const TENANTS_SHEET = 'Tenants';
const UNITS_SHEET = 'Units';
const TENANCIES_SHEET = 'Tenancies';
const FAMILY_SHEET = 'FamilyMembers';
const WINGS_SHEET = 'Wings';
const LANDLORDS_SHEET = 'Landlords';
const CLAUSES_SHEET = 'Clauses';
const WING_MONTHLY_SHEET = 'WingMonthlyConfig';
const TENANT_READINGS_SHEET = 'TenantMonthlyReadings';
const BILL_LINES_SHEET = 'BillLines';
const PAYMENTS_SHEET = 'Payments';
const ATTACHMENTS_SHEET = 'Attachments';
const INDEX_SHEET = 'Index';
const TENANCY_RENT_REVISIONS_SHEET = 'TenancyRentRevisions';

/********* HEADERS *********/
const TENANTS_HEADERS = [
  'tenant_id',
  'full_name',
  'mobile',
  'aadhaar',
  'occupation',
  'permanent_address',
  'created_at',
];

const UNITS_HEADERS = [
  'unit_id',
  'wing',
  'unit_number',
  'floor',
  'direction',
  'meter_number',
  'landlord_id',
  'notes',
  'is_occupied',
  'current_tenancy_id',
  'created_at',
];

const TENANCIES_HEADERS = [
  'tenancy_id',
  'tenant_id',
  'grn_number',
  'unit_id',
  'landlord_id',
  'agreement_date',
  'commencement_date',
  'end_date',
  'status',
  'vacate_reason',
  'security_deposit',
  'rent_payable_day',
  'tenant_notice_months',
  'landlord_notice_months',
  'pet_policy',
  'late_rent_per_day',
  'late_grace_days',
  'rent_revision_unit',
  'rent_revision_number',
  'created_at',
  'rent_increase_amount',
];

const LANDLORD_HEADERS = [
  'landlord_id',
  'name',
  'aadhaar',
  'address',
  'created_at',
];

const FAMILY_HEADERS = [
  'member_id',
  'tenant_id',
  'name',
  'relationship',
  'occupation',
  'aadhaar',
  'address',
  'created_at',
];

const CLAUSES_HEADERS = [
  'Section',
  'SortOrder',
  'Enabled',
  'ClauseHtml',
];

const WING_MONTHLY_HEADERS = [
  'wing_month_id',
  'month_key',
  'wing',
  'electricity_rate',
  'sweeping_per_flat',
  'motor_prev',
  'motor_new',
  'motor_units',
  'created_at',
];

const TENANT_READING_HEADERS = [
  'reading_id',
  'month_key',
  'tenancy_id',
  'prev_reading',
  'new_reading',
  'included',
  'override_rent',
  'notes',
  'created_at',
];

const TENANCY_RENT_REVISION_HEADERS = [
  'revision_id',
  'tenancy_id',
  'effective_month',
  'rent_amount',
  'note',
  'created_at',
];

const BILL_LINE_HEADERS = [
  'bill_line_id',
  'month_key',
  'tenancy_id',
  'rent_amount',
  'electricity_units',
  'electricity_amount',
  'motor_share_amount',
  'sweep_amount',
  'total_amount',
  'payable_date',
  'generated_at',
  'amount_paid',
  'is_paid',
];

const PAYMENT_HEADERS = [
  'payment_id',
  'payment_date',
  'bill_line_id',
  'tenant_id',
  'amount',
  'mode',
  'reference',
  'notes',
  'attachment_id',
  'created_at',
];

const ATTACHMENT_HEADERS = [
  'attachment_id',
  'file_name',
  'file_url',
  'file_drive_id',
  'uploaded_at',
];



const driveShareCache = {};
const LOOKUP_CACHE_TTL_SECONDS = 300;

/********* COMMON HELPERS *********/
function parseIsoDate_(iso) {
  if (!iso) return '';
  if (iso instanceof Date) return iso;
  try {
    return new Date(iso.toString());
  } catch (err) {
    return '';
  }
}

function normalizeMonthKey_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM');
  }
  const str = value.toString().trim();
  const compact = str.match(/^(\d{4})(\d{2})$/);
  if (compact) return `${compact[1]}-${compact[2]}`;
  const dashed = str.match(/^(\d{4})[-/.](\d{1,2})/);
  if (dashed) return `${dashed[1]}-${dashed[2].padStart(2, '0')}`;
  return str;
}

function formatDateIso_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : parseIsoDate_(value);
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd');
}

function formatDateTime_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : parseIsoDate_(value);
  if (!(d instanceof Date) || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
}

function formatMonthLabelForDisplay_(monthKey) {
  const normalized = normalizeMonthKey_(monthKey || '');
  const match = normalized.match(/^(\d{4})-(\d{2})$/);
  if (!match) return normalized;
  const d = new Date(Number(match[1]), Number(match[2]) - 1, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Kolkata', 'MMM yyyy');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getMasterSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureHeaderRow_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const existingSet = new Set(
    existing.map((val) => (val || '').toString().trim()).filter(Boolean)
  );

  let changed = false;
  const updated = existing.slice();
  headers.forEach((header) => {
    const label = (header || '').toString().trim();
    if (!label) return;
    if (!existingSet.has(label)) {
      updated.push(label);
      existingSet.add(label);
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, updated.length).setValues([updated]);
  }
}

function getSheetWithHeaders_(name, headers) {
  const ss = getMasterSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaderRow_(sheet, headers);
  return sheet;
}

function buildHeaderIndex_(headers) {
  return headers.reduce((map, key, idx) => {
    map[key] = idx;
    return map;
  }, {});
}

function updateUnitOccupancy_(unitId, tenancyId, occupied) {
  if (!unitId) return;
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const match = units.find((u) => u.unit_id === unitId);
  if (!match) return;
  match.is_occupied = !!occupied;
  match.current_tenancy_id = occupied ? tenancyId : '';
  upsertUnique_(UNITS_SHEET, UNITS_HEADERS, ['unit_id'], match);
}

function normalizeBoolean_(val) {
  if (val === true || val === false) return val;
  if (typeof val === 'string') {
    return val.toLowerCase() !== 'false' && val !== '';
  }
  return !!val;
}

function readTable_(sheetName, headers) {
  const sheet = getSheetWithHeaders_(sheetName, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .map((row) => {
      const record = {};
      headers.forEach((key, idx) => {
        record[key] = row[idx];
      });
      return record;
    })
    .filter((r) => Object.values(r).some((v) => v !== ''));
}

function readTableColumns_(sheetName, headers, columns) {
  const sheet = getSheetWithHeaders_(sheetName, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const headerIndex = buildHeaderIndex_(headers);
  const requested = (Array.isArray(columns) ? columns : []).filter(Boolean);
  const indices = requested.map((key) => headerIndex[key]).filter((idx) => idx !== undefined);
  if (!indices.length) return [];
  const maxIndex = Math.max.apply(null, indices);
  const values = sheet.getRange(2, 1, lastRow - 1, maxIndex + 1).getValues();
  return values
    .map((row) => {
      const record = {};
      requested.forEach((key) => {
        const idx = headerIndex[key];
        record[key] = idx !== undefined ? row[idx] : '';
      });
      return record;
    })
    .filter((r) => Object.values(r).some((v) => v !== ''));
}

function getScriptCache_() {
  return CacheService.getScriptCache();
}

function getCachedJson_(key) {
  if (!key) return null;
  try {
    const cache = getScriptCache_();
    const payload = cache.get(key);
    if (!payload) return null;
    return JSON.parse(payload);
  } catch (err) {
    return null;
  }
}

function setCachedJson_(key, value, ttlSeconds) {
  if (!key) return;
  try {
    const cache = getScriptCache_();
    cache.put(key, JSON.stringify(value), ttlSeconds || LOOKUP_CACHE_TTL_SECONDS);
  } catch (err) {
    // Ignore cache write failures.
  }
}

function getCachedLookup_(key, buildFn, ttlSeconds) {
  const cached = getCachedJson_(key);
  if (cached) return cached;
  if (typeof buildFn !== 'function') return {};
  const built = buildFn() || {};
  setCachedJson_(key, built, ttlSeconds);
  return built;
}

function buildLookupByKey_(rows, keyField) {
  const map = {};
  rows.forEach((row) => {
    const key = row && row[keyField];
    if (key !== undefined && key !== null && key !== '') {
      map[key] = row;
    }
  });
  return map;
}

function getMonthIndex_(monthKey) {
  const normalized = normalizeMonthKey_(monthKey || '');
  const match = normalized.match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  return Number(match[1]) * 12 + (Number(match[2]) - 1);
}

function isWithinRecentMonths_(monthKey, monthsBack) {
  const limit = Number(monthsBack) || 0;
  if (!limit || limit <= 0) return true;
  const idx = getMonthIndex_(monthKey);
  if (idx === null) return true;
  const now = new Date();
  const currentIdx = now.getFullYear() * 12 + now.getMonth();
  return idx >= currentIdx - (limit - 1);
}

function writeTableRows_(sheetName, headers, rows) {
  const sheet = getSheetWithHeaders_(sheetName, headers);
  const mapped = rows.map((record) => headers.map((key) => record[key] ?? ''));
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, mapped.length, headers.length).setValues(mapped);
}

function upsertUnique_(sheetName, headers, uniqueColumns, record) {
  const sheet = getSheetWithHeaders_(sheetName, headers);
  const lastRow = sheet.getLastRow();
  const headerIndex = buildHeaderIndex_(headers);
  const normalizeKeyValue = (col, value) => {
    if (col === 'month_key') return normalizeMonthKey_(value);
    if (col === 'wing') return (value || '').toString().trim().toLowerCase();
    return value;
  };
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const matches = uniqueColumns.every(
        (col) => normalizeKeyValue(col, row[headerIndex[col]]) == normalizeKeyValue(col, record[col])
      );
      if (matches) {
        const targetRow = i + 2;
        sheet.getRange(targetRow, 1, 1, headers.length).setValues([
          headers.map((key) => record[key] ?? ''),
        ]);
        return record;
      }
    }
  }
  sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([
    headers.map((key) => record[key] ?? ''),
  ]);
  return record;
}

function ensureDriveFileShareable_(url) {
  const idMatch = url && url.match(/\/d\/([^/]+)/);
  const id = idMatch && idMatch[1];
  if (!id) return { originalUrl: url, viewUrl: url };
  if (driveShareCache[id]) return { originalUrl: url, viewUrl: driveShareCache[id] };
  try {
    const file = DriveApp.getFileById(id);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const viewUrl = file.getUrl();
    driveShareCache[id] = viewUrl;
    return { originalUrl: url, viewUrl };
  } catch (err) {
    Logger.log('Drive share failed: ' + err);
    return { originalUrl: url, viewUrl: url };
  }
}

function extractDriveFileId_(url) {
  if (!url) return '';
  const match = url.match(/\/d\/([^/]+)/) || url.match(/[?&]id=([^&#]+)/);
  return match && match[1] ? match[1] : '';
}

function handleAttachmentPreview_(attachmentUrl) {
  const id = extractDriveFileId_(attachmentUrl);
  if (!id) return jsonResponse({ ok: false, previewUrl: '', attachmentUrl: attachmentUrl || '' });
  try {
    const file = DriveApp.getFileById(id);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const ensured = ensureDriveFileShareable_(file.getUrl());
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const sizeLimit = 2 * 1024 * 1024;
    if (bytes.length > sizeLimit) {
      return jsonResponse({
        ok: true,
        previewUrl: '',
        attachmentUrl: ensured.viewUrl,
      });
    }
    const dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(bytes)}`;
    return jsonResponse({
      ok: true,
      previewUrl: dataUrl,
      attachmentUrl: ensured.viewUrl,
    });
  } catch (err) {
    Logger.log('Attachment preview failed: ' + err);
    return jsonResponse({ ok: false, previewUrl: '', attachmentUrl: attachmentUrl || '' });
  }
}

function sanitizeFileSegment_(value, fallback) {
  const raw = (value || '').toString().trim();
  if (!raw) return fallback || '';
  const cleaned = raw.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  return cleaned || (fallback || '');
}

function getFileExtension_(mimeType, originalName) {
  const name = (originalName || '').toString();
  const dot = name.lastIndexOf('.');
  if (dot >= 0 && dot < name.length - 1) {
    const ext = name.substring(dot + 1).replace(/[^A-Za-z0-9]/g, '');
    if (ext) return ext.toLowerCase();
  }
  const normalized = (mimeType || '').toString().toLowerCase();
  if (normalized.includes('jpeg')) return 'jpg';
  if (normalized.includes('png')) return 'png';
  if (normalized.includes('webp')) return 'webp';
  if (normalized.includes('gif')) return 'gif';
  return 'png';
}

/********* CLAUSES *********/
function readUnifiedClauses_() {
  const sheet = getSheetWithHeaders_(CLAUSES_SHEET, CLAUSES_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { tenant: [], landlord: [], penalties: [], misc: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, CLAUSES_HEADERS.length).getValues();
  const sections = { tenant: [], landlord: [], penalties: [], misc: [] };
  values.forEach((row) => {
    const section = (row[0] || '').toString().toLowerCase();
    const entry = { sortOrder: Number(row[1]) || 0, enabled: normalizeBoolean_(row[2]), html: row[3] || '' };
    if (sections[section]) sections[section].push(entry);
  });
  Object.keys(sections).forEach((key) => sections[key].sort((a, b) => a.sortOrder - b.sortOrder));
  return sections;
}



function writeUnifiedClauses_(payload) {
  const sheet = getSheetWithHeaders_(CLAUSES_SHEET, CLAUSES_HEADERS);
  const rows = [];
  ['tenant', 'landlord', 'penalties', 'misc'].forEach((sectionKey) => {
    const items = Array.isArray(payload[sectionKey]) ? payload[sectionKey] : [];
    items.forEach((item, idx) => {
      rows.push([
        sectionKey,
        Number(item.sortOrder || idx + 1),
        normalizeBoolean_(item.enabled),
        item.html || item.text || '',
      ]);
    });
  });
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, CLAUSES_HEADERS.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, CLAUSES_HEADERS.length).setValues(rows);
}

/********* TENANTS + TENANCIES *********/
function mapTenantPayload_(payload) {
  const template = payload.templateData || {};
  const updates = payload.updates || {};
  const tenantId = payload.tenantId || template.tenant_id || Utilities.getUuid();
  const unitId = payload.unitId || template.unit_id || Utilities.getUuid();
  const tenancyId = payload.forceNewTenancyId || payload.tenancyId || template.tenancy_id || Utilities.getUuid();
  const landlordId = payload.landlordId || updates.landlordId || template.landlord_id || '';
  const now = new Date();

  const tenant = {
    tenant_id: tenantId,
    full_name: updates.tenantFullName || template.Tenant_Full_Name || '',
    mobile: updates.tenantMobile || template.tenant_mobile || '',
    aadhaar: updates.tenantAadhaar || template.tenant_Aadhar || '',
    occupation: payload.Tenant_occupation || updates.tenantOccupation || template.Tenant_occupation || '',
    permanent_address: updates.tenantPermanentAddress || template.Tenant_Permanent_Address || '',
    created_at: template.created_at || now,
  };

  const unit = {
    unit_id: unitId,
    wing: payload.wing || updates.wing || template.wing || '',
    unit_number: updates.unitNumber || payload.unit_number || template.unit_number || payload.flatNumber || '',
    floor: payload.floor_of_building || updates.floor || template.floor_of_building || '',
    direction: updates.direction || template.direction_build || '',
    meter_number: updates.meterNumber || template.meter_number || '',
    landlord_id: landlordId || template.landlord_id || '',
    notes: '',
    is_occupied: true,
    current_tenancy_id: tenancyId,
    created_at: template.created_at || now,
  };

  const activeFlag = typeof updates.activeTenant !== 'undefined' ? updates.activeTenant : payload.activeTenant;
  const tenancy = {
    tenancy_id: tenancyId,
    tenant_id: tenantId,
    grn_number: updates.grnNumber || template['GRN number'] || payload.grn || '',
    unit_id: unitId,
    landlord_id: landlordId,
    agreement_date: formatDateIso_(updates.agreementDateRaw || template.agreement_date_raw || template.agreement_date || ''),
    commencement_date: formatDateIso_(updates.tenancyCommencementRaw || template.tenancy_comm_raw || template.tenancy_comm || ''),
    end_date: formatDateIso_(updates.tenancyEndRaw || template.tenancy_end_raw || ''),
    status: activeFlag === false || (typeof activeFlag === 'string' && activeFlag.toLowerCase() === 'no') ? 'ENDED' : 'ACTIVE',
    vacate_reason: updates.vacateReason || template.vacateReason || '',
    security_deposit: updates.securityDeposit || template.secu_depo || '',
    rent_payable_day: updates.payableDate || template.payable_date_raw || '',
    tenant_notice_months: updates.tenantNoticeMonths || template.notice_num_t || '',
    landlord_notice_months: updates.landlordNoticeMonths || template.notice_num_l || '',
    pet_policy: updates.pet_policy || template.pet_text_area || '',
    late_rent_per_day: updates.lateRentPerDay || template.late_rent || '',
    late_grace_days: updates.lateGracePeriodDays || template.late_days || '',
    rent_revision_unit: updates.rentRevisionUnit || template['rent_rev year_mon'] || '',
    rent_revision_number: updates.rentRevisionNumber || template.rent_rev_number || '',
    rent_increase_amount: updates.rentIncreaseAmount || template.rent_inc || '',
    created_at: template.created_at || now,
  };

  const family = Array.isArray(payload.familyMembers)
    ? payload.familyMembers.map((fm) => ({
      member_id: Utilities.getUuid(),
      tenant_id: tenantId,
      name: fm.name || '',
      relationship: fm.relationship || '',
      occupation: fm.occupation || '',
      aadhaar: fm.aadhaar || '',
      address: fm.address || '',
      created_at: now,
    }))
    : [];

  const rentAmountValue = Number(updates.rentAmount || template.rent_amount);

  return { tenant, unit, tenancy, family, rentAmountValue: isNaN(rentAmountValue) ? null : rentAmountValue };
}

function handleSaveUnit_(payload) {
  const now = new Date();
  const unit = {
    unit_id: payload.unitId || Utilities.getUuid(),
    wing: payload.wing || '',
    unit_number: payload.unitNumber || payload.flatNumber || '',
    floor: payload.floor || '',
    direction: payload.direction || '',
    meter_number: payload.meterNumber || '',
    landlord_id: payload.landlordId || '',
    notes: payload.notes || '',
    is_occupied: normalizeBoolean_(payload.isOccupied),
    current_tenancy_id: payload.currentTenancyId || '',
    created_at: payload.created_at || now,
  };

  upsertUnique_(UNITS_SHEET, UNITS_HEADERS, ['unit_id'], unit);
  return jsonResponse({ ok: true, unit });
}

function handleSaveLandlord_(payload = {}) {
  const landlord = {
    landlord_id: payload.landlordId || Utilities.getUuid(),
    name: payload.name || '',
    aadhaar: payload.aadhaar || '',
    address: payload.address || '',
    created_at: new Date(),
  };

  upsertUnique_(LANDLORDS_SHEET, LANDLORD_HEADERS, ['landlord_id'], landlord);
  return jsonResponse({ ok: true, landlord });
}

function handleDeleteLandlord_(payload = {}) {
  const landlordId = payload.landlordId || '';
  if (!landlordId) return jsonResponse({ ok: false, error: 'Landlord ID missing' });

  const remaining = readTable_(LANDLORDS_SHEET, LANDLORD_HEADERS).filter((l) => l.landlord_id !== landlordId);
  const sheet = getSheetWithHeaders_(LANDLORDS_SHEET, LANDLORD_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, LANDLORD_HEADERS.length).clearContent();
  if (remaining.length) {
    const rows = remaining.map((r) => LANDLORD_HEADERS.map((key) => r[key] ?? ''));
    sheet.getRange(2, 1, rows.length, LANDLORD_HEADERS.length).setValues(rows);
  }
  return jsonResponse({ ok: true, landlordId });
}

function handleDeleteUnit_(payload) {
  const unitId = payload.unitId || '';
  if (!unitId) return jsonResponse({ ok: false, error: 'Unit ID missing' });
  const remaining = readTable_(UNITS_SHEET, UNITS_HEADERS).filter((u) => u.unit_id !== unitId);
  const sheet = getSheetWithHeaders_(UNITS_SHEET, UNITS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, UNITS_HEADERS.length).clearContent();
  if (remaining.length) {
    const rows = remaining.map((r) => UNITS_HEADERS.map((key) => r[key] ?? ''));
    sheet.getRange(2, 1, rows.length, UNITS_HEADERS.length).setValues(rows);
  }
  return jsonResponse({ ok: true, unitId });
}

function handleRemoveWing_(payload) {
  const wing = (payload.wing || '').toString().trim();
  if (!wing) return jsonResponse({ ok: false, error: 'Wing missing' });
  const sheet = getSheetWithHeaders_(WINGS_SHEET, ['Wing']);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonResponse({ ok: true, wing });
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const retained = values.filter((r) => (r[0] || '').toString().trim().toLowerCase() !== wing.toLowerCase());
  sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  if (retained.length) sheet.getRange(2, 1, retained.length, 1).setValues(retained);
  return jsonResponse({ ok: true, wing });
}

function handleSaveTenant_(payload) {
  const mapped = mapTenantPayload_(payload || {});
  upsertUnique_(TENANTS_SHEET, TENANTS_HEADERS, ['tenant_id'], mapped.tenant);
  upsertUnique_(UNITS_SHEET, UNITS_HEADERS, ['unit_id'], mapped.unit);
  upsertUnique_(TENANCIES_SHEET, TENANCIES_HEADERS, ['tenancy_id'], mapped.tenancy);

  updateUnitOccupancy_(mapped.tenancy.unit_id, mapped.tenancy.tenancy_id, (mapped.tenancy.status || '').toUpperCase() === 'ACTIVE');

  if (mapped.family && mapped.family.length) {
    writeTableRows_(FAMILY_SHEET, FAMILY_HEADERS, mapped.family);
  }

  // Only create initial rent revision for new tenancies
  if (mapped.rentAmountValue !== null) {
    const effectiveMonth =
      normalizeMonthKey_(mapped.tenancy.commencement_date || mapped.tenancy.agreement_date) || normalizeMonthKey_(new Date());
    if (effectiveMonth) {
      // Check if a revision already exists for this tenancy at this effective month
      const existingRevisions = listTenancyRentRevisions_(mapped.tenancy.tenancy_id);
      const existingRevisionForMonth = existingRevisions.find(
        (r) => normalizeMonthKey_(r.effective_month) === effectiveMonth
      );

      // Only create revision if one doesn't exist yet for this month
      if (!existingRevisionForMonth) {
        upsertTenancyRentRevision_({
          tenancyId: mapped.tenancy.tenancy_id,
          effectiveMonth,
          rentAmount: mapped.rentAmountValue,
          note: 'Initial rent',
        });
      }
    }
  }

  return jsonResponse({ ok: true, message: 'Tenant saved', tenantId: mapped.tenant.tenant_id, tenancyId: mapped.tenancy.tenancy_id });
}

function handleUpdateTenant_(payload) {
  const createNewTenancy = normalizeBoolean_(payload && payload.createNewTenancy);
  const previousTenancyId = payload.previousTenancyId || payload.tenancyId;
  const keepPreviousActive = normalizeBoolean_(payload && payload.keepPreviousActive);
  const forcedTenancyId = createNewTenancy ? (payload.forceNewTenancyId || Utilities.getUuid()) : payload.tenancyId;
  const mapped = mapTenantPayload_({ ...(payload || {}), forceNewTenancyId: forcedTenancyId });
  const existingTenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const previousTenancy = existingTenancies.find((t) => t.tenancy_id === previousTenancyId);

  upsertUnique_(TENANTS_SHEET, TENANTS_HEADERS, ['tenant_id'], mapped.tenant);
  upsertUnique_(UNITS_SHEET, UNITS_HEADERS, ['unit_id'], mapped.unit);

  if (createNewTenancy && previousTenancy && !keepPreviousActive) {
    previousTenancy.status = 'ENDED';
    previousTenancy.end_date = previousTenancy.end_date || formatDateIso_(payload.updates?.tenancyEndRaw || new Date());
    upsertUnique_(TENANCIES_SHEET, TENANCIES_HEADERS, ['tenancy_id'], previousTenancy);
  }

  upsertUnique_(TENANCIES_SHEET, TENANCIES_HEADERS, ['tenancy_id'], mapped.tenancy);

  if (previousTenancy && previousTenancy.unit_id && previousTenancy.unit_id !== mapped.tenancy.unit_id && !keepPreviousActive) {
    updateUnitOccupancy_(previousTenancy.unit_id, '', false);
  }
  updateUnitOccupancy_(mapped.tenancy.unit_id, mapped.tenancy.tenancy_id, (mapped.tenancy.status || '').toUpperCase() === 'ACTIVE');

  if (mapped.family) {
    const sheet = getSheetWithHeaders_(FAMILY_SHEET, FAMILY_HEADERS);
    const lastRow = sheet.getLastRow();
    const existing = readTable_(FAMILY_SHEET, FAMILY_HEADERS).filter((row) => row.tenant_id !== mapped.tenant.tenant_id);
    sheet.getRange(2, 1, Math.max(0, lastRow - 1), FAMILY_HEADERS.length).clearContent();
    const retained = existing.concat(mapped.family);
    if (retained.length) {
      const rows = retained.map((r) => FAMILY_HEADERS.map((key) => r[key] ?? ''));
      sheet.getRange(2, 1, rows.length, FAMILY_HEADERS.length).setValues(rows);
    }
  }

  // Only create/update rent revision if rent has actually changed or it's a new tenancy
  if (mapped.rentAmountValue !== null) {
    const effectiveMonth =
      normalizeMonthKey_(mapped.tenancy.commencement_date || mapped.tenancy.agreement_date) || normalizeMonthKey_(new Date());
    if (effectiveMonth) {
      // Get existing revisions for this tenancy
      const existingRevisions = listTenancyRentRevisions_(mapped.tenancy.tenancy_id);
      const existingRevisionForMonth = existingRevisions.find(
        (r) => normalizeMonthKey_(r.effective_month) === effectiveMonth
      );

      // Only create/update revision if:
      // 1. It's a new tenancy (no existing revision for this month), OR
      // 2. The rent amount has actually changed
      const shouldCreateRevision = !existingRevisionForMonth ||
        (Number(existingRevisionForMonth.rent_amount) !== Number(mapped.rentAmountValue));

      if (shouldCreateRevision) {
        const note = existingRevisionForMonth
          ? existingRevisionForMonth.note || ''
          : (existingRevisions.length === 0 ? 'Initial rent' : '');
        upsertTenancyRentRevision_({
          tenancyId: mapped.tenancy.tenancy_id,
          effectiveMonth,
          rentAmount: mapped.rentAmountValue,
          note,
        });
      }
    }
  }

  return jsonResponse({ ok: true, message: 'Tenant updated', tenantId: mapped.tenant.tenant_id, tenancyId: mapped.tenancy.tenancy_id, createdNewTenancy: createNewTenancy });
}

function buildTenantDirectory_() {
  const tenants = readTable_(TENANTS_SHEET, TENANTS_HEADERS);
  const tenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const families = readTable_(FAMILY_SHEET, FAMILY_HEADERS);
  const landlords = readTable_(LANDLORDS_SHEET, LANDLORD_HEADERS);
  const revisions = readTable_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS)
    .map((r) => ({
      ...r,
      effective_month: normalizeMonthKey_(r.effective_month),
      rent_amount: Number(r.rent_amount) || 0,
    }))
    .sort((a, b) => {
      // Sort by effective_month descending (most recent first)
      const monthCompare = (b.effective_month || '').localeCompare(a.effective_month || '');
      if (monthCompare !== 0) return monthCompare;

      // If same month, sort by created_at descending (most recent first)
      const aTime = a.created_at ? new Date(a.created_at).getTime() : 0;
      const bTime = b.created_at ? new Date(b.created_at).getTime() : 0;
      return bTime - aTime;
    });

  const revisionCache = revisions.reduce((m, r) => {
    if (!m[r.tenancy_id]) m[r.tenancy_id] = [];
    m[r.tenancy_id].push(r);
    return m;
  }, {});

  const landlordById = landlords.reduce((m, l) => {
    m[l.landlord_id] = l;
    return m;
  }, {});

  const unitById = units.reduce((m, u) => {
    m[u.unit_id] = u;
    return m;
  }, {});

  const familyByTenant = families.reduce((m, f) => {
    if (!m[f.tenant_id]) m[f.tenant_id] = [];
    m[f.tenant_id].push({
      name: f.name || '',
      relationship: f.relationship || '',
      occupation: f.occupation || '',
      aadhaar: f.aadhaar || '',
      address: f.address || '',
    });
    return m;
  }, {});

  const tenanciesByTenant = tenancies.reduce((m, t) => {
    const list = m[t.tenant_id] || [];
    list.push(t);
    m[t.tenant_id] = list;
    return m;
  }, {});

  return tenancies.map((tenancy) => {
    const tenant = tenants.find((t) => t.tenant_id === tenancy.tenant_id) || {};
    const unit = unitById[tenancy.unit_id] || {};
    const landlord = landlordById[tenancy.landlord_id || unit.landlord_id] || {};
    const templateData = tenancy.templateData || {};
    const tenancyHistory = (tenanciesByTenant[tenancy.tenant_id] || [])
      .slice()
      .sort((a, b) => new Date(b.commencement_date || b.agreement_date || b.created_at || '').getTime() - new Date(a.commencement_date || a.agreement_date || a.created_at || '').getTime())
      .map((t) => ({
        tenancyId: t.tenancy_id,
        unitLabel: buildUnitLabel_(unitById[t.unit_id] || {}),
        startDate: t.commencement_date || t.agreement_date || '',
        endDate: t.end_date || '',
        status: t.status || '',
        grnNumber: t.grn_number || tenant.grn_number || '',
        currentRent: getLatestRentForTenancy_(t.tenancy_id, revisionCache) ?? 0,
      }));
    const latestRent = getLatestRentForTenancy_(tenancy.tenancy_id, revisionCache);
    const resolvedRent = latestRent ?? 0;
    return {
      tenantId: tenant.tenant_id,
      tenancyId: tenancy.tenancy_id,
      grnNumber: tenancy.grn_number || tenant.grn_number || '',
      tenantFullName: tenant.full_name || '',
      tenantOccupation: tenant.occupation || '',
      tenantPermanentAddress: tenant.permanent_address || '',
      tenantAadhaar: tenant.aadhaar || '',
      tenantMobile: tenant.mobile || '',
      wing: unit.wing || '',
      unitId: unit.unit_id || '',
      unitNumber: unit.unit_number || '',
      floor: unit.floor || '',
      direction: unit.direction || '',
      meterNumber: unit.meter_number || '',
      landlordId: landlord.landlord_id || '',
      landlordName: landlord.name || '',
      landlordAadhaar: landlord.aadhaar || '',
      landlordAddress: landlord.address || '',
      unitOccupied: normalizeBoolean_(unit.is_occupied),
      rentAmount: resolvedRent,
      currentRent: resolvedRent,
      payableDate: tenancy.rent_payable_day || '',
      securityDeposit: tenancy.security_deposit || '',
      rentIncrease: tenancy.rent_increase_amount || '',
      rentRevisionNumber: tenancy.rent_revision_number || '',
      rentRevisionUnit: tenancy.rent_revision_unit || '',
      tenantNoticeMonths: tenancy.tenant_notice_months || '',
      landlordNoticeMonths: tenancy.landlord_notice_months || '',
      lateRentPerDay: tenancy.late_rent_per_day || '',
      lateGracePeriodDays: tenancy.late_grace_days || '',
      agreementDate: tenancy.agreement_date || '',
      tenancyCommencement: tenancy.commencement_date || '',
      tenancyEndDate: tenancy.end_date || '',
      activeTenant: (tenancy.status || '').toString().toUpperCase() === 'ACTIVE',
      vacateReason: tenancy.vacate_reason || '',
      family: familyByTenant[tenant.tenant_id] || [],
      tenancyHistory,
      templateData: {
        tenant_id: tenant.tenant_id,
        tenancy_id: tenancy.tenancy_id,
        unit_id: unit.unit_id,
        'GRN number': tenancy.grn_number || tenant.grn_number || '',
        Tenant_Full_Name: tenant.full_name || '',
        Tenant_Permanent_Address: tenant.permanent_address || '',
        tenant_Aadhar: tenant.aadhaar || '',
        tenant_mobile: tenant.mobile || '',
        tenant_occupation: tenant.occupation || '',
        wing: unit.wing || '',
        unit_number: unit.unit_number || '',
        floor_of_building: unit.floor || '',
        direction_build: unit.direction || '',
        meter_number: unit.meter_number || '',
        landlord_id: landlord.landlord_id || '',
        Landlord_name: landlord.name || templateData.Landlord_name || '',
        landlord_address: landlord.address || templateData.landlord_address || '',
        landlord_aadhar: landlord.aadhaar || templateData.landlord_aadhar || '',
        rent_amount: resolvedRent,
        payable_date_raw: tenancy.rent_payable_day || '',
        secu_depo: tenancy.security_deposit || '',
        rent_inc: tenancy.rent_increase_amount || '',
        rent_rev_number: tenancy.rent_revision_number || '',
        'rent_rev year_mon': tenancy.rent_revision_unit || '',
        notice_num_t: tenancy.tenant_notice_months || '',
        notice_num_l: tenancy.landlord_notice_months || '',
        pet_text_area: tenancy.pet_policy || '',
        late_rent: tenancy.late_rent_per_day || '',
        late_days: tenancy.late_grace_days || '',
        agreement_date_raw: tenancy.agreement_date || '',
        tenancy_comm_raw: tenancy.commencement_date || '',
        tenancy_end_raw: tenancy.end_date || '',
      },
    };
  });
}

function buildUnitLabel_(unit) {
  if (!unit) return '';
  const parts = [unit.wing, unit.unit_number].filter(Boolean);
  return parts.join(' - ') || unit.unit_id || '';
}

/********* TENANCY RENT REVISIONS *********/
function listTenancyRentRevisions_(tenancyId) {
  if (!tenancyId) return [];
  return readTable_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS)
    .filter((r) => r.tenancy_id === tenancyId)
    .map((r) => ({
      ...r,
      effective_month: normalizeMonthKey_(r.effective_month),
      rent_amount: Number(r.rent_amount) || 0,
    }))
    .sort((a, b) => {
      // Sort by effective_month descending (most recent first)
      const monthCompare = (b.effective_month || '').localeCompare(a.effective_month || '');
      if (monthCompare !== 0) return monthCompare;

      // If same month, sort by created_at descending (most recent first)
      const aTime = a.created_at ? new Date(a.created_at).getTime() : 0;
      const bTime = b.created_at ? new Date(b.created_at).getTime() : 0;
      return bTime - aTime;
    });
}

function upsertTenancyRentRevision_(payload = {}) {
  const tenancyId = payload.tenancyId || payload.tenancy_id;
  const effectiveMonth = normalizeMonthKey_(payload.effectiveMonth || payload.effective_month);
  const rentAmount = Number(payload.rentAmount ?? payload.rent_amount);
  const note = payload.note || '';

  if (!tenancyId) throw new Error('tenancyId required');
  if (!effectiveMonth || !/^\d{4}-\d{2}$/.test(effectiveMonth)) throw new Error('Invalid effective month');
  if (isNaN(rentAmount) || rentAmount < 0) throw new Error('Invalid rent amount');

  const record = {
    revision_id: payload.revision_id || payload.revisionId || Utilities.getUuid(),
    tenancy_id: tenancyId,
    effective_month: effectiveMonth,
    rent_amount: rentAmount,
    note,
    created_at: payload.created_at || new Date(),
  };

  upsertUnique_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS, ['tenancy_id', 'effective_month'], record);
  return record;
}

function deleteTenancyRentRevision_(revisionId) {
  if (!revisionId) return false;
  const remaining = readTable_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS).filter(
    (r) => r.revision_id !== revisionId
  );
  const sheet = getSheetWithHeaders_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, TENANCY_RENT_REVISION_HEADERS.length).clearContent();
  if (remaining.length) {
    const rows = remaining.map((r) => TENANCY_RENT_REVISION_HEADERS.map((key) => r[key] ?? ''));
    sheet.getRange(2, 1, rows.length, TENANCY_RENT_REVISION_HEADERS.length).setValues(rows);
  }
  return true;
}

function getLatestRentForTenancy_(tenancyId, revisionCache) {
  if (!tenancyId) return null;
  const revisions = revisionCache ? revisionCache[tenancyId] || [] : listTenancyRentRevisions_(tenancyId);
  if (!revisions.length) return null;

  let latestMonth = '';
  let latestCreated = 0;
  let latestAmount = null;

  revisions.forEach((rev) => {
    const effective = normalizeMonthKey_(rev.effective_month);
    const createdAt = rev.created_at ? new Date(rev.created_at).getTime() : 0;
    const monthCompare = (effective || '').localeCompare(latestMonth || '');
    if (latestAmount === null || monthCompare > 0 || (monthCompare === 0 && createdAt > latestCreated)) {
      latestMonth = effective || '';
      latestCreated = createdAt;
      const amount = Number(rev.rent_amount);
      latestAmount = isNaN(amount) ? 0 : amount;
    }
  });

  return latestAmount;
}

/********* BILLING *********/
function computeMotorShare_(config, includedCount) {
  const motorUnits = Number(config.motor_new || 0) - Number(config.motor_prev || 0);
  const rate = Number(config.electricity_rate || 0);
  if (!includedCount || includedCount <= 0) return 0;
  return Math.round(((motorUnits * rate) / includedCount) * 100) / 100;
}

function roundToTwo_(value) {
  const num = Number(value || 0);
  if (isNaN(num)) return 0;
  return Math.round(num * 100) / 100;
}

function roundToNearest_(value) {
  const num = Number(value || 0);
  if (isNaN(num)) return 0;
  return roundToTwo_(Math.round(num));
}

function handleSaveBillingRecord_(payload) {
  const monthKey = normalizeMonthKey_(payload.monthKey || '');
  const wing = (payload.wing || '').toString().trim();
  if (!monthKey || !wing) return jsonResponse({ ok: false, error: 'Missing month or wing' });

  const wingNormalized = wing.toLowerCase();
  const unitById = readTable_(UNITS_SHEET, UNITS_HEADERS).reduce((m, u) => {
    m[u.unit_id] = u;
    return m;
  }, {});

  const meta = payload.meta || {};
  const wingConfig = {
    wing_month_id: Utilities.getUuid(),
    month_key: monthKey,
    wing,
    electricity_rate: meta.electricityRate || '',
    sweeping_per_flat: meta.sweepingPerFlat || '',
    motor_prev: meta.motorPrev || '',
    motor_new: meta.motorNew || '',
    motor_units: (Number(meta.motorNew || 0) || 0) - (Number(meta.motorPrev || 0) || 0),
    created_at: new Date(),
  };
  upsertUnique_(WING_MONTHLY_SHEET, WING_MONTHLY_HEADERS, ['month_key', 'wing'], wingConfig);

  const tenants = Array.isArray(payload.tenants) ? payload.tenants : [];
  const tenancyById = {};
  const tenancyByTenant = {};
  const tenancyByGrn = {};
  readTable_(TENANCIES_SHEET, TENANCIES_HEADERS).forEach((t) => {
    tenancyById[t.tenancy_id] = t;
    if (!tenancyByTenant[t.tenant_id]) tenancyByTenant[t.tenant_id] = [];
    tenancyByTenant[t.tenant_id].push(t);
    const grnKey = (t.grn_number || '').toString().toLowerCase();
    if (grnKey) {
      if (!tenancyByGrn[grnKey]) tenancyByGrn[grnKey] = [];
      tenancyByGrn[grnKey].push(t);
    }
  });

  const rentRevisionCache = readTable_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS).reduce((m, r) => {
    const tenancyId = r.tenancy_id;
    const effectiveMonth = normalizeMonthKey_(r.effective_month);
    if (!tenancyId || !effectiveMonth) return m;
    if (!m[tenancyId]) m[tenancyId] = [];
    m[tenancyId].push({
      ...r,
      effective_month: effectiveMonth,
      rent_amount: Number(r.rent_amount) || 0,
    });
    return m;
  }, {});
  Object.values(rentRevisionCache).forEach((list) =>
    list.sort((a, b) => {
      // Sort by effective_month descending (most recent first)
      const monthCompare = (b.effective_month || '').localeCompare(a.effective_month || '');
      if (monthCompare !== 0) return monthCompare;

      // If same month, sort by created_at descending (most recent first)
      const aTime = a.created_at ? new Date(a.created_at).getTime() : 0;
      const bTime = b.created_at ? new Date(b.created_at).getTime() : 0;
      return bTime - aTime;
    })
  );

  const readingRows = [];
  const billRows = [];
  const includedTenancies = [];

  tenants.forEach((tenant) => {
    const grnKey = (tenant.grn || tenant.tenantKey || tenant.tenantName || '').toString().toLowerCase();
    const tenancyFromPayload = tenant.tenancyId && tenancyById[tenant.tenancyId];
    const tenancyOptions = tenancyFromPayload
      ? [tenancyFromPayload]
      : tenant.tenantId && tenancyByTenant[tenant.tenantId]
        ? tenancyByTenant[tenant.tenantId]
        : tenancyByGrn[grnKey]
          ? tenancyByGrn[grnKey]
          : [];
    const tenancyMatchesWing = (t) => {
      const unit = unitById[t.unit_id];
      if (!unit) return false;
      return (unit.wing || '').toString().trim().toLowerCase() === wingNormalized;
    };
    const matchingTenancies = tenancyOptions.filter(
      (t) => tenancyMatchesWing(t) && (t.status || '').toUpperCase() === 'ACTIVE'
    );
    const resolvedTenancies = matchingTenancies.length
      ? matchingTenancies
      : tenancyOptions.filter(tenancyMatchesWing);
    if (!resolvedTenancies.length) return;

    const included = normalizeBoolean_(tenant.included);
    const prev = Number(tenant.prevReading || 0) || 0;
    const next = Number(tenant.newReading || 0) || 0;
    const units = Math.max(next - prev, 0);
    resolvedTenancies.forEach((tenancy) => {
      if (included) includedTenancies.push(tenancy.tenancy_id);

      const reading = {
        reading_id: Utilities.getUuid(),
        month_key: monthKey,
        tenancy_id: tenancy.tenancy_id,
        prev_reading: tenant.prevReading || '',
        new_reading: tenant.newReading || '',
        included,
        override_rent: tenant.override_rent || tenant.rentAmount || '',
        notes: tenant.notes || '',
        created_at: new Date(),
      };
      readingRows.push(reading);
    });
  });

  const motorPerTenant = computeMotorShare_(wingConfig, includedTenancies.length);
  const rate = Number(wingConfig.electricity_rate || 0);
  const sweep = Number(wingConfig.sweeping_per_flat || 0);
  const payableDate = (tenant) => tenant.rent_payable_day || '';

  readingRows.forEach((reading) => {
    const tenancy = tenancyById[reading.tenancy_id];
    if (!tenancy) return;

    const latestRent = getLatestRentForTenancy_(tenancy.tenancy_id, rentRevisionCache);

    // Prioritize override_rent, then latest rent from revisions
    const rent = reading.override_rent ? Number(reading.override_rent) : (latestRent || 0);
    const units = Math.max(Number(reading.new_reading || 0) - Number(reading.prev_reading || 0), 0);
    const electricityAmount = roundToTwo_(units * rate);
    const sweepAmount = normalizeBoolean_(reading.included) ? roundToTwo_(sweep) : 0;
    const motorShare = normalizeBoolean_(reading.included) ? roundToTwo_(motorPerTenant) : 0;
    const totalBeforeRound = Number(rent) + electricityAmount + sweepAmount + motorShare;
    const total = normalizeBoolean_(reading.included) ? roundToNearest_(totalBeforeRound) : 0;

    billRows.push({
      bill_line_id: Utilities.getUuid(),
      month_key: monthKey,
      tenancy_id: reading.tenancy_id,
      rent_amount: rent,
      electricity_units: units,
      electricity_amount: electricityAmount,
      motor_share_amount: motorShare,
      sweep_amount: sweepAmount,
      total_amount: total,
      payable_date: payableDate(tenancy),
      generated_at: new Date(),
      amount_paid: 0,
      is_paid: total <= 0,
    });
  });

  // Persist readings and bills (replace existing month+tenancy combos)
  persistUniqueByKeys_(TENANT_READINGS_SHEET, TENANT_READING_HEADERS, ['month_key', 'tenancy_id'], readingRows);
  persistUniqueByKeys_(BILL_LINES_SHEET, BILL_LINE_HEADERS, ['month_key', 'tenancy_id'], billRows);

  const { bills, coverage } = handleFetchGeneratedBills_();
  return jsonResponse({ ok: true, message: 'Billing saved', bills, coverage });
}

function handleGetBillingRecord_(monthKeyRaw, wingRaw) {
  const monthKey = normalizeMonthKey_(monthKeyRaw || '');
  const wing = (wingRaw || '').toString().trim();
  const wingNormalized = wing.toLowerCase();
  if (!monthKey || !wing) return jsonResponse({ ok: false, error: 'Missing month or wing' });

  const wingConfigs = readTable_(WING_MONTHLY_SHEET, WING_MONTHLY_HEADERS);
  const configRow = wingConfigs.find(
    (c) => normalizeMonthKey_(c.month_key) === monthKey && (c.wing || '').toString().trim().toLowerCase() === wingNormalized
  );
  const hasConfig = !!configRow;
  const config = configRow || {
    month_key: monthKey,
    wing,
    electricity_rate: '',
    sweeping_per_flat: '',
    motor_prev: '',
    motor_new: '',
    motor_units: '',
  };

  const tenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const tenants = readTable_(TENANTS_SHEET, TENANTS_HEADERS);
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const rentRevisionCache = readTable_(TENANCY_RENT_REVISIONS_SHEET, TENANCY_RENT_REVISION_HEADERS).reduce((m, r) => {
    const tenancyId = r.tenancy_id;
    const effectiveMonth = normalizeMonthKey_(r.effective_month);
    if (!tenancyId || !effectiveMonth) return m;
    if (!m[tenancyId]) m[tenancyId] = [];
    m[tenancyId].push({
      ...r,
      effective_month: effectiveMonth,
      rent_amount: Number(r.rent_amount) || 0,
    });
    return m;
  }, {});
  Object.values(rentRevisionCache).forEach((list) =>
    list.sort((a, b) => {
      // Sort by effective_month descending (most recent first)
      const monthCompare = (b.effective_month || '').localeCompare(a.effective_month || '');
      if (monthCompare !== 0) return monthCompare;

      // If same month, sort by created_at descending (most recent first)
      const aTime = a.created_at ? new Date(a.created_at).getTime() : 0;
      const bTime = b.created_at ? new Date(b.created_at).getTime() : 0;
      return bTime - aTime;
    })
  );

  const unitMap = units.reduce((m, u) => {
    if ((u.wing || '').toString().trim().toLowerCase() === wingNormalized) m[u.unit_id] = u;
    return m;
  }, {});

  const tenancyMap = tenancies.reduce((m, t) => {
    if (unitMap[t.unit_id]) m[t.tenancy_id] = t;
    return m;
  }, {});

  const readings = readTable_(TENANT_READINGS_SHEET, TENANT_READING_HEADERS).filter((r) => {
    const matchesMonth = normalizeMonthKey_(r.month_key) === monthKey;
    const matchesWing = !!tenancyMap[r.tenancy_id];
    return matchesMonth && matchesWing;
  });
  const hasReadings = readings.length > 0;

  const tenantMap = tenants.reduce((m, t) => {
    m[t.tenant_id] = t;
    return m;
  }, {});
  const tenantsPayload = readings.map((reading) => {
    const tenancy = tenancyMap[reading.tenancy_id] || {};
    const tenant = tenantMap[tenancy.tenant_id] || {};
    const unit = unitMap[tenancy.unit_id] || {};
    const latestRent = getLatestRentForTenancy_(reading.tenancy_id, rentRevisionCache);
    return {
      tenancyId: reading.tenancy_id,
      tenantKey: tenancy.grn_number || tenant.full_name || '',
      tenantName: tenant.full_name || '',
      wing: unit.wing || '',
      unitNumber: unit.unit_number || '',
      prevReading: reading.prev_reading || '',
      newReading: reading.new_reading || '',
      included: normalizeBoolean_(reading.included),
      override_rent: reading.override_rent || '',
      rentAmount: reading.override_rent || latestRent || '',
      payableDate: tenancy.rent_payable_day || '',
      direction: unit.direction || '',
      floor: unit.floor || '',
      meterNumber: unit.meter_number || '',
    };
  });

  return jsonResponse({
    ok: true,
    monthKey,
    monthLabel: formatMonthLabelForDisplay_(monthKey),
    wing: config.wing || wing,
    hasConfig,
    hasReadings,
    meta: {
      month_key: config.month_key,
      wing: config.wing,
      electricityRate: config.electricity_rate || '',
      sweepingPerFlat: config.sweeping_per_flat || '',
      motorPrev: config.motor_prev || '',
      motorNew: config.motor_new || '',
      motor_units: config.motor_units || '',
    },
    tenants: tenantsPayload,
  });
}

function persistUniqueByKeys_(sheetName, headers, keys, records) {
  const sheet = getSheetWithHeaders_(sheetName, headers);
  const headerIndex = buildHeaderIndex_(headers);
  const existing = readTable_(sheetName, headers);
  const normalizeKeyValue = (key, value) => {
    if (key === 'month_key') return normalizeMonthKey_(value);
    return value;
  };

  const filtered = existing.filter((row) => {
    return !records.some((incoming) =>
      keys.every((k) => normalizeKeyValue(k, incoming[k]) == normalizeKeyValue(k, row[k]))
    );
  });
  const merged = filtered.concat(records);
  const rows = merged.map((r) => headers.map((key) => r[key] ?? ''));
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function deriveBillPaymentState_(bill) {
  const totalAmount = Number(bill.total_amount) || 0;
  const amountPaidRaw = bill.amount_paid;
  const isPaidRaw = bill.is_paid;
  const hasAmountPaid = amountPaidRaw !== '' && amountPaidRaw !== null && typeof amountPaidRaw !== 'undefined';
  const hasIsPaid = isPaidRaw !== '' && isPaidRaw !== null && typeof isPaidRaw !== 'undefined';
  const amountPaid = hasAmountPaid ? Number(amountPaidRaw) || 0 : null;
  const isPaid = hasIsPaid ? normalizeBoolean_(isPaidRaw) : null;
  const paidByTotal = totalAmount <= 0;
  const paidByAmount = amountPaid !== null && amountPaid + 0.005 >= totalAmount;
  let derivedIsPaid = isPaid;
  if (derivedIsPaid === null) {
    if (amountPaid !== null) {
      derivedIsPaid = paidByTotal || paidByAmount;
    } else if (paidByTotal) {
      derivedIsPaid = true;
    }
  }
  return { totalAmount, amountPaid, isPaid: derivedIsPaid };
}

function handleFetchBillsMinimal_(statusRaw, monthsBackRaw) {
  const status = (statusRaw || '').toString().trim().toLowerCase();
  const monthsBack = Number(monthsBackRaw) || 0;
  const billColumns = [
    'bill_line_id',
    'month_key',
    'tenancy_id',
    'total_amount',
    'amount_paid',
    'is_paid',
    'payable_date',
  ];
  const bills = readTableColumns_(BILL_LINES_SHEET, BILL_LINE_HEADERS, billColumns);
  const filteredBills = (status === 'pending' || status === 'paid'
    ? bills.filter((bill) => {
      const state = deriveBillPaymentState_(bill);
      return status === 'paid' ? state.isPaid === true : state.isPaid !== true;
    })
    : bills).filter((bill) => isWithinRecentMonths_(bill.month_key, monthsBack));

  if (!filteredBills.length) {
    return { bills: [] };
  }

  const tenancyColumns = ['tenancy_id', 'tenant_id', 'unit_id', 'grn_number'];
  const tenantColumns = ['tenant_id', 'full_name'];
  const unitColumns = ['unit_id', 'wing', 'unit_number'];

  const tenancyMap = getCachedLookup_(
    `lookup:tenancies:${tenancyColumns.join(',')}`,
    () => buildLookupByKey_(readTableColumns_(TENANCIES_SHEET, TENANCIES_HEADERS, tenancyColumns), 'tenancy_id'),
    LOOKUP_CACHE_TTL_SECONDS
  );
  const tenantMap = getCachedLookup_(
    `lookup:tenants:${tenantColumns.join(',')}`,
    () => buildLookupByKey_(readTableColumns_(TENANTS_SHEET, TENANTS_HEADERS, tenantColumns), 'tenant_id'),
    LOOKUP_CACHE_TTL_SECONDS
  );
  const unitMap = getCachedLookup_(
    `lookup:units:${unitColumns.join(',')}`,
    () => buildLookupByKey_(readTableColumns_(UNITS_SHEET, UNITS_HEADERS, unitColumns), 'unit_id'),
    LOOKUP_CACHE_TTL_SECONDS
  );

  const billPayload = filteredBills.map((bill) => {
    const tenancy = tenancyMap[bill.tenancy_id] || {};
    const tenant = tenantMap[tenancy.tenant_id] || {};
    const unit = unitMap[tenancy.unit_id] || {};
    const paymentState = deriveBillPaymentState_(bill);
    const totalAmount = paymentState.totalAmount;
    const amountPaid = paymentState.amountPaid;
    const isPaid = paymentState.isPaid;
    const remainingAmount = amountPaid !== null
      ? Math.max(0, totalAmount - amountPaid)
      : (isPaid === true ? 0 : null);

    return {
      monthKey: bill.month_key,
      monthLabel: formatMonthLabelForDisplay_(bill.month_key),
      wing: unit.wing || '',
      unitNumber: unit.unit_number || '',
      tenantKey: tenancy.grn_number || tenant.full_name || '',
      tenantName: tenant.full_name || '',
      totalAmount,
      amountPaid,
      remainingAmount,
      isPaid,
      payableDate: bill.payable_date || '',
      billLineId: bill.bill_line_id,
      tenancyId: bill.tenancy_id,
    };
  });

  return { bills: billPayload };
}

function handleFetchBillDetails_(billLineIdRaw) {
  const billLineId = (billLineIdRaw || '').toString().trim();
  if (!billLineId) return { ok: false, error: 'Missing billLineId' };

  const bills = readTable_(BILL_LINES_SHEET, BILL_LINE_HEADERS);
  const bill = bills.find((b) => b.bill_line_id === billLineId);
  if (!bill) return { ok: false, error: 'Bill not found' };

  const tenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const tenants = readTable_(TENANTS_SHEET, TENANTS_HEADERS);
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const readings = readTable_(TENANT_READINGS_SHEET, TENANT_READING_HEADERS);
  const config = readTable_(WING_MONTHLY_SHEET, WING_MONTHLY_HEADERS);

  const tenancy = tenancies.find((t) => t.tenancy_id === bill.tenancy_id) || {};
  const tenant = tenants.find((t) => t.tenant_id === tenancy.tenant_id) || {};
  const unit = units.find((u) => u.unit_id === tenancy.unit_id) || {};

  const normalizedMonth = normalizeMonthKey_(bill.month_key);
  const reading = readings.find((r) =>
    r.tenancy_id === bill.tenancy_id && normalizeMonthKey_(r.month_key) === normalizedMonth
  ) || {};
  const cfg = config.find((c) =>
    normalizeMonthKey_(c.month_key) === normalizedMonth &&
    (c.wing || '').toString().trim().toLowerCase() === (unit.wing || '').toString().trim().toLowerCase()
  ) || {};

  const paymentState = deriveBillPaymentState_(bill);
  const totalAmount = paymentState.totalAmount;
  const amountPaid = paymentState.amountPaid;
  const isPaid = paymentState.isPaid;
  const remainingAmount = amountPaid !== null
    ? Math.max(0, totalAmount - amountPaid)
    : (isPaid === true ? 0 : null);

  return {
    ok: true,
    bill: {
      monthKey: bill.month_key,
      monthLabel: formatMonthLabelForDisplay_(bill.month_key),
      wing: unit.wing || '',
      tenantKey: tenancy.grn_number || tenant.full_name || '',
      tenantName: tenant.full_name || '',
      rentAmount: Number(bill.rent_amount) || 0,
      electricityAmount: Number(bill.electricity_amount) || 0,
      motorShare: Number(bill.motor_share_amount) || 0,
      sweepAmount: Number(bill.sweep_amount) || 0,
      totalAmount,
      amountPaid,
      remainingAmount,
      isPaid,
      included: normalizeBoolean_(reading.included),
      payableDate: bill.payable_date || '',
      prevReading: reading.prev_reading || '',
      newReading: reading.new_reading || '',
      electricityRate: cfg.electricity_rate || '',
      sweepingPerFlat: cfg.sweeping_per_flat || '',
      motorPrev: cfg.motor_prev || '',
      motorNew: cfg.motor_new || '',
      billLineId: bill.bill_line_id,
      tenancyId: bill.tenancy_id,
    },
  };
}

function handleFetchGeneratedBills_(statusRaw) {
  const bills = readTable_(BILL_LINES_SHEET, BILL_LINE_HEADERS);
  const tenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const tenants = readTable_(TENANTS_SHEET, TENANTS_HEADERS);
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const readings = readTable_(TENANT_READINGS_SHEET, TENANT_READING_HEADERS);
  const config = readTable_(WING_MONTHLY_SHEET, WING_MONTHLY_HEADERS);

  const tenancyMap = tenancies.reduce((m, t) => {
    m[t.tenancy_id] = t;
    return m;
  }, {});
  const tenantMap = tenants.reduce((m, t) => {
    m[t.tenant_id] = t;
    return m;
  }, {});
  const unitMap = units.reduce((m, u) => {
    m[u.unit_id] = u;
    return m;
  }, {});
  const readingMap = readings.reduce((m, r) => {
    m[`${r.month_key}__${r.tenancy_id}`] = r;
    return m;
  }, {});
  const configMap = config.reduce((m, c) => {
    m[`${c.month_key}__${c.wing}`] = c;
    return m;
  }, {});

  const coverage = config
    .map((c) => ({
      monthKey: normalizeMonthKey_(c.month_key),
      wing: (c.wing || '').toString().trim(),
    }))
    .filter((c) => c.monthKey && c.wing);

  const status = (statusRaw || '').toString().trim().toLowerCase();
  const filteredBills = status === 'pending' || status === 'paid'
    ? bills.filter((bill) => {
      const state = deriveBillPaymentState_(bill);
      return status === 'paid' ? state.isPaid === true : state.isPaid !== true;
    })
    : bills;

  const billPayload = filteredBills.map((bill) => {
    const tenancy = tenancyMap[bill.tenancy_id] || {};
    const tenant = tenantMap[tenancy.tenant_id] || {};
    const unit = unitMap[tenancy.unit_id] || {};
    const reading = readingMap[`${bill.month_key}__${bill.tenancy_id}`] || {};
    const cfg = configMap[`${bill.month_key}__${unit.wing || ''}`] || {};
    const monthLabel = formatMonthLabelForDisplay_(bill.month_key);
    const paymentState = deriveBillPaymentState_(bill);
    const totalAmount = paymentState.totalAmount;
    const amountPaid = paymentState.amountPaid;
    const isPaid = paymentState.isPaid;
    const remainingAmount = amountPaid !== null
      ? Math.max(0, totalAmount - amountPaid)
      : (isPaid === true ? 0 : null);
    return {
      monthKey: bill.month_key,
      monthLabel,
      wing: unit.wing || '',
      unitNumber: unit.unit_number || '',
      tenantKey: tenancy.grn_number || tenant.full_name || '',
      tenantName: tenant.full_name || '',
      rentAmount: Number(bill.rent_amount) || 0,
      electricityAmount: Number(bill.electricity_amount) || 0,
      motorShare: Number(bill.motor_share_amount) || 0,
      sweepAmount: Number(bill.sweep_amount) || 0,
      totalAmount,
      amountPaid,
      remainingAmount,
      isPaid,
      included: normalizeBoolean_(reading.included),
      payableDate: bill.payable_date || '',
      prevReading: reading.prev_reading || '',
      newReading: reading.new_reading || '',
      electricityRate: cfg.electricity_rate || '',
      sweepingPerFlat: cfg.sweeping_per_flat || '',
      motorPrev: cfg.motor_prev || '',
      motorNew: cfg.motor_new || '',
      billLineId: bill.bill_line_id,
      tenancyId: bill.tenancy_id,
    };
  });

  return { bills: billPayload, coverage };
}

/********* PAYMENTS *********/
function savePaymentAttachment_(dataUrl, paymentId, originalName, tenantName, monthKey) {
  if (!dataUrl) return { attachment_id: '', attachmentName: '', attachmentUrl: '' };
  const match = dataUrl.match(/^data:([^;]+);base64,(.*)$/);
  if (!match) return { attachment_id: '', attachmentName: '', attachmentUrl: '' };
  const mimeType = match[1] || 'application/octet-stream';
  const cleanBase64 = match[2].replace(/\s/g, '');
  const bytes = Utilities.base64Decode(cleanBase64);
  const safeTenant = sanitizeFileSegment_(tenantName, 'Tenant');
  const normalizedMonthKey = normalizeMonthKey_(monthKey);
  const safeMonth = normalizedMonthKey || sanitizeFileSegment_(monthKey, 'month');
  const safePayment = sanitizeFileSegment_(paymentId || '', 'payment');
  const blobName = `${safeTenant}_${safeMonth}_${safePayment}`;
  const blob = Utilities.newBlob(bytes, mimeType).setName(blobName);
  const folder = getOrCreateFolder_('Payment Proofs');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const ensured = ensureDriveFileShareable_(file.getUrl());
  const attachment = {
    attachment_id: Utilities.getUuid(),
    file_drive_id: file.getId(),
    file_name: file.getName(),
    file_url: ensured.viewUrl,
    uploaded_at: new Date(),
  };
  upsertUnique_(ATTACHMENTS_SHEET, ATTACHMENT_HEADERS, ['attachment_id'], attachment);
  return { attachment_id: attachment.attachment_id, attachmentName: attachment.file_name, attachmentUrl: attachment.file_url };
}

function deleteAttachmentById_(attachmentId) {
  if (!attachmentId) return false;
  const rows = readTable_(ATTACHMENTS_SHEET, ATTACHMENT_HEADERS);
  const remaining = rows.filter((r) => r.attachment_id !== attachmentId);
  const target = rows.find((r) => r.attachment_id === attachmentId);

  if (target && target.file_drive_id) {
    try {
      const file = DriveApp.getFileById(target.file_drive_id);
      file.setTrashed(true);
    } catch (err) {
      // Ignore drive deletion errors
    }
  }

  const sheet = getSheetWithHeaders_(ATTACHMENTS_SHEET, ATTACHMENT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, ATTACHMENT_HEADERS.length).clearContent();
  if (remaining.length) {
    const rowsOut = remaining.map((r) => ATTACHMENT_HEADERS.map((key) => r[key] ?? ''));
    sheet.getRange(2, 1, rowsOut.length, ATTACHMENT_HEADERS.length).setValues(rowsOut);
  }
  return true;
}

function getOrCreateFolder_(name) {
  const existing = DriveApp.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(name);
}

function updateBillPaymentStatus_(billLineIds) {
  const list = Array.isArray(billLineIds) ? billLineIds : [billLineIds];
  const ids = list
    .map((id) => (id || '').toString().trim())
    .filter(Boolean);
  if (!ids.length) return;

  const idSet = ids.reduce((m, id) => {
    m[id] = true;
    return m;
  }, {});
  const totals = ids.reduce((m, id) => {
    m[id] = 0;
    return m;
  }, {});

  const payments = readTableColumns_(PAYMENTS_SHEET, PAYMENT_HEADERS, ['bill_line_id', 'amount']);
  payments.forEach((payment) => {
    const id = payment.bill_line_id;
    if (!idSet[id]) return;
    totals[id] += Number(payment.amount) || 0;
  });

  const sheet = getSheetWithHeaders_(BILL_LINES_SHEET, BILL_LINE_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const headerIndex = buildHeaderIndex_(BILL_LINE_HEADERS);
  const idCol = headerIndex.bill_line_id + 1;
  const totalCol = headerIndex.total_amount + 1;
  const amountPaidCol = headerIndex.amount_paid + 1;
  const isPaidCol = headerIndex.is_paid + 1;
  const rowCount = lastRow - 1;
  const idValues = sheet.getRange(2, idCol, rowCount, 1).getValues();
  const totalValues = sheet.getRange(2, totalCol, rowCount, 1).getValues();

  for (let i = 0; i < rowCount; i += 1) {
    const id = idValues[i][0];
    if (!idSet[id]) continue;
    const total = Number(totalValues[i][0]) || 0;
    const paid = Math.round((totals[id] || 0) * 100) / 100;
    const isPaid = total <= 0 || paid + 0.005 >= total;
    sheet.getRange(i + 2, amountPaidCol, 1, 2).setValues([[paid, isPaid]]);
  }
}

function handleSavePayment_(payload = {}) {
  const paymentId = payload.id || Utilities.getUuid();
  const paymentDate = parseIsoDate_(payload.date) || new Date();
  let billLineId = payload.billLineId || '';
  let tenantId = payload.tenantId || '';
  let resolvedMonthKey = payload.monthKey || '';
  let resolvedTenantName = payload.tenantName || '';

  // Fallback: If billLineId is missing but we have tenancy+month, find the bill
  if (!billLineId && payload.tenancyId && payload.monthKey) {
    const bills = readTable_(BILL_LINES_SHEET, BILL_LINE_HEADERS);
    const match = bills.find(b =>
      b.tenancy_id === payload.tenancyId &&
      normalizeMonthKey_(b.month_key) === normalizeMonthKey_(payload.monthKey)
    );
    if (match) billLineId = match.bill_line_id;
  }

  if (billLineId) {
    const billLookup = readTableColumns_(BILL_LINES_SHEET, BILL_LINE_HEADERS, ['bill_line_id', 'month_key', 'tenancy_id']);
    const billMatch = billLookup.find((b) => b.bill_line_id == billLineId);
    if (billMatch) {
      resolvedMonthKey = normalizeMonthKey_(billMatch.month_key) || resolvedMonthKey || '';
      if (!tenantId || !resolvedTenantName) {
        const tenancyLookup = readTableColumns_(TENANCIES_SHEET, TENANCIES_HEADERS, ['tenancy_id', 'tenant_id']);
        const tenancyMatch = tenancyLookup.find((t) => t.tenancy_id == billMatch.tenancy_id);
        if (tenancyMatch) {
          tenantId = tenantId || tenancyMatch.tenant_id || '';
          if (!resolvedTenantName) {
            const tenantLookup = readTableColumns_(TENANTS_SHEET, TENANTS_HEADERS, ['tenant_id', 'full_name']);
            const tenantMatch = tenantLookup.find((t) => t.tenant_id == tenancyMatch.tenant_id);
            resolvedTenantName = (tenantMatch && tenantMatch.full_name) || resolvedTenantName;
          }
        }
      }
    }
  }
  const attachment = payload.attachmentDataUrl
    ? savePaymentAttachment_(
      payload.attachmentDataUrl,
      paymentId,
      payload.attachmentName || '',
      resolvedTenantName,
      resolvedMonthKey
    )
    : {
      attachment_id: payload.attachmentId || '',
      attachmentName: payload.attachmentName || '',
      attachmentUrl: payload.attachmentUrl || '',
    };

  const record = {
    payment_id: paymentId,
    payment_date: paymentDate,
    bill_line_id: billLineId,
    tenant_id: tenantId,
    amount: Number(payload.amount) || 0,
    mode: payload.mode || '',
    reference: payload.reference || '',
    notes: payload.notes || '',
    attachment_id: attachment.attachment_id || '',
    created_at: new Date(),
  };

  upsertUnique_(PAYMENTS_SHEET, PAYMENT_HEADERS, ['payment_id'], record);
  if (billLineId) {
    updateBillPaymentStatus_([billLineId]);
  }

  return { ok: true, payment: mapPaymentRow_(record) };
}

function mapPaymentRow_(record) {
  const bills = readTable_(BILL_LINES_SHEET, BILL_LINE_HEADERS);
  const tenancies = readTable_(TENANCIES_SHEET, TENANCIES_HEADERS);
  const tenants = readTable_(TENANTS_SHEET, TENANTS_HEADERS);
  const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
  const attachmentLookup = readTable_(ATTACHMENTS_SHEET, ATTACHMENT_HEADERS).reduce((m, a) => {
    m[a.attachment_id] = a;
    return m;
  }, {});

  const bill = bills.find((b) => b.bill_line_id == record.bill_line_id) || {};
  const tenancy = tenancies.find((t) => t.tenancy_id == bill.tenancy_id) || {};
  const tenant = tenants.find((t) => t.tenant_id == (record.tenant_id || tenancy.tenant_id)) || {};
  const unit = units.find((u) => u.unit_id == tenancy.unit_id) || {};
  const attachment = attachmentLookup[record.attachment_id] || {};

  return {
    id: record.payment_id,
    date: formatDateIso_(record.payment_date),
    amount: Number(record.amount) || 0,
    mode: record.mode || '',
    reference: record.reference || '',
    notes: record.notes || '',
    tenantKey: tenancy.grn_number || tenant.full_name || '',
    tenantName: tenant.full_name || '',
    wing: unit.wing || '',
    attachmentName: attachment.file_name || '',
    attachmentUrl: attachment.file_url || '',
    attachmentId: attachment.attachment_id || '',
    monthKey: bill.month_key || '',
    monthLabel: formatMonthLabelForDisplay_(bill.month_key || ''),
    billTotal: Number(bill.total_amount) || 0,
    rentAmount: Number(bill.rent_amount) || 0,
    electricityAmount: Number(bill.electricity_amount) || 0,
    motorShare: Number(bill.motor_share_amount) || 0,
    sweepAmount: Number(bill.sweep_amount) || 0,
    prevReading: '',
    newReading: '',
    payableDate: bill.payable_date || '',
    createdAt: formatDateTime_(record.created_at),
  };
}

function handleFetchPayments_() {
  const payments = readTable_(PAYMENTS_SHEET, PAYMENT_HEADERS);
  return payments.map((p) => mapPaymentRow_(p)).sort((a, b) => {
    const aTime = a.createdAt ? new Date(a.createdAt).getTime() : (a.date ? new Date(a.date).getTime() : 0);
    const bTime = b.createdAt ? new Date(b.createdAt).getTime() : (b.date ? new Date(b.date).getTime() : 0);
    return bTime - aTime;
  });
}

/********* HTTP ENTRY POINTS *********/
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action || '').toLowerCase();
    if (action === 'wings') {
      const ss = getMasterSpreadsheet_();
      let sheet = ss.getSheetByName(WINGS_SHEET);
      if (!sheet) return jsonResponse({ wings: [] });
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return jsonResponse({ wings: [] });
      const values = sheet.getRange(1, 1, lastRow, 1).getValues();
      const wings = values
        .map((r) => (r[0] || '').toString().trim())
        .filter((v, idx) => v && !(idx === 0 && v.toLowerCase() === 'wing'));
      return jsonResponse({ wings });
    }

    if (action === 'clauses') {
      const unified = readUnifiedClauses_();
      const sections = unified || readLegacyClauses_();
      return jsonResponse({ ok: true, tenant: sections.tenant || [], landlord: sections.landlord || [], penalties: sections.penalties || [], misc: sections.misc || [] });
    }

    if (action === 'tenants') {
      return jsonResponse({ ok: true, tenants: buildTenantDirectory_() });
    }

    if (action === 'units') {
      const units = readTable_(UNITS_SHEET, UNITS_HEADERS);
      return jsonResponse({ ok: true, units });
    }

    if (action === 'landlords') {
      const landlords = readTable_(LANDLORDS_SHEET, LANDLORD_HEADERS);
      return jsonResponse({ ok: true, landlords });
    }

    if (action === 'billsminimal') {
      const { bills } = handleFetchBillsMinimal_(e.parameter && e.parameter.status, e.parameter && e.parameter.monthsBack);
      return jsonResponse({ ok: true, bills });
    }

    if (action === 'generatedbills') {
      const { bills, coverage } = handleFetchGeneratedBills_(e.parameter && e.parameter.status);
      return jsonResponse({ ok: true, bills, coverage });
    }

    if (action === 'billdetails') {
      const result = handleFetchBillDetails_(e.parameter && e.parameter.billLineId);
      return jsonResponse(result);
    }

    if (action === 'getbillingrecord') {
      return handleGetBillingRecord_(e.parameter.month, e.parameter.wing);
    }

    if (action === 'payments') {
      const payments = handleFetchPayments_();
      return jsonResponse({ ok: true, payments });
    }

    if (action === 'attachmentpreview') {
      return handleAttachmentPreview_(e.parameter && e.parameter.attachmentUrl);
    }

    return jsonResponse({ ok: true, message: 'GET OK' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return jsonResponse({ ok: false, error: 'No postData.contents' });
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'saveTenant') return handleSaveTenant_(body.payload);
    if (action === 'updateTenant') return handleUpdateTenant_(body.payload);
    if (action === 'saveUnit') return handleSaveUnit_(body.payload || {});
    if (action === 'deleteUnit') return handleDeleteUnit_(body.payload || {});
    if (action === 'saveLandlord') return handleSaveLandlord_(body.payload || {});
    if (action === 'deleteLandlord') return handleDeleteLandlord_(body.payload || {});
    if (action === 'saveBillingRecord') return handleSaveBillingRecord_(body.payload);
    if (action === 'savePayment') return jsonResponse(handleSavePayment_(body.payload || {}));
    if (action === 'uploadPaymentAttachment') {
      const payload = body.payload || {};
      try {
        const result = savePaymentAttachment_(payload.dataUrl || payload.attachmentDataUrl || '', payload.paymentId || '', payload.attachmentName || '', payload.tenantName || '', payload.monthKey || '');
        return jsonResponse({ ok: true, attachment: result });
      } catch (err) {
        return jsonResponse({ ok: false, error: String(err) });
      }
    }
    if (action === 'deleteAttachment') {
      const attachmentId = body.payload && body.payload.attachmentId;
      const ok = deleteAttachmentById_(attachmentId);
      return jsonResponse({ ok, attachmentId });
    }
    if (action === 'saveClauses') {
      writeUnifiedClauses_(body.payload || {});
      return jsonResponse({ ok: true, message: 'Clauses saved to Google Sheets' });
    }
    if (action === 'getRentRevisions') {
      const tenancyId = body.payload && (body.payload.tenancyId || body.payload.tenancy_id);
      const revisions = tenancyId ? listTenancyRentRevisions_(tenancyId) : [];
      return jsonResponse({ ok: true, revisions });
    }
    if (action === 'saveRentRevision') {
      try {
        const record = upsertTenancyRentRevision_(body.payload || {});
        const revisions = listTenancyRentRevisions_(record.tenancy_id);
        return jsonResponse({ ok: true, revision: record, revisions });
      } catch (err) {
        return jsonResponse({ ok: false, error: String(err) });
      }
    }
    if (action === 'deleteRentRevision') {
      const revisionId = body.payload && (body.payload.revisionId || body.payload.revision_id);
      const ok = deleteTenancyRentRevision_(revisionId);
      return jsonResponse({ ok, revisionId });
    }
    if (action === 'addWing') {
      const wing = (body.payload && body.payload.wing || '').toString().trim();
      if (!wing) return jsonResponse({ ok: false, error: 'Wing missing' });
      upsertUnique_(WINGS_SHEET, ['Wing'], ['Wing'], { Wing: wing });
      return jsonResponse({ ok: true, message: 'Wing added' });
    }
    if (action === 'removeWing') return handleRemoveWing_(body.payload || {});
    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}
