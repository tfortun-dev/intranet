const TRIPS_SHEET = 'Trajets';
const VEHICLE_SPREADSHEET_ID = '1RgFrPm-ojzIM9Cqgco5fD8jUA1RMZyIqdtsxAOjh1HE';
const VEHICLE_INFO_SHEET = 'INFOS A LIRE';

const TRIP_HEADERS = [
  'id','name','phone','email','role','date','time',
  'startLabel','startLon','startLat','endLabel','endLon','endLat',
  'vehicleType','companySite','registration','routeJson',
  'createdAt','status','editToken',
  'tripType','returnDate','returnTime','daysOfWeek','privacyMode','privacyRadiusKm',
  'personalVehicleReason','companyVehicleSheet','companyVehicleModel'
];

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Covoiturage interne - Renaud Distribution')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTripsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Le script doit etre lie au Google Sheet Covoiturage interne RD.');
  let sheet = ss.getSheetByName(TRIPS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TRIPS_SHEET);
    sheet.getRange(1, 1, 1, TRIP_HEADERS.length).setValues([TRIP_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  ensureTripHeaders_(sheet);
  return sheet;
}

function ensureTripHeaders_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = TRIP_HEADERS.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function headerMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return { headers, map };
}

function listTrips() {
  const sheet = getTripsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const hm = headerMap_(sheet);
  const rows = sheet.getRange(2, 1, lastRow - 1, hm.headers.length).getValues();
  const today = new Date();
  today.setHours(0,0,0,0);
  const cutoff = new Date(today.getTime() - 7 * 86400000);

  return rows
    .map(row => rowToTrip_(row, hm.map))
    .filter(t => t && t.status === 'active')
    .filter(t => t.tripType === 'regular' || !t.date || new Date(t.date + 'T23:59:59') >= cutoff)
    .sort((a,b) => {
      if (a.tripType !== b.tripType) return a.tripType === 'regular' ? -1 : 1;
      return String(a.date || '').localeCompare(String(b.date || '')) || String(a.time || '').localeCompare(String(b.time || ''));
    });
}

function createTrip(input) {
  input = input || {};
  const trip = sanitizeTrip_(input.trip || input);
  validateTrip_(trip);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let vehicleReservation = null;
    if (trip.tripType === 'punctual' && trip.role === 'driver' && trip.vehicleType === 'company') {
      vehicleReservation = reserveCompanyVehicle_(trip);
      trip.registration = vehicleReservation.registration;
      trip.companySite = 'Le Pallet';
      trip.companyVehicleSheet = vehicleReservation.sheetName;
      trip.companyVehicleModel = vehicleReservation.model;
    }

    const sheet = getTripsSheet_();
    const hm = headerMap_(sheet);
    const id = Utilities.getUuid();
    const managementToken = randomToken_();
    const createdAt = new Date().toISOString();
    const record = tripToRecord_(trip, id, createdAt, sha256Hex_(managementToken));
    const row = hm.headers.map(h => Object.prototype.hasOwnProperty.call(record, h) ? record[h] : '');
    sheet.appendRow(row);

    return {
      trip: Object.assign({}, trip, { id, createdAt, status: 'active', source: 'shared' }),
      managementToken,
      vehicleReservation
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteTrip(id, managementToken) {
  id = clean_(id, 100);
  managementToken = clean_(managementToken, 300);
  if (!id || !managementToken) throw new Error('Identifiant ou jeton de suppression manquant.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getTripsSheet_();
    const hm = headerMap_(sheet);
    const idCol = hm.map.id;
    const tokenCol = hm.map.editToken;
    const statusCol = hm.map.status;
    if ([idCol, tokenCol, statusCol].some(v => v == null)) throw new Error('Structure du tableau Trajets incomplete.');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Trajet introuvable.');
    const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
    const expected = sha256Hex_(managementToken);
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== id) continue;
      const row = i + 2;
      const stored = String(sheet.getRange(row, tokenCol + 1).getValue() || '');
      if (!constantTimeEqual_(stored, expected)) throw new Error('Ce trajet ne peut etre retire que depuis le navigateur qui l’a publie.');
      sheet.getRange(row, statusCol + 1).setValue('deleted');
      return { success: true };
    }
    throw new Error('Trajet introuvable.');
  } finally {
    lock.releaseLock();
  }
}

function getCompanyVehicleAvailability(request) {
  request = request || {};
  const start = parseRequestedDateTime_(request.startDate, request.startTime);
  const end = parseRequestedDateTime_(request.returnDate, request.returnTime);
  if (!start || !end || end <= start) throw new Error('Renseignez un depart et un retour valides.');

  const ss = SpreadsheetApp.openById(VEHICLE_SPREADSHEET_ID);
  const sheets = ss.getSheets().filter(s => {
    const name = s.getName();
    return name !== VEHICLE_INFO_SHEET && !s.isSheetHidden() && /\bDISPO\b/i.test(name) && !/NON\s*DISPO/i.test(name);
  });

  return sheets.map(sheet => inspectVehicleSheet_(sheet, start, end));
}

function inspectVehicleSheet_(sheet, requestedStart, requestedEnd) {
  const name = sheet.getName();
  const title = String(sheet.getRange('A1').getDisplayValue() || name);
  const registration = (title.match(/\b[A-Z]{2}-\d{3}-[A-Z]{2}\b/) || name.match(/\b[A-Z]{2}-\d{3}-[A-Z]{2}\b/) || [''])[0];
  const model = title.replace(registration, '').replace(/^\s*[\/-]+\s*/, '').trim();
  const lastRow = Math.max(4, sheet.getLastRow());
  const values = lastRow >= 5 ? sheet.getRange(5, 1, lastRow - 4, 9).getValues() : [];
  let conflict = null;
  let stopFrom = null;

  values.forEach(row => {
    const borrower = clean_(row[0], 150);
    if (!borrower && !row[2] && !row[5]) return;
    const start = combineSheetDateTime_(row[2], row[3]);
    let end = combineSheetDateTime_(row[5], row[6]);
    if (/STOP\s*RESA/i.test(borrower) && start) {
      stopFrom = stopFrom && stopFrom < start ? stopFrom : start;
      return;
    }
    if (!start && end) {
      const assumedStart = new Date(end.getTime() - 24 * 3600000);
      if (overlaps_(requestedStart, requestedEnd, assumedStart, end)) conflict = conflict || { borrower, start: assumedStart, end };
      return;
    }
    if (start && !end) end = new Date(start.getTime() + 24 * 3600000);
    if (start && end && overlaps_(requestedStart, requestedEnd, start, end)) {
      conflict = conflict || { borrower, start, end };
    }
  });

  const stopped = stopFrom && requestedEnd > stopFrom;
  const available = !conflict && !stopped;
  return {
    sheetName: name,
    registration,
    model,
    available,
    conflict: conflict ? {
      borrower: conflict.borrower,
      start: toLocalIso_(conflict.start),
      end: toLocalIso_(conflict.end)
    } : null,
    stopFrom: stopFrom ? toLocalIso_(stopFrom) : null
  };
}

function reserveCompanyVehicle_(trip) {
  const start = parseRequestedDateTime_(trip.date, trip.time);
  const end = parseRequestedDateTime_(trip.returnDate, trip.returnTime);
  if (!start || !end || end <= start) throw new Error('Le retour du vehicule doit etre renseigne.');
  if (!trip.companyVehicleSheet) throw new Error('Selectionnez un vehicule de societe.');

  const ss = SpreadsheetApp.openById(VEHICLE_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(trip.companyVehicleSheet);
  if (!sheet || sheet.isSheetHidden() || /NON\s*DISPO/i.test(sheet.getName())) throw new Error('Ce vehicule n’est plus disponible.');

  const availability = inspectVehicleSheet_(sheet, start, end);
  if (!availability.available) throw new Error('Ce vehicule vient d’etre reserve sur ce creneau. Actualisez les disponibilites.');

  const lastRow = Math.max(4, sheet.getLastRow());
  const values = lastRow >= 5 ? sheet.getRange(5, 1, lastRow - 4, 9).getValues() : [];
  let insertAt = lastRow + 1;
  for (let i = 0; i < values.length; i++) {
    const rowStart = combineSheetDateTime_(values[i][2], values[i][3]);
    if (rowStart && rowStart > start) { insertAt = i + 5; break; }
  }
  if (insertAt <= lastRow) {
    sheet.insertRowsBefore(insertAt, 1);
    const source = Math.max(5, insertAt - 1);
    sheet.getRange(source, 1, 1, 9).copyFormatToRange(sheet, 1, 9, insertAt, insertAt);
  }

  const row = [
    trip.name,
    trip.end.label,
    start,
    start,
    '',
    end,
    end,
    '',
    'Reservation creee depuis l’outil de covoiturage interne'
  ];
  sheet.getRange(insertAt, 1, 1, 9).setValues([row]);
  sheet.getRange(insertAt, 3).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(insertAt, 4).setNumberFormat('hh"h"mm');
  sheet.getRange(insertAt, 6).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(insertAt, 7).setNumberFormat('hh"h"mm');

  return {
    sheetName: sheet.getName(),
    registration: availability.registration,
    model: availability.model,
    reserved: true
  };
}

function tripToRecord_(trip, id, createdAt, tokenHash) {
  return {
    id,
    name: trip.name,
    phone: trip.phone,
    email: trip.email,
    role: trip.role,
    date: trip.date || '',
    time: trip.time || '',
    startLabel: trip.start.label,
    startLon: trip.start.coordinates[0],
    startLat: trip.start.coordinates[1],
    endLabel: trip.end.label,
    endLon: trip.end.coordinates[0],
    endLat: trip.end.coordinates[1],
    vehicleType: trip.vehicleType,
    companySite: trip.companySite,
    registration: trip.registration,
    routeJson: JSON.stringify(trip.route || {}),
    createdAt,
    status: 'active',
    editToken: tokenHash,
    tripType: trip.tripType,
    returnDate: trip.returnDate || '',
    returnTime: trip.returnTime || '',
    daysOfWeek: (trip.daysOfWeek || []).join(','),
    privacyMode: trip.privacyMode || 'exact',
    privacyRadiusKm: Number(trip.privacyRadiusKm || 0),
    personalVehicleReason: trip.personalVehicleReason || '',
    companyVehicleSheet: trip.companyVehicleSheet || '',
    companyVehicleModel: trip.companyVehicleModel || ''
  };
}

function rowToTrip_(row, map) {
  if (!row || !value_(row, map, 'id')) return null;
  let route = {};
  try { route = JSON.parse(String(value_(row, map, 'routeJson') || '{}')); } catch (e) { route = {}; }
  const tripType = clean_(value_(row, map, 'tripType'), 20) || 'punctual';
  const daysRaw = clean_(value_(row, map, 'daysOfWeek'), 100);
  return {
    id: String(value_(row, map, 'id') || ''),
    name: String(value_(row, map, 'name') || ''),
    phone: String(value_(row, map, 'phone') || ''),
    email: String(value_(row, map, 'email') || ''),
    role: String(value_(row, map, 'role') || ''),
    date: normalizeDate_(value_(row, map, 'date')),
    time: normalizeTime_(value_(row, map, 'time')),
    returnDate: normalizeDate_(value_(row, map, 'returnDate')),
    returnTime: normalizeTime_(value_(row, map, 'returnTime')),
    start: {
      label: String(value_(row, map, 'startLabel') || ''),
      coordinates: [Number(value_(row, map, 'startLon')), Number(value_(row, map, 'startLat'))]
    },
    end: {
      label: String(value_(row, map, 'endLabel') || ''),
      coordinates: [Number(value_(row, map, 'endLon')), Number(value_(row, map, 'endLat'))]
    },
    vehicleType: String(value_(row, map, 'vehicleType') || 'none'),
    companySite: String(value_(row, map, 'companySite') || ''),
    registration: String(value_(row, map, 'registration') || ''),
    companyVehicleSheet: String(value_(row, map, 'companyVehicleSheet') || ''),
    companyVehicleModel: String(value_(row, map, 'companyVehicleModel') || ''),
    personalVehicleReason: String(value_(row, map, 'personalVehicleReason') || ''),
    route,
    createdAt: normalizeIso_(value_(row, map, 'createdAt')),
    status: String(value_(row, map, 'status') || 'active'),
    tripType,
    daysOfWeek: daysRaw ? daysRaw.split(',').filter(Boolean) : [],
    privacyMode: String(value_(row, map, 'privacyMode') || 'exact'),
    privacyRadiusKm: Number(value_(row, map, 'privacyRadiusKm') || 0),
    recurrence: tripType === 'regular' ? 'weekly' : 'once',
    tolerance: 90,
    maxDetour: 20,
    seats: String(value_(row, map, 'role') || '') === 'driver' ? 1 : 0,
    source: 'shared'
  };
}

function sanitizeTrip_(input) {
  input = input || {};
  const start = input.start || {};
  const end = input.end || {};
  return {
    name: clean_(input.name, 100),
    phone: clean_(input.phone, 30),
    email: clean_(input.email, 160),
    role: clean_(input.role, 20),
    tripType: clean_(input.tripType, 20) || 'punctual',
    date: clean_(input.date, 10),
    time: clean_(input.time, 5),
    returnDate: clean_(input.returnDate, 10),
    returnTime: clean_(input.returnTime, 5),
    daysOfWeek: Array.isArray(input.daysOfWeek) ? input.daysOfWeek.map(v => clean_(v, 3)).filter(Boolean) : [],
    privacyMode: clean_(input.privacyMode, 20) || 'exact',
    privacyRadiusKm: Number(input.privacyRadiusKm || 0),
    start: {
      label: clean_(start.label, 250),
      coordinates: [Number(start.coordinates && start.coordinates[0]), Number(start.coordinates && start.coordinates[1])]
    },
    end: {
      label: clean_(end.label, 250),
      coordinates: [Number(end.coordinates && end.coordinates[0]), Number(end.coordinates && end.coordinates[1])]
    },
    vehicleType: clean_(input.vehicleType, 30) || 'none',
    companySite: clean_(input.companySite, 100),
    registration: clean_(input.registration, 30),
    companyVehicleSheet: clean_(input.companyVehicleSheet, 220),
    companyVehicleModel: clean_(input.companyVehicleModel, 160),
    personalVehicleReason: clean_(input.personalVehicleReason, 300),
    route: sanitizeRoute_(input.route)
  };
}

function validateTrip_(trip) {
  if (!trip.name) throw new Error('Nom obligatoire.');
  if (!trip.phone && !trip.email) throw new Error('Un telephone ou un e-mail est obligatoire.');
  if (!['driver','passenger'].includes(trip.role)) throw new Error('Role invalide.');
  if (!['punctual','regular'].includes(trip.tripType)) throw new Error('Type de trajet invalide.');
  if (!trip.start.label || !trip.end.label || !trip.start.coordinates.every(Number.isFinite) || !trip.end.coordinates.every(Number.isFinite)) {
    throw new Error('Depart et arrivee invalides.');
  }
  if (trip.tripType === 'punctual') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(trip.date) || !/^\d{2}:\d{2}$/.test(trip.time)) throw new Error('Date ou heure invalide.');
    if (trip.role === 'driver' && trip.vehicleType === 'company') {
      if (!trip.returnDate || !trip.returnTime || !trip.companyVehicleSheet) throw new Error('Le retour et le vehicule de societe sont obligatoires.');
    }
    if (trip.role === 'driver' && trip.vehicleType === 'personal' && !trip.personalVehicleReason) {
      throw new Error('Expliquez pourquoi vous utilisez votre voiture personnelle.');
    }
  } else {
    if (!trip.daysOfWeek.length) throw new Error('Selectionnez au moins un jour habituel.');
    if (!/^\d{2}:\d{2}$/.test(trip.time)) throw new Error('Heure habituelle invalide.');
    if (!['exact','approx1km'].includes(trip.privacyMode)) throw new Error('Mode de confidentialite invalide.');
  }
}

function sanitizeRoute_(route) {
  route = route || {};
  const out = {
    distance: Number(route.distance),
    duration: Number(route.duration),
    provider: clean_(route.provider, 50),
    estimated: Boolean(route.estimated),
    geometry: route.geometry || null
  };
  if (!Number.isFinite(out.distance)) out.distance = 0;
  if (!Number.isFinite(out.duration)) out.duration = 0;
  return out;
}

function parseRequestedDateTime_(date, time) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || '')) || !/^\d{2}:\d{2}$/.test(String(time || ''))) return null;
  const d = new Date(String(date) + 'T' + String(time) + ':00');
  return isNaN(d.getTime()) ? null : d;
}

function combineSheetDateTime_(dateValue, timeValue) {
  let y, m, d;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    y = dateValue.getFullYear(); m = dateValue.getMonth(); d = dateValue.getDate();
  } else {
    const text = String(dateValue || '').trim();
    let match = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
    if (match) { d = Number(match[1]); m = Number(match[2]) - 1; y = Number(match[3]); }
    else {
      match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
      if (!match) return null;
      y = Number(match[1]); m = Number(match[2]) - 1; d = Number(match[3]);
    }
  }

  let hh = 0, mm = 0;
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    hh = timeValue.getHours(); mm = timeValue.getMinutes();
  } else {
    const t = String(timeValue || '').trim().toLowerCase().replace(/\s/g, '');
    const match = /^(\d{1,2})(?:h|:)?(\d{0,2})/.exec(t);
    if (match) { hh = Number(match[1]); mm = Number(match[2] || 0); }
  }
  const out = new Date(y, m, d, hh, mm, 0, 0);
  return isNaN(out.getTime()) ? null : out;
}

function overlaps_(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && aEnd > bStart;
}

function toLocalIso_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', "yyyy-MM-dd'T'HH:mm:ss");
}

function value_(row, map, key) {
  return map[key] == null ? '' : row[map[key]];
}

function clean_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 500);
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM-dd');
  return String(value || '').slice(0,10);
}

function normalizeTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Europe/Paris', 'HH:mm');
  const s = String(value || '').trim();
  const m = /^(\d{1,2})(?:h|:)(\d{0,2})/i.exec(s);
  if (m) return String(Number(m[1])).padStart(2,'0') + ':' + String(Number(m[2] || 0)).padStart(2,'0');
  return s.slice(0,5);
}

function normalizeIso_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
  return String(value || '');
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function constantTimeEqual_(a,b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i=0;i<a.length;i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}