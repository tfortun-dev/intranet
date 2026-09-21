const SHEET_NAME = 'Trajets';
const HEADERS = [
  'id','name','phone','email','role','date','time',
  'startLabel','startLon','startLat','endLabel','endLon','endLat',
  'vehicleType','companySite','registration','routeJson',
  'createdAt','status','editToken'
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
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (current.join('|') !== HEADERS.join('|')) {
    throw new Error('Les colonnes de l\'onglet Trajets ne correspondent pas au format attendu.');
  }
  return sheet;
}

function listTrips() {
  const sheet = getTripsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const now = new Date();

  return rows
    .map(rowToTrip_)
    .filter(trip => trip && trip.status === 'active')
    .filter(trip => {
      if (!trip.date) return true;
      const d = new Date(trip.date + 'T23:59:59');
      return isNaN(d.getTime()) || d >= new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7);
    })
    .sort((a, b) => String(a.date || '').localeCompare(String(b.date || '')) || String(a.time || '').localeCompare(String(b.time || '')));
}

function createTrip(input) {
  const trip = sanitizeTrip_(input);
  validateTrip_(trip);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getTripsSheet_();
    const id = Utilities.getUuid();
    const managementToken = randomToken_();
    const editToken = sha256Hex_(managementToken);
    const createdAt = new Date().toISOString();

    const row = [
      id,
      trip.name,
      trip.phone,
      trip.email,
      trip.role,
      trip.date,
      trip.time,
      trip.start.label,
      trip.start.coordinates[0],
      trip.start.coordinates[1],
      trip.end.label,
      trip.end.coordinates[0],
      trip.end.coordinates[1],
      trip.vehicleType,
      trip.companySite,
      trip.registration,
      JSON.stringify(trip.route || {}),
      createdAt,
      'active',
      editToken
    ];

    sheet.appendRow(row);

    return {
      trip: Object.assign({}, trip, {
        id,
        createdAt,
        status: 'active',
        source: 'shared'
      }),
      managementToken
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteTrip(id, managementToken) {
  id = String(id || '').trim();
  managementToken = String(managementToken || '').trim();
  if (!id || !managementToken) throw new Error('Identifiant ou jeton de suppression manquant.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getTripsSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Trajet introuvable.');

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const expectedHash = sha256Hex_(managementToken);

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== id) continue;

      const rowNumber = i + 2;
      const storedHash = String(sheet.getRange(rowNumber, 20).getValue() || '');
      if (!constantTimeEqual_(storedHash, expectedHash)) {
        throw new Error('Ce trajet ne peut pas etre supprime depuis ce navigateur.');
      }

      sheet.getRange(rowNumber, 19).setValue('deleted');
      return { success: true };
    }

    throw new Error('Trajet introuvable.');
  } finally {
    lock.releaseLock();
  }
}

function rowToTrip_(row) {
  if (!row || !row[0]) return null;

  let route = {};
  try {
    route = row[16] ? JSON.parse(String(row[16])) : {};
  } catch (e) {
    route = {};
  }

  return {
    id: String(row[0] || ''),
    name: String(row[1] || ''),
    phone: String(row[2] || ''),
    email: String(row[3] || ''),
    role: String(row[4] || ''),
    date: normalizeDate_(row[5]),
    time: normalizeTime_(row[6]),
    start: {
      label: String(row[7] || ''),
      coordinates: [Number(row[8]), Number(row[9])]
    },
    end: {
      label: String(row[10] || ''),
      coordinates: [Number(row[11]), Number(row[12])]
    },
    vehicleType: String(row[13] || 'none'),
    companySite: String(row[14] || ''),
    registration: String(row[15] || ''),
    route,
    createdAt: normalizeIso_(row[17]),
    status: String(row[18] || 'active'),
    recurrence: 'once',
    tolerance: 90,
    maxDetour: 20,
    seats: String(row[4] || '') === 'driver' ? 1 : 0,
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
    date: clean_(input.date, 10),
    time: clean_(input.time, 5),
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
    route: sanitizeRoute_(input.route)
  };
}

function validateTrip_(trip) {
  if (!trip.name) throw new Error('Nom obligatoire.');
  if (!trip.phone && !trip.email) throw new Error('Un telephone ou un e-mail est obligatoire.');
  if (!['driver', 'passenger'].includes(trip.role)) throw new Error('Role invalide.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(trip.date)) throw new Error('Date invalide.');
  if (!/^\d{2}:\d{2}$/.test(trip.time)) throw new Error('Heure invalide.');
  if (!trip.start.label || !trip.end.label) throw new Error('Depart et arrivee obligatoires.');
  if (!trip.start.coordinates.every(Number.isFinite) || !trip.end.coordinates.every(Number.isFinite)) {
    throw new Error('Coordonnees de trajet invalides.');
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

function clean_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 500);
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10);
}

function normalizeTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Europe/Paris', 'HH:mm');
  }
  return String(value || '').slice(0, 5);
}

function normalizeIso_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
  return String(value || '');
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function constantTimeEqual_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
