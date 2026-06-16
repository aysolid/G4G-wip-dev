/**
 * G4G Research Operations Web App
 * Main Server-Side Code
 *
 * This file contains the core web app functionality including:
 * - Web app entry points (doGet)
 * - Page routing
 * - Session management
 */

// ============================================
// CONFIGURATION
// ============================================

const CONFIG = {
  APP_NAME: 'G4G Research Operations',
  VERSION: '1.0.0',
  SITES: ['UGA', 'Missouri'],
  PERIODS: ['Spring', 'Summer', 'Fall'],
  ROLES: ['admin', 'facilitator', 'viewer'],
  PARTICIPANT_STATUSES: ['active', 'withdrawn', 'completed'],
  CHECKLIST_STATUSES: ['not_started', 'completed', 'missing'],

  // The 18 non-negotiable protocol items
  INSTRUMENTS: [
    { number: 1, name: 'Consent Form', category: 'enrollment' },
    { number: 2, name: 'Assent Form / Pre-test', category: 'enrollment' },
    { number: 3, name: 'Lesson 1 Journal', category: 'lesson' },
    { number: 4, name: 'Lesson 2 Journal', category: 'lesson' },
    { number: 5, name: 'Lesson 3 Journal', category: 'lesson' },
    { number: 6, name: 'Lesson 3 - Design a Game Worksheet', category: 'worksheet' },
    { number: 7, name: 'Lesson 4 Journal', category: 'lesson' },
    { number: 8, name: 'Lesson 4 - Paper Prototyping Worksheet', category: 'worksheet' },
    { number: 9, name: 'Lesson 5 Journal', category: 'lesson' },
    { number: 10, name: 'Lesson 5 - Debugging Worksheet', category: 'worksheet' },
    { number: 11, name: 'Lesson 6 Journal', category: 'lesson' },
    { number: 12, name: 'Lesson 6 - Game Refinement Worksheet', category: 'worksheet' },
    { number: 13, name: 'Lesson 7 Journal', category: 'lesson' },
    { number: 14, name: 'Lesson 7 - Playtesting Feedback Guide', category: 'worksheet' },
    { number: 15, name: 'Lesson 8 Journal', category: 'lesson' },
    { number: 16, name: 'Post-Test', category: 'post' },
    { number: 17, name: 'Participant Feedback Survey', category: 'post' },
    { number: 18, name: 'Parent Satisfaction Survey', category: 'post' }
  ]
};

const RUNTIME_CACHE = {
  configMap: null,
  globalProtocolItems: null,
  rolloutProtocolItems: {},
  sheetSnapshots: {}
};

function getSheetSnapshot(sheetName, options) {
  const key = sheetName + '::' + (options && options.ensureFn ? options.ensureFn.name : 'raw');
  if (RUNTIME_CACHE.sheetSnapshots[key]) return RUNTIME_CACHE.sheetSnapshots[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) {
    const empty = { sheet: sheet, headers: [], data: [] };
    RUNTIME_CACHE.sheetSnapshots[key] = empty;
    return empty;
  }
  const headers = (options && options.ensureFn)
    ? options.ensureFn(sheet)
    : sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
  const rowCount = sheet.getLastRow();
  const colCount = headers.length || sheet.getLastColumn();
  const data = rowCount > 0 ? sheet.getRange(1, 1, rowCount, colCount).getValues() : [];
  const snapshot = { sheet: sheet, headers: headers, data: data };
  RUNTIME_CACHE.sheetSnapshots[key] = snapshot;
  return snapshot;
}

function invalidateSheetSnapshot(sheetName) {
  Object.keys(RUNTIME_CACHE.sheetSnapshots).forEach(key => {
    if (key === sheetName || key.indexOf(sheetName + '::') === 0) {
      delete RUNTIME_CACHE.sheetSnapshots[key];
    }
  });
}

function invalidateConfigRuntimeCache() {
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
  invalidateSheetSnapshot('Config');
}

function getGlobalProtocolItems() {
  if (RUNTIME_CACHE.globalProtocolItems) return RUNTIME_CACHE.globalProtocolItems;
  const configMap = getConfigMap();
  const raw = configMap.PROTOCOL_ITEMS;
  if (raw) {
    try {
      const parsed = JSON.parse(raw || '[]');
      if (Array.isArray(parsed) && parsed.length > 0) {
        RUNTIME_CACHE.globalProtocolItems = parsed;
        return parsed;
      }
    } catch (e) {}
  }
  RUNTIME_CACHE.globalProtocolItems = CONFIG.INSTRUMENTS;
  return CONFIG.INSTRUMENTS;
}

function getRolloutProtocolItemsInternal(rolloutId) {
  if (RUNTIME_CACHE.rolloutProtocolItems[String(rolloutId)]) {
    return RUNTIME_CACHE.rolloutProtocolItems[String(rolloutId)];
  }
  const globalItems = getGlobalProtocolItems();
  let overrides = {};
  const configMap = getConfigMap();
  if (configMap.ROLLOUT_PROTOCOL_OVERRIDES) {
    try { overrides = JSON.parse(configMap.ROLLOUT_PROTOCOL_OVERRIDES || '{}') || {}; } catch (e) {}
  }
  const selectedIds = overrides[rolloutId];
  if (!Array.isArray(selectedIds) || selectedIds.length === 0) {
    RUNTIME_CACHE.rolloutProtocolItems[String(rolloutId)] = globalItems;
    return globalItems;
  }
  const selectedMap = {};
  selectedIds.forEach(id => selectedMap[String(id)] = true);
  const filtered = globalItems.filter(item => selectedMap[String(item.number)]);
  RUNTIME_CACHE.rolloutProtocolItems[String(rolloutId)] = filtered;
  return filtered;
}

function getConfigMap() {
  if (RUNTIME_CACHE.configMap) return RUNTIME_CACHE.configMap;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const result = {};
  if (!configSheet || configSheet.getLastRow() < 2) {
    RUNTIME_CACHE.configMap = result;
    return result;
  }
  const data = configSheet.getRange(1, 1, configSheet.getLastRow(), configSheet.getLastColumn()).getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  for (let i = 1; i < data.length; i++) result[data[i][keyCol]] = data[i][valueCol];
  RUNTIME_CACHE.configMap = result;
  return result;
}

function getProtocolItemsForFilter(rolloutFilter) {
  if (rolloutFilter && rolloutFilter !== 'All') return getRolloutProtocolItemsInternal(rolloutFilter);
  return getGlobalProtocolItems();
}

// ============================================
// WEB APP ENTRY POINTS
// ============================================

/**
 * Main entry point for the web app
 * Handles GET requests and routes to appropriate pages
 */
function doGet(e) {
  const page = e.parameter.page || 'index';
  const sessionToken = e.parameter.token || '';

  // Validate session for protected pages
  const protectedPages = ['users', 'cohorts', 'participants', 'participant-detail', 'reports'];

  if (protectedPages.includes(page)) {
    const session = validateSession(sessionToken);
    if (!session) {
      return HtmlService.createTemplateFromFile('login')
        .evaluate()
        .setTitle(CONFIG.APP_NAME + ' - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  try {
    const template = HtmlService.createTemplateFromFile(page);
    return template.evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('Error loading page: ' + error.message);
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle(CONFIG.APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Include HTML files (for modular HTML)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get the web app URL
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Get configuration for client-side use
 */
function getConfig() {
  const protocols = getGlobalProtocolItems();
  const recordingTypes = getSessionRecordingTypes();
  return {
    appName: CONFIG.APP_NAME,
    version: CONFIG.VERSION,
    sites: CONFIG.SITES,
    periods: CONFIG.PERIODS,
    roles: CONFIG.ROLES,
    instruments: protocols,
    participantStatuses: CONFIG.PARTICIPANT_STATUSES,
    checklistStatuses: CONFIG.CHECKLIST_STATUSES,
    sessionRecordingTypes: recordingTypes,
    enrollmentFields: getEnrollmentFieldsInternal()
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

// ============================================
// MEDIA & FIELD RECORDS
// ============================================

function ensureMediaRecordSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheetIfNotExists(ss, 'ParticipantMedia', [
    'mediaId', 'rolloutId', 'participantId', 'site', 'mediaType', 'title', 'driveUrl',
    'thumbnailUrl', 'playtesterParticipantId', 'playtesterName', 'notes', 'createdAt',
    'createdBy', 'updatedAt', 'updatedBy', 'status'
  ]);
  createSheetIfNotExists(ss, 'CohortMedia', [
    'mediaId', 'rolloutId', 'site', 'mediaCategory', 'title', 'driveUrl', 'thumbnailUrl',
    'description', 'capturedDate', 'createdAt', 'createdBy', 'updatedAt', 'updatedBy', 'status'
  ]);
}

function buildRowFromObject(headers, data) {
  return headers.map(header => data[header] !== undefined ? data[header] : '');
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, idx) => obj[header] = row[idx]);
  return obj;
}

function getParticipantsForMediaRollout_(token, rolloutId) {
  const result = getAllParticipants(token, { rolloutId: rolloutId });
  if (!result.success) return result;
  const participants = (result.participants || []).filter(participant => String(participant.rolloutId) === String(rolloutId));
  return { success: true, participants: participants };
}

function getParticipantMediaByRollout(token, rolloutId, mediaType) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  if (!rolloutId || rolloutId === 'All') return { success: false, message: 'Select a specific cohort first' };
  ensureMediaRecordSheets();
  const participantsResult = getParticipantsForMediaRollout_(token, rolloutId);
  if (!participantsResult.success) return participantsResult;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipantMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const media = [];
  for (let i = 1; i < data.length; i++) {
    const item = rowToObject(headers, data[i]);
    if (String(item.rolloutId) !== String(rolloutId)) continue;
    if (String(item.mediaType) !== String(mediaType)) continue;
    if (String(item.status || 'active').toLowerCase() === 'deleted') continue;
    media.push(item);
  }
  return { success: true, participants: participantsResult.participants, media: media };
}

function saveParticipantMedia(token, media) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  ensureMediaRecordSheets();
  const payload = media || {};
  if (!payload.rolloutId || !payload.participantId || !payload.mediaType) return { success: false, message: 'rolloutId, participantId, and mediaType are required' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipantMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const now = new Date().toISOString();
  const existingId = String(payload.mediaId || '').trim();
  const mediaIdCol = headers.indexOf('mediaId');
  const rolloutCol = headers.indexOf('rolloutId');
  const participantCol = headers.indexOf('participantId');
  const typeCol = headers.indexOf('mediaType');
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    const byId = existingId && String(data[i][mediaIdCol]) === existingId;
    const byNaturalKey = String(data[i][rolloutCol]) === String(payload.rolloutId) && String(data[i][participantCol]) === String(payload.participantId) && String(data[i][typeCol]) === String(payload.mediaType) && String(data[i][headers.indexOf('status')] || 'active') !== 'deleted';
    if (byId || byNaturalKey) { rowIndex = i + 1; break; }
  }
  const existing = rowIndex > 0 ? rowToObject(headers, data[rowIndex - 1]) : {};
  const record = Object.assign({}, existing, {
    mediaId: existing.mediaId || existingId || generateUUID(),
    rolloutId: payload.rolloutId,
    participantId: payload.participantId,
    site: payload.site || existing.site || '',
    mediaType: payload.mediaType,
    title: payload.title || existing.title || '',
    driveUrl: String(payload.driveUrl || '').trim(),
    thumbnailUrl: String(payload.thumbnailUrl || existing.thumbnailUrl || '').trim(),
    playtesterParticipantId: payload.playtesterParticipantId || '',
    playtesterName: payload.playtesterName || '',
    notes: payload.notes || '',
    createdAt: existing.createdAt || now,
    createdBy: existing.createdBy || currentUser.fullName,
    updatedAt: now,
    updatedBy: currentUser.fullName,
    status: payload.status || 'active'
  });
  const row = buildRowFromObject(headers, record);
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  else appendRowsAsPlainText(sheet, [row]);
  invalidateSheetSnapshot('ParticipantMedia');
  return { success: true, message: 'Participant media saved', media: record };
}

function saveParticipantMediaBatch(token, mediaItems) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const items = Array.isArray(mediaItems) ? mediaItems : [];
  if (!items.length) return { success: false, message: 'No media records were provided' };
  ensureMediaRecordSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipantMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const mediaIdCol = headers.indexOf('mediaId');
  const rolloutCol = headers.indexOf('rolloutId');
  const participantCol = headers.indexOf('participantId');
  const typeCol = headers.indexOf('mediaType');
  const statusCol = headers.indexOf('status');
  const now = new Date().toISOString();
  const existingById = {};
  const existingByNaturalKey = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNumber = i + 1;
    const mediaId = String(row[mediaIdCol] || '').trim();
    const naturalKey = [
      row[rolloutCol],
      row[participantCol],
      row[typeCol]
    ].map(value => String(value || '')).join('::');
    const isDeleted = String(row[statusCol] || 'active').toLowerCase() === 'deleted';
    if (mediaId) existingById[mediaId] = { rowNumber: rowNumber, row: row };
    if (!isDeleted) existingByNaturalKey[naturalKey] = { rowNumber: rowNumber, row: row };
  }

  const rowsToAppend = [];
  let savedCount = 0;
  items.forEach(item => {
    const payload = item || {};
    if (!payload.rolloutId || !payload.participantId || !payload.mediaType) return;
    const existingId = String(payload.mediaId || '').trim();
    const naturalKey = [
      payload.rolloutId,
      payload.participantId,
      payload.mediaType
    ].map(value => String(value || '')).join('::');
    const match = (existingId && existingById[existingId]) || existingByNaturalKey[naturalKey];
    const existing = match ? rowToObject(headers, match.row) : {};
    const record = Object.assign({}, existing, {
      mediaId: existing.mediaId || existingId || generateUUID(),
      rolloutId: payload.rolloutId,
      participantId: payload.participantId,
      site: payload.site || existing.site || '',
      mediaType: payload.mediaType,
      title: payload.title || existing.title || '',
      driveUrl: String(payload.driveUrl || '').trim(),
      thumbnailUrl: String(payload.thumbnailUrl || existing.thumbnailUrl || '').trim(),
      playtesterParticipantId: payload.playtesterParticipantId || '',
      playtesterName: payload.playtesterName || '',
      notes: payload.notes || '',
      createdAt: existing.createdAt || now,
      createdBy: existing.createdBy || currentUser.fullName,
      updatedAt: now,
      updatedBy: currentUser.fullName,
      status: payload.status || 'active'
    });
    const row = buildRowFromObject(headers, record);
    if (match) {
      sheet.getRange(match.rowNumber, 1, 1, headers.length).setValues([row]);
    } else {
      rowsToAppend.push(row);
    }
    savedCount++;
  });

  if (rowsToAppend.length) appendRowsAsPlainText(sheet, rowsToAppend);
  invalidateSheetSnapshot('ParticipantMedia');
  return { success: true, message: `${savedCount} participant media record${savedCount === 1 ? '' : 's'} saved`, savedCount: savedCount };
}

function deleteParticipantMedia(token, mediaId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  ensureMediaRecordSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipantMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const mediaIdCol = headers.indexOf('mediaId');
  const statusCol = headers.indexOf('status');
  const updatedAtCol = headers.indexOf('updatedAt');
  const updatedByCol = headers.indexOf('updatedBy');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][mediaIdCol]) === String(mediaId)) {
      const row = data[i].slice(0, headers.length);
      row[statusCol] = 'deleted';
      row[updatedAtCol] = new Date().toISOString();
      row[updatedByCol] = currentUser.fullName;
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateSheetSnapshot('ParticipantMedia');
      return { success: true, message: 'Participant media removed' };
    }
  }
  return { success: false, message: 'Media record not found' };
}

function getCohortMediaByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  if (!rolloutId || rolloutId === 'All') return { success: false, message: 'Select a specific cohort first' };
  ensureMediaRecordSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CohortMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const media = [];
  for (let i = 1; i < data.length; i++) {
    const item = rowToObject(headers, data[i]);
    if (String(item.rolloutId) !== String(rolloutId)) continue;
    if (String(item.status || 'active').toLowerCase() === 'deleted') continue;
    media.push(item);
  }
  return { success: true, media: media };
}

function saveCohortMedia(token, media) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  ensureMediaRecordSheets();
  const payload = media || {};
  if (!payload.rolloutId || !payload.mediaCategory) return { success: false, message: 'rolloutId and mediaCategory are required' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CohortMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const now = new Date().toISOString();
  const existingId = String(payload.mediaId || '').trim();
  let rowIndex = -1;
  const mediaIdCol = headers.indexOf('mediaId');
  for (let i = 1; i < data.length; i++) {
    if (existingId && String(data[i][mediaIdCol]) === existingId) { rowIndex = i + 1; break; }
  }
  const existing = rowIndex > 0 ? rowToObject(headers, data[rowIndex - 1]) : {};
  const record = Object.assign({}, existing, {
    mediaId: existing.mediaId || existingId || generateUUID(),
    rolloutId: payload.rolloutId,
    site: payload.site || existing.site || '',
    mediaCategory: payload.mediaCategory,
    title: payload.title || '',
    driveUrl: String(payload.driveUrl || '').trim(),
    thumbnailUrl: String(payload.thumbnailUrl || '').trim(),
    description: payload.description || '',
    capturedDate: payload.capturedDate || '',
    createdAt: existing.createdAt || now,
    createdBy: existing.createdBy || currentUser.fullName,
    updatedAt: now,
    updatedBy: currentUser.fullName,
    status: payload.status || 'active'
  });
  const row = buildRowFromObject(headers, record);
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  else appendRowsAsPlainText(sheet, [row]);
  invalidateSheetSnapshot('CohortMedia');
  return { success: true, message: 'Cohort media saved', media: record };
}

function deleteCohortMedia(token, mediaId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  ensureMediaRecordSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CohortMedia');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const mediaIdCol = headers.indexOf('mediaId');
  const statusCol = headers.indexOf('status');
  const updatedAtCol = headers.indexOf('updatedAt');
  const updatedByCol = headers.indexOf('updatedBy');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][mediaIdCol]) === String(mediaId)) {
      const row = data[i].slice(0, headers.length);
      row[statusCol] = 'deleted';
      row[updatedAtCol] = new Date().toISOString();
      row[updatedByCol] = currentUser.fullName;
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateSheetSnapshot('CohortMedia');
      return { success: true, message: 'Cohort media removed' };
    }
  }
  return { success: false, message: 'Media record not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}


function getFieldNotesByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const sessions = (sessionsResult.sessions || []).map(function(s) {
    return Object.assign({}, s, { fieldNotesLink: s.fieldNotesLink || '' });
  });
  return { success: true, sessions: sessions };
}

function saveFieldNotesLink(token, sessionId, link) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionIdIdx = headers.indexOf('sessionId');
  if (sessionIdIdx === -1) return { success: false, message: 'sessionId column missing' };

  let fieldNotesIdx = headers.indexOf('fieldNotesLink');
  if (fieldNotesIdx === -1) {
    headers.push('fieldNotesLink');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fieldNotesIdx = headers.length - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIdIdx]) === String(sessionId)) {
      sheet.getRange(i + 1, fieldNotesIdx + 1).setValue(String(link || '').trim());
      invalidateSheetSnapshot('StudySessions');
      return { success: true, message: 'Field notes link saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}


function getFieldNotesByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const sessions = (sessionsResult.sessions || []).map(function(s) {
    return Object.assign({}, s, { fieldNotesLink: s.fieldNotesLink || '' });
  });
  return { success: true, sessions: sessions };
}

function saveFieldNotesLink(token, sessionId, link) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionIdIdx = headers.indexOf('sessionId');
  if (sessionIdIdx === -1) return { success: false, message: 'sessionId column missing' };

  let fieldNotesIdx = headers.indexOf('fieldNotesLink');
  if (fieldNotesIdx === -1) {
    headers.push('fieldNotesLink');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fieldNotesIdx = headers.length - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIdIdx]) === String(sessionId)) {
      sheet.getRange(i + 1, fieldNotesIdx + 1).setValue(String(link || '').trim());
      invalidateSheetSnapshot('StudySessions');
      return { success: true, message: 'Field notes link saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}


function getFieldNotesByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const sessions = (sessionsResult.sessions || []).map(function(s) {
    return Object.assign({}, s, { fieldNotesLink: s.fieldNotesLink || '' });
  });
  return { success: true, sessions: sessions };
}

function saveFieldNotesLink(token, sessionId, link) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionIdIdx = headers.indexOf('sessionId');
  if (sessionIdIdx === -1) return { success: false, message: 'sessionId column missing' };

  let fieldNotesIdx = headers.indexOf('fieldNotesLink');
  if (fieldNotesIdx === -1) {
    headers.push('fieldNotesLink');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fieldNotesIdx = headers.length - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIdIdx]) === String(sessionId)) {
      sheet.getRange(i + 1, fieldNotesIdx + 1).setValue(String(link || '').trim());
      invalidateSheetSnapshot('StudySessions');
      return { success: true, message: 'Field notes link saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}

/**
 * Get configuration for client-side use
 */
function getConfig() {
  const protocols = getGlobalProtocolItems();
  const recordingTypes = getSessionRecordingTypes();
  return {
    appName: CONFIG.APP_NAME,
    version: CONFIG.VERSION,
    sites: CONFIG.SITES,
    periods: CONFIG.PERIODS,
    roles: CONFIG.ROLES,
    instruments: protocols,
    participantStatuses: CONFIG.PARTICIPANT_STATUSES,
    checklistStatuses: CONFIG.CHECKLIST_STATUSES,
    sessionRecordingTypes: recordingTypes,
    enrollmentFields: getEnrollmentFieldsInternal()
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}


function getFieldNotesByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const sessions = (sessionsResult.sessions || []).map(function(s) {
    return Object.assign({}, s, { fieldNotesLink: s.fieldNotesLink || '' });
  });
  return { success: true, sessions: sessions };
}

function saveFieldNotesLink(token, sessionId, link) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionIdIdx = headers.indexOf('sessionId');
  if (sessionIdIdx === -1) return { success: false, message: 'sessionId column missing' };

  let fieldNotesIdx = headers.indexOf('fieldNotesLink');
  if (fieldNotesIdx === -1) {
    headers.push('fieldNotesLink');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fieldNotesIdx = headers.length - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIdIdx]) === String(sessionId)) {
      sheet.getRange(i + 1, fieldNotesIdx + 1).setValue(String(link || '').trim());
      invalidateSheetSnapshot('StudySessions');
      return { success: true, message: 'Field notes link saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

/**
 * Get configuration for client-side use
 */
function getConfig() {
  const protocols = getGlobalProtocolItems();
  const recordingTypes = getSessionRecordingTypes();
  return {
    appName: CONFIG.APP_NAME,
    version: CONFIG.VERSION,
    sites: CONFIG.SITES,
    periods: CONFIG.PERIODS,
    roles: CONFIG.ROLES,
    instruments: protocols,
    participantStatuses: CONFIG.PARTICIPANT_STATUSES,
    checklistStatuses: CONFIG.CHECKLIST_STATUSES,
    sessionRecordingTypes: recordingTypes,
    enrollmentFields: getEnrollmentFieldsInternal()
  };
}

function getSessionRecordingTypes() {
  const map = getConfigMap();
  const raw = map.SESSION_RECORDING_TYPES;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed.map(v => String(v).trim()).filter(Boolean);
    } catch (e) {}
  }
  return ['GoPro', 'Tascam', 'Meeting Owl'];
}

function saveSessionRecordingTypes(token, types) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const normalized = (types || []).map(v => String(v || '').trim()).filter(Boolean);
  if (!normalized.length) return { success: false, message: 'At least one recording type is required' };
  upsertConfigValue('SESSION_RECORDING_TYPES', JSON.stringify(normalized), 'Session recording input labels');
  return { success: true, types: normalized };
}

function getSessionRecordingsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const types = getSessionRecordingTypes();
  const sessions = (sessionsResult.sessions || []).map(s => {
    let jsonLinks = {};
    if (s.recordingLinksJson) {
      try { jsonLinks = JSON.parse(s.recordingLinksJson || '{}') || {}; } catch (e) {}
    }
    const links = Object.assign({
      'GoPro': s.goproLink || '',
      'Tascam': s.tascamLink || '',
      'Meeting Owl': s.meetingOwlLink || ''
    }, jsonLinks);
    return Object.assign({}, s, { recordingLinks: links });
  });
  return { success: true, sessions: sessions, recordingTypes: types };
}


function getFieldNotesByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const sessionsResult = getSessionsByRollout(token, rolloutId);
  if (!sessionsResult.success) return sessionsResult;
  const sessions = (sessionsResult.sessions || []).map(function(s) {
    return Object.assign({}, s, { fieldNotesLink: s.fieldNotesLink || '' });
  });
  return { success: true, sessions: sessions };
}

function saveFieldNotesLink(token, sessionId, link) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessionIdIdx = headers.indexOf('sessionId');
  if (sessionIdIdx === -1) return { success: false, message: 'sessionId column missing' };

  let fieldNotesIdx = headers.indexOf('fieldNotesLink');
  if (fieldNotesIdx === -1) {
    headers.push('fieldNotesLink');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fieldNotesIdx = headers.length - 1;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sessionIdIdx]) === String(sessionId)) {
      sheet.getRange(i + 1, fieldNotesIdx + 1).setValue(String(link || '').trim());
      invalidateSheetSnapshot('StudySessions');
      return { success: true, message: 'Field notes link saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function saveSessionRecordingLinks(token, sessionId, links) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') return { success: false, message: 'Unauthorized' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('StudySessions');
  if (!sheet) return { success: false, message: 'StudySessions not found' };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ensureCol = name => {
    let idx = headers.indexOf(name);
    if (idx === -1) {
      headers.push(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      idx = headers.length - 1;
    }
    return idx + 1;
  };
  const col = name => headers.indexOf(name) + 1;
  const jsonCol = ensureCol('recordingLinksJson');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('sessionId')]) === String(sessionId)) {
      if (col('goproLink') > 0) sheet.getRange(i + 1, col('goproLink')).setValue(links.GoPro || '');
      if (col('tascamLink') > 0) sheet.getRange(i + 1, col('tascamLink')).setValue(links.Tascam || '');
      if (col('meetingOwlLink') > 0) sheet.getRange(i + 1, col('meetingOwlLink')).setValue(links['Meeting Owl'] || '');
      sheet.getRange(i + 1, jsonCol).setValue(JSON.stringify(links || {}));
      return { success: true, message: 'Session recording links saved' };
    }
  }
  return { success: false, message: 'Session not found' };
}

function upsertConfigValue(key, value, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  const descCol = headers.indexOf('description');
  const updatedAtCol = headers.indexOf('updatedAt');
  const timestamp = new Date().toISOString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      const row = data[i].slice(0, headers.length);
      row[valueCol] = value;
      row[descCol] = description || '';
      row[updatedAtCol] = timestamp;
      configSheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      invalidateConfigRuntimeCache();
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', timestamp]);
  invalidateConfigRuntimeCache();
}


// ============================================
// DIGITAL GOOGLE FORMS PROTOCOL SYNC
// ============================================

const DIGITAL_FORM_SYNC_CONFIG_KEY = 'DIGITAL_FORM_SYNC_CONFIG';
const DIGITAL_FORM_SYNC_WATERMARKS_KEY = 'DIGITAL_FORM_SYNC_WATERMARKS';
const DIGITAL_FORM_SYNC_SYSTEM_USER = 'Google Forms Sync';
const DIGITAL_FORM_SYNC_ACCESS_KEY = 'DIGITAL_FORM_SYNC_ACCESS_ADMINS';
const DIGITAL_FORM_SYNC_SUPER_ADMIN_ID = '03456e13-c1fc-45c4-ad4a-28e06d23cb6d';
const DIGITAL_FORM_SYNC_SUPER_ADMIN_USERNAME = 'david';
const ENROLLMENT_FIELD_MANAGER_ACCESS_KEY = 'ENROLLMENT_FIELD_MANAGER_ACCESS_ADMINS';

function isDigitalFormSyncSuperAdmin(user) {
  if (!user || user.role !== 'admin') return false;
  const userId = String(user.userId || '').trim();
  const username = String(user.username || '').trim().toLowerCase();
  return userId === DIGITAL_FORM_SYNC_SUPER_ADMIN_ID || username === DIGITAL_FORM_SYNC_SUPER_ADMIN_USERNAME;
}

function getDigitalFormSyncAllowedAdminIdsInternal() {
  return getSensitiveFeatureAllowedAdminIdsInternal(DIGITAL_FORM_SYNC_ACCESS_KEY);
}

function getEnrollmentFieldManagerAllowedAdminIdsInternal() {
  return getSensitiveFeatureAllowedAdminIdsInternal(ENROLLMENT_FIELD_MANAGER_ACCESS_KEY);
}

function getSensitiveFeatureAllowedAdminIdsInternal(configKey) {
  const allowed = {};
  allowed[DIGITAL_FORM_SYNC_SUPER_ADMIN_ID] = true;
  const configMap = getConfigMap();
  const raw = configMap[configKey];
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const ids = Array.isArray(parsed) ? parsed : (parsed && Array.isArray(parsed.allowedAdminIds) ? parsed.allowedAdminIds : []);
      ids.forEach(id => {
        const clean = String(id || '').trim();
        if (clean) allowed[clean] = true;
      });
    } catch (e) {
      String(raw).split(',').forEach(id => {
        const clean = String(id || '').trim();
        if (clean) allowed[clean] = true;
      });
    }
  }
  return Object.keys(allowed);
}

function isUserAllowedForSensitiveFeature(user, configKey) {
  if (!user || user.role !== 'admin') return false;
  if (isDigitalFormSyncSuperAdmin(user)) return true;
  const userId = String(user.userId || '').trim();
  return !!userId && getSensitiveFeatureAllowedAdminIdsInternal(configKey).indexOf(userId) !== -1;
}

function canAccessDigitalFormSync(user) {
  return isUserAllowedForSensitiveFeature(user, DIGITAL_FORM_SYNC_ACCESS_KEY);
}

function canAccessEnrollmentFieldManager(user) {
  return isUserAllowedForSensitiveFeature(user, ENROLLMENT_FIELD_MANAGER_ACCESS_KEY);
}

function getDigitalFormSyncAccessControlForUser(user) {
  return {
    hasAccess: canAccessDigitalFormSync(user),
    isSuperAdmin: isDigitalFormSyncSuperAdmin(user),
    superAdminUserId: DIGITAL_FORM_SYNC_SUPER_ADMIN_ID,
    allowedAdminIds: getDigitalFormSyncAllowedAdminIdsInternal()
  };
}

function getEnrollmentFieldManagerAccessControlForUser(user) {
  return {
    hasAccess: canAccessEnrollmentFieldManager(user),
    isSuperAdmin: isDigitalFormSyncSuperAdmin(user),
    superAdminUserId: DIGITAL_FORM_SYNC_SUPER_ADMIN_ID,
    allowedAdminIds: getEnrollmentFieldManagerAllowedAdminIdsInternal()
  };
}

function getAdminUsersForDigitalFormSyncAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const idx = header => headers.indexOf(header);
  const allowed = getDigitalFormSyncAllowedAdminIdsInternal();
  return values.slice(1)
    .map(row => ({
      userId: row[idx('userId')] || '',
      username: row[idx('username')] || '',
      fullName: row[idx('fullName')] || row[idx('name')] || '',
      role: row[idx('role')] || '',
      site: row[idx('site')] || '',
      status: row[idx('status')] || ''
    }))
    .filter(user => String(user.role).toLowerCase() === 'admin' && String(user.status || 'active').toLowerCase() !== 'inactive')
    .map(user => Object.assign({}, user, {
      isSuperAdmin: isDigitalFormSyncSuperAdmin(user),
      hasAccess: isDigitalFormSyncSuperAdmin(user) || allowed.indexOf(String(user.userId || '').trim()) !== -1
    }));
}

function saveDigitalFormSyncAccess(token, adminUserIds) {
  const currentUser = validateSession(token);
  if (!isDigitalFormSyncSuperAdmin(currentUser)) {
    return { success: false, message: 'Only the primary Digital Forms Sync administrator can manage access.' };
  }
  const allowed = {};
  allowed[DIGITAL_FORM_SYNC_SUPER_ADMIN_ID] = true;
  (Array.isArray(adminUserIds) ? adminUserIds : []).forEach(id => {
    const clean = String(id || '').trim();
    if (clean) allowed[clean] = true;
  });
  upsertConfigValue(DIGITAL_FORM_SYNC_ACCESS_KEY, JSON.stringify(Object.keys(allowed)), 'Admin users allowed to manage Digital Forms Sync');
  return {
    success: true,
    message: 'Digital Forms Sync access updated',
    accessControl: getDigitalFormSyncAccessControlForUser(currentUser),
    adminUsers: getAdminUsersForDigitalFormSyncAccess()
  };
}

function saveSensitiveFeatureAllowedAdminIdsInternal(configKey, adminUserIds, description) {
  const allowed = {};
  allowed[DIGITAL_FORM_SYNC_SUPER_ADMIN_ID] = true;
  (Array.isArray(adminUserIds) ? adminUserIds : []).forEach(id => {
    const clean = String(id || '').trim();
    if (clean) allowed[clean] = true;
  });
  upsertConfigValue(configKey, JSON.stringify(Object.keys(allowed)), description);
}

function getAdminFeatureAccessData(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  const isSuperAdmin = isDigitalFormSyncSuperAdmin(currentUser);
  const digitalAllowed = getDigitalFormSyncAllowedAdminIdsInternal();
  const enrollmentAllowed = getEnrollmentFieldManagerAllowedAdminIdsInternal();
  return {
    success: true,
    isSuperAdmin: isSuperAdmin,
    digitalFormsSync: getDigitalFormSyncAccessControlForUser(currentUser),
    enrollmentFieldManager: getEnrollmentFieldManagerAccessControlForUser(currentUser),
    adminUsers: isSuperAdmin ? getAdminUsersForDigitalFormSyncAccess().map(user => Object.assign({}, user, {
      digitalFormsSyncAccess: user.isSuperAdmin || digitalAllowed.indexOf(String(user.userId || '').trim()) !== -1,
      enrollmentFieldManagerAccess: user.isSuperAdmin || enrollmentAllowed.indexOf(String(user.userId || '').trim()) !== -1
    })) : []
  };
}

function saveAdminFeatureAccess(token, access) {
  const currentUser = validateSession(token);
  if (!isDigitalFormSyncSuperAdmin(currentUser)) {
    return { success: false, message: 'Only david can manage access to sensitive admin features.' };
  }
  const payload = access || {};
  saveSensitiveFeatureAllowedAdminIdsInternal(
    DIGITAL_FORM_SYNC_ACCESS_KEY,
    payload.digitalFormsSyncAdminIds || [],
    'Admin users allowed to manage Digital Forms Sync'
  );
  saveSensitiveFeatureAllowedAdminIdsInternal(
    ENROLLMENT_FIELD_MANAGER_ACCESS_KEY,
    payload.enrollmentFieldManagerAdminIds || [],
    'Admin users allowed to manage Enrollment Field Manager'
  );
  return Object.assign({ message: 'Admin feature access updated' }, getAdminFeatureAccessData(token));
}

function ensureDigitalFormSyncSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheetIfNotExists(ss, 'FormResponseSyncLog', [
    'syncId', 'sourceWorkbookId', 'sourceSheetName', 'sourceRowNumber', 'sourceTimestamp',
    'responseHash', 'rawName', 'normalizedName', 'instrumentNumber', 'instrumentName',
    'candidateRolloutId', 'candidateSessionId', 'participantId', 'checklistId', 'matchStatus',
    'confidenceScore', 'matchedBy', 'matchedAt', 'notes', 'sourceLink'
  ]);
  createSheetIfNotExists(ss, 'FormResponseMatchQueue', [
    'queueId', 'createdAt', 'sourceWorkbookId', 'sourceSheetName', 'sourceRowNumber', 'sourceTimestamp',
    'rawName', 'normalizedName', 'instrumentNumber', 'instrumentName', 'candidateRolloutId',
    'candidateSessionId', 'suggestedParticipantId', 'suggestedParticipantName', 'confidenceScore',
    'status', 'reviewedBy', 'reviewedAt', 'notes', 'sourceLink', 'responseHash'
  ]);
  createSheetIfNotExists(ss, 'ParticipantAliases', [
    'aliasId', 'participantId', 'alias', 'normalizedAlias', 'createdAt', 'createdBy', 'notes'
  ]);
}

function getDefaultDigitalFormSyncConfig() {
  return {
    enabled: false,
    masterWorkbookId: '',
    masterWorkbookIds: { UGA: '', Missouri: '' },
    autoMatchThreshold: 90,
    reviewMatchThreshold: 70,
    lateResponseWindowDays: 0,
    enableRecentCompletedFallback: false,
    recentCompletedFallbackDays: 30,
    recentCompletedFallbackMaxCohorts: 1,
    allowRecentCompletedAutoMatch: false,
    preferPresentAttendance: true,
    updateMissingChecklistItems: true,
    saveAliasesOnReview: true,
    mappings: [],
    siteMappings: { UGA: [], Missouri: [] }
  };
}

function extractSpreadsheetId(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  return raw.replace(/[?#].*$/, '');
}

function normalizeDigitalFormSyncConfig(config) {
  const defaults = getDefaultDigitalFormSyncConfig();
  const input = config || {};
  const normalized = Object.assign({}, defaults, input);
  const workbookIdsInput = input.masterWorkbookIds || {};
  const legacyWorkbookId = extractSpreadsheetId(normalized.masterWorkbookId || normalized.masterWorkbookUrl || '');
  normalized.masterWorkbookIds = {
    UGA: extractSpreadsheetId(workbookIdsInput.UGA || workbookIdsInput.uga || input.ugaMasterWorkbookId || ''),
    Missouri: extractSpreadsheetId(workbookIdsInput.Missouri || workbookIdsInput.missouri || input.missouriMasterWorkbookId || '')
  };
  if (legacyWorkbookId && !normalized.masterWorkbookIds.UGA && !normalized.masterWorkbookIds.Missouri) {
    normalized.masterWorkbookIds.UGA = legacyWorkbookId;
  }
  normalized.masterWorkbookId = normalized.masterWorkbookIds.UGA || normalized.masterWorkbookIds.Missouri || legacyWorkbookId;
  normalized.autoMatchThreshold = Math.min(100, Math.max(0, Number(normalized.autoMatchThreshold) || defaults.autoMatchThreshold));
  normalized.reviewMatchThreshold = Math.min(normalized.autoMatchThreshold, Math.max(0, Number(normalized.reviewMatchThreshold) || defaults.reviewMatchThreshold));
  normalized.lateResponseWindowDays = Math.max(0, Number(normalized.lateResponseWindowDays) || 0);
  normalized.enableRecentCompletedFallback = !!normalized.enableRecentCompletedFallback;
  normalized.recentCompletedFallbackDays = Math.max(0, Number(normalized.recentCompletedFallbackDays) || defaults.recentCompletedFallbackDays);
  normalized.recentCompletedFallbackMaxCohorts = Math.max(1, Number(normalized.recentCompletedFallbackMaxCohorts) || defaults.recentCompletedFallbackMaxCohorts);
  normalized.allowRecentCompletedAutoMatch = !!normalized.allowRecentCompletedAutoMatch;

  const normalizeMapping = mapping => ({
    enabled: !!mapping.enabled,
    instrumentNumber: String(mapping.instrumentNumber || '').trim(),
    instrumentName: String(mapping.instrumentName || '').trim(),
    sheetName: String(mapping.sheetName || '').trim(),
    timestampColumn: String(mapping.timestampColumn || 'Timestamp').trim() || 'Timestamp',
    nameColumn: String(mapping.nameColumn || '').trim(),
    allowLateResponses: mapping.allowLateResponses !== false
  });
  const legacyMappings = Array.isArray(input.mappings) ? input.mappings.map(normalizeMapping).filter(mapping => mapping.instrumentNumber && mapping.sheetName) : [];
  const inputSiteMappings = input.siteMappings || {};
  normalized.siteMappings = { UGA: [], Missouri: [] };
  ['UGA', 'Missouri'].forEach(site => {
    const siteInput = inputSiteMappings[site] || inputSiteMappings[String(site).toLowerCase()] || null;
    normalized.siteMappings[site] = Array.isArray(siteInput) && siteInput.length
      ? siteInput.map(normalizeMapping).filter(mapping => mapping.instrumentNumber && mapping.sheetName)
      : legacyMappings.slice();
  });
  normalized.mappings = legacyMappings.length ? legacyMappings : (normalized.siteMappings.UGA || []);
  return normalized;
}

function getDigitalFormMappingsForSite(config, site) {
  const normalizedSite = String(site || '').toLowerCase() === 'missouri' ? 'Missouri' : 'UGA';
  const siteMappings = config && config.siteMappings && Array.isArray(config.siteMappings[normalizedSite]) ? config.siteMappings[normalizedSite] : [];
  return siteMappings.length ? siteMappings : ((config && config.mappings) || []);
}

function getDigitalFormSyncConfigInternal() {
  const configMap = getConfigMap();
  if (!configMap[DIGITAL_FORM_SYNC_CONFIG_KEY]) return getDefaultDigitalFormSyncConfig();
  try {
    return normalizeDigitalFormSyncConfig(JSON.parse(configMap[DIGITAL_FORM_SYNC_CONFIG_KEY] || '{}'));
  } catch (e) {
    return getDefaultDigitalFormSyncConfig();
  }
}

function getDigitalFormSyncWatermarksInternal() {
  const configMap = getConfigMap();
  if (!configMap[DIGITAL_FORM_SYNC_WATERMARKS_KEY]) return {};
  try { return JSON.parse(configMap[DIGITAL_FORM_SYNC_WATERMARKS_KEY] || '{}') || {}; } catch (e) { return {}; }
}

function saveDigitalFormSyncWatermarksInternal(watermarks) {
  upsertConfigValue(DIGITAL_FORM_SYNC_WATERMARKS_KEY, JSON.stringify(watermarks || {}), 'Per-form response sync watermarks');
}

function saveDigitalFormSyncConfig(token, config) {
  const currentUser = validateSession(token);
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  ensureDigitalFormSyncSheets();
  const normalized = normalizeDigitalFormSyncConfig(config || {});
  upsertConfigValue(DIGITAL_FORM_SYNC_CONFIG_KEY, JSON.stringify(normalized), 'Google Forms digital protocol sync configuration');
  return { success: true, message: 'Digital Forms Sync configuration saved', config: normalized };
}

function getDigitalFormSyncConfig(token) {
  const currentUser = validateSession(token);
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  ensureDigitalFormSyncSheets();
  return { success: true, config: getDigitalFormSyncConfigInternal(), accessControl: getDigitalFormSyncAccessControlForUser(currentUser) };
}

function openDigitalFormWorkbook(workbookId) {
  const id = extractSpreadsheetId(workbookId);
  if (!id) throw new Error('Master workbook ID is required');
  return SpreadsheetApp.openById(id);
}

function getWorkbookTabsForDigitalSync(workbookId) {
  const workbook = openDigitalFormWorkbook(workbookId);
  return workbook.getSheets().map(sheet => ({
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    lastRow: sheet.getLastRow(),
    headers: sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean) : []
  }));
}

function testDigitalFormWorkbook(token, workbookId) {
  const currentUser = validateSession(token);
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  try {
    const tabs = getWorkbookTabsForDigitalSync(workbookId);
    return { success: true, message: 'Connected to master workbook', tabs: tabs };
  } catch (error) {
    return { success: false, message: 'Unable to connect: ' + error.message };
  }
}

function getDigitalFormSyncAdminData(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  if (!canAccessDigitalFormSync(currentUser)) {
    return {
      success: true,
      restricted: true,
      message: 'Digital Forms Sync is restricted. Ask the primary administrator to grant access.',
      accessControl: getDigitalFormSyncAccessControlForUser(currentUser)
    };
  }
  ensureDigitalFormSyncSheets();
  const config = getDigitalFormSyncConfigInternal();
  let tabs = [];
  let workbookError = '';
  const tabsBySite = {};
  ['UGA', 'Missouri'].forEach(site => {
    const workbookId = config.masterWorkbookIds && config.masterWorkbookIds[site];
    if (!workbookId) return;
    try {
      tabsBySite[site] = getWorkbookTabsForDigitalSync(workbookId);
      if (!tabs.length) tabs = tabsBySite[site];
    } catch (e) {
      tabsBySite[site] = [];
      workbookError += (workbookError ? '; ' : '') + site + ': ' + e.message;
    }
  });
  if (!tabs.length && config.masterWorkbookId) {
    try {
      tabs = getWorkbookTabsForDigitalSync(config.masterWorkbookId);
      if (!tabsBySite.UGA || !tabsBySite.UGA.length) tabsBySite.UGA = tabs;
    } catch (e) {
      workbookError += (workbookError ? '; ' : '') + e.message;
    }
  }
  return {
    success: true,
    config: config,
    tabs: tabs,
    tabsBySite: tabsBySite,
    workbookError: workbookError,
    protocolItems: getGlobalProtocolItems(),
    queue: getDigitalFormMatchQueueInternal({ status: 'pending', limit: 50 }),
    triggers: getDigitalFormSyncTriggerSummary(),
    accessControl: getDigitalFormSyncAccessControlForUser(currentUser),
    adminUsers: isDigitalFormSyncSuperAdmin(currentUser) ? getAdminUsersForDigitalFormSyncAccess() : []
  };
}

function normalizePersonNameForMatch(name) {
  return String(name || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function levenshteinDistance(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a === b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;
  const prev = [];
  const curr = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= b.length; j++) prev[j] = curr[j];
  }
  return prev[b.length];
}

function fuzzyStringScore(a, b) {
  const left = normalizePersonNameForMatch(a);
  const right = normalizePersonNameForMatch(b);
  if (!left || !right) return 0;
  if (left === right) return 100;
  const distance = levenshteinDistance(left, right);
  const maxLen = Math.max(left.length, right.length) || 1;
  return Math.max(0, Math.round((1 - distance / maxLen) * 100));
}

function scoreParticipantNameMatch(rawName, participant, aliases, presentParticipantIds) {
  const entered = normalizePersonNameForMatch(rawName);
  const fullName = normalizePersonNameForMatch(participant.fullName);
  if (!entered || !fullName) return { score: 0, reason: 'Missing name' };
  const enteredTokens = entered.split(' ').filter(Boolean);
  const fullTokens = fullName.split(' ').filter(Boolean);
  let score = fuzzyStringScore(entered, fullName);
  let reason = 'Fuzzy full-name comparison';
  if (entered === fullName) {
    score = 100;
    reason = 'Exact full-name match';
  } else if (enteredTokens.length === 1 && fullTokens[0] === enteredTokens[0]) {
    score = Math.max(score, 82);
    reason = 'First-name match';
  } else if (enteredTokens.length >= 2 && fullTokens.length >= 2 && enteredTokens[0] === fullTokens[0] && enteredTokens[enteredTokens.length - 1] === fullTokens[fullTokens.length - 1]) {
    score = Math.max(score, 95);
    reason = 'First and last name match';
  } else if (enteredTokens[0] && fullTokens[0] && enteredTokens[0] === fullTokens[0]) {
    score = Math.max(score, 74);
    reason = 'Shared first name with fuzzy remainder';
  }

  const aliasList = aliases[String(participant.participantId)] || [];
  aliasList.forEach(alias => {
    const aliasScore = fuzzyStringScore(entered, alias.normalizedAlias || alias.alias);
    if (aliasScore > score) {
      score = aliasScore;
      reason = 'Participant alias match';
    }
  });

  if (presentParticipantIds && presentParticipantIds[String(participant.participantId)] && score > 0) {
    score = Math.min(100, score + 5);
    reason += ' + present attendance';
  }
  return { score: score, reason: reason };
}

function chooseBestParticipantMatch(rawName, candidates, aliases, presentParticipantIds) {
  const scored = candidates.map(participant => {
    const scoredMatch = scoreParticipantNameMatch(rawName, participant, aliases, presentParticipantIds);
    return Object.assign({}, participant, { score: scoredMatch.score, reason: scoredMatch.reason });
  }).sort((a, b) => b.score - a.score || String(a.fullName || '').localeCompare(String(b.fullName || '')));

  const best = scored[0] || null;
  const second = scored[1] || null;
  if (!best) return { participant: null, score: 0, reason: 'No candidate participants' };

  const enteredTokens = normalizePersonNameForMatch(rawName).split(' ').filter(Boolean);
  if (enteredTokens.length === 1) {
    const sameFirstName = scored.filter(candidate => normalizePersonNameForMatch(candidate.fullName).split(' ')[0] === enteredTokens[0]);
    if (sameFirstName.length > 1) {
      return { participant: best, score: Math.min(best.score, 69), reason: 'First name is not unique in candidate cohort' };
    }
  }

  if (second && best.score - second.score < 8 && best.score < 95) {
    return { participant: best, score: Math.min(best.score, 79), reason: 'Close competing participant match' };
  }
  return { participant: best, score: best.score, reason: best.reason };
}

function getDigitalFormMatchQueueInternal(options) {
  ensureDigitalFormSyncSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('FormResponseMatchQueue');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const statusFilter = options && options.status;
  const limit = Number(options && options.limit) || 100;
  const rows = [];
  for (let i = data.length - 1; i >= 1 && rows.length < limit; i--) {
    const row = data[i];
    const status = row[headers.indexOf('status')];
    if (statusFilter && status !== statusFilter) continue;
    rows.push({
      queueId: row[headers.indexOf('queueId')],
      createdAt: row[headers.indexOf('createdAt')],
      sourceSheetName: row[headers.indexOf('sourceSheetName')],
      sourceRowNumber: row[headers.indexOf('sourceRowNumber')],
      sourceTimestamp: row[headers.indexOf('sourceTimestamp')],
      rawName: row[headers.indexOf('rawName')],
      instrumentNumber: row[headers.indexOf('instrumentNumber')],
      instrumentName: row[headers.indexOf('instrumentName')],
      candidateRolloutId: row[headers.indexOf('candidateRolloutId')],
      candidateSessionId: row[headers.indexOf('candidateSessionId')],
      suggestedParticipantId: row[headers.indexOf('suggestedParticipantId')],
      suggestedParticipantName: row[headers.indexOf('suggestedParticipantName')],
      confidenceScore: row[headers.indexOf('confidenceScore')],
      status: status,
      notes: row[headers.indexOf('notes')],
      sourceLink: row[headers.indexOf('sourceLink')]
    });
  }
  return rows;
}

function buildDigitalSyncDatasets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const sessionsSnapshot = getSheetSnapshot('StudySessions');
  const attendanceSnapshot = getSheetSnapshot('SessionAttendance');
  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const aliasSheet = ss.getSheetByName('ParticipantAliases');
  const aliases = {};
  if (aliasSheet && aliasSheet.getLastRow() > 1) {
    const aliasData = aliasSheet.getDataRange().getValues();
    const aliasHeaders = aliasData[0] || [];
    for (let i = 1; i < aliasData.length; i++) {
      const pid = String(aliasData[i][aliasHeaders.indexOf('participantId')] || '');
      if (!pid) continue;
      aliases[pid] = aliases[pid] || [];
      aliases[pid].push({ alias: aliasData[i][aliasHeaders.indexOf('alias')], normalizedAlias: aliasData[i][aliasHeaders.indexOf('normalizedAlias')] });
    }
  }

  const rolloutsSnapshot = getSheetSnapshot('StudyRollouts');
  const rHeaders = rolloutsSnapshot.headers;
  const rolloutsById = {};
  for (let i = 1; i < rolloutsSnapshot.data.length; i++) {
    const row = rolloutsSnapshot.data[i];
    const rollout = {
      rolloutId: row[rHeaders.indexOf('rolloutId')],
      site: row[rHeaders.indexOf('site')],
      schoolName: row[rHeaders.indexOf('schoolName')],
      period: row[rHeaders.indexOf('period')],
      year: row[rHeaders.indexOf('year')],
      status: row[rHeaders.indexOf('status')]
    };
    if (rollout.rolloutId) rolloutsById[String(rollout.rolloutId)] = rollout;
  }

  const pHeaders = participantsSnapshot.headers;
  const participants = [];
  const participantsByRollout = {};
  for (let i = 1; i < participantsSnapshot.data.length; i++) {
    const row = participantsSnapshot.data[i];
    const participant = {
      participantId: row[pHeaders.indexOf('participantId')],
      fullName: row[pHeaders.indexOf('fullName')],
      site: row[pHeaders.indexOf('site')],
      rolloutId: row[pHeaders.indexOf('rolloutId')],
      status: row[pHeaders.indexOf('status')]
    };
    if (!participant.participantId) continue;
    participants.push(participant);
    participantsByRollout[String(participant.rolloutId)] = participantsByRollout[String(participant.rolloutId)] || [];
    participantsByRollout[String(participant.rolloutId)].push(participant);
  }

  const sHeaders = sessionsSnapshot.headers;
  const sessionsByDate = {};
  const sessionById = {};
  const rolloutSessionBounds = {};
  for (let i = 1; i < sessionsSnapshot.data.length; i++) {
    const row = sessionsSnapshot.data[i];
    const session = {
      sessionId: row[sHeaders.indexOf('sessionId')],
      rolloutId: row[sHeaders.indexOf('rolloutId')],
      sessionNumber: row[sHeaders.indexOf('sessionNumber')],
      sessionDate: normalizeSessionDateValue(row[sHeaders.indexOf('sessionDate')]),
      status: row[sHeaders.indexOf('status')]
    };
    const dateOnly = getDateOnly(parseSessionDate(session.sessionDate));
    if (!dateOnly) continue;
    session.dateOnly = dateOnly;
    sessionById[String(session.sessionId)] = session;
    sessionsByDate[dateOnly] = sessionsByDate[dateOnly] || [];
    sessionsByDate[dateOnly].push(session);
    const rolloutIdKey = String(session.rolloutId || '');
    rolloutSessionBounds[rolloutIdKey] = rolloutSessionBounds[rolloutIdKey] || { firstDate: dateOnly, lastDate: dateOnly, sessionCount: 0 };
    if (compareDateStrings(dateOnly, rolloutSessionBounds[rolloutIdKey].firstDate) < 0) rolloutSessionBounds[rolloutIdKey].firstDate = dateOnly;
    if (compareDateStrings(dateOnly, rolloutSessionBounds[rolloutIdKey].lastDate) > 0) rolloutSessionBounds[rolloutIdKey].lastDate = dateOnly;
    rolloutSessionBounds[rolloutIdKey].sessionCount++;
  }

  const completedRolloutsBySite = {};
  const todayKey = getDateOnly(new Date());
  Object.keys(rolloutsById).forEach(rolloutId => {
    const bounds = rolloutSessionBounds[rolloutId];
    if (!bounds || !bounds.lastDate) return;
    if (compareDateStrings(bounds.lastDate, todayKey) > 0) return;
    const rollout = Object.assign({}, rolloutsById[rolloutId], bounds);
    const siteKey = String(rollout.site || '').toLowerCase();
    completedRolloutsBySite[siteKey] = completedRolloutsBySite[siteKey] || [];
    completedRolloutsBySite[siteKey].push(rollout);
  });
  Object.keys(completedRolloutsBySite).forEach(siteKey => completedRolloutsBySite[siteKey].sort((a, b) => compareDateStrings(b.lastDate, a.lastDate)));

  const aHeaders = attendanceSnapshot.headers;
  const presentBySession = {};
  for (let i = 1; i < attendanceSnapshot.data.length; i++) {
    const row = attendanceSnapshot.data[i];
    if (String(row[aHeaders.indexOf('status')] || '') !== 'present') continue;
    const sessionId = String(row[aHeaders.indexOf('sessionId')] || '');
    presentBySession[sessionId] = presentBySession[sessionId] || {};
    presentBySession[sessionId][String(row[aHeaders.indexOf('participantId')] || '')] = true;
  }

  const cHeaders = checklistSnapshot.headers;
  const checklistByParticipantInstrument = {};
  for (let i = 1; i < checklistSnapshot.data.length; i++) {
    const row = checklistSnapshot.data[i];
    const key = String(row[cHeaders.indexOf('participantId')]) + '|' + String(row[cHeaders.indexOf('instrumentNumber')]);
    checklistByParticipantInstrument[key] = { rowIndex: i, row: row };
  }

  return {
    participants: participants,
    participantsByRollout: participantsByRollout,
    rolloutsById: rolloutsById,
    completedRolloutsBySite: completedRolloutsBySite,
    sessionsByDate: sessionsByDate,
    sessionById: sessionById,
    presentBySession: presentBySession,
    aliases: aliases,
    checklistSnapshot: checklistSnapshot,
    checklistByParticipantInstrument: checklistByParticipantInstrument
  };
}

function getDigitalSessionMatchesForResponse(responseDate, datasets, config, mapping) {
  if (!responseDate) return { sessions: [], strategy: 'no_timestamp', notes: 'No response date available' };
  const exact = datasets.sessionsByDate[responseDate] || [];
  if (exact.length) return { sessions: exact, strategy: 'session_date', notes: 'Matched by exact session date ' + responseDate };
  if (!mapping.allowLateResponses || !config.lateResponseWindowDays) {
    return { sessions: [], strategy: 'no_session', notes: 'No DARTS session found for response date ' + responseDate };
  }
  const target = parseDateOnlyString(responseDate);
  if (!target) return { sessions: [], strategy: 'no_timestamp', notes: 'Response date could not be parsed' };
  const matches = [];
  Object.keys(datasets.sessionsByDate).forEach(dateKey => {
    const sessionDate = parseDateOnlyString(dateKey);
    if (!sessionDate) return;
    const diffDays = Math.abs(Math.round((target.getTime() - sessionDate.getTime()) / 86400000));
    if (diffDays <= config.lateResponseWindowDays) {
      datasets.sessionsByDate[dateKey].forEach(session => matches.push(session));
    }
  });
  return matches.length
    ? { sessions: matches, strategy: 'late_window', notes: 'Matched by late response window around ' + responseDate }
    : { sessions: [], strategy: 'no_session', notes: 'No DARTS session found within late-response window for ' + responseDate };
}

function findSessionsForDigitalResponse(responseDate, datasets, config, mapping) {
  return getDigitalSessionMatchesForResponse(responseDate, datasets, config, mapping).sessions;
}

function getRecentCompletedFallbackCandidates(responseDate, site, datasets, config) {
  if (!config.enableRecentCompletedFallback) return { candidates: [], rollout: null, strategy: '', notes: '' };
  const siteKey = String(site || '').toLowerCase();
  if (!siteKey) return { candidates: [], rollout: null, strategy: '', notes: '' };
  const responseDateKey = responseDate || getDateOnly(new Date());
  const responseDateObj = parseDateOnlyString(responseDateKey);
  const maxCohorts = Math.max(1, Number(config.recentCompletedFallbackMaxCohorts) || 1);
  const graceDays = Math.max(0, Number(config.recentCompletedFallbackDays) || 0);
  const recentRollouts = (datasets.completedRolloutsBySite[siteKey] || []).filter(rollout => {
    if (!responseDateObj || !rollout.lastDate) return true;
    if (compareDateStrings(rollout.lastDate, responseDateKey) > 0) return false;
    const lastDate = parseDateOnlyString(rollout.lastDate);
    if (!lastDate) return true;
    const diffDays = Math.round((responseDateObj.getTime() - lastDate.getTime()) / 86400000);
    return diffDays >= 0 && diffDays <= graceDays;
  }).slice(0, maxCohorts);
  const candidates = [];
  recentRollouts.forEach(rollout => {
    (datasets.participantsByRollout[String(rollout.rolloutId)] || []).forEach(participant => {
      if (String(participant.status || '').toLowerCase() !== 'withdrawn') candidates.push(participant);
    });
  });
  const rollout = recentRollouts[0] || null;
  return {
    candidates: candidates,
    rollout: rollout,
    strategy: candidates.length ? 'recent_completed_cohort' : '',
    notes: candidates.length && rollout ? 'Fallback searched recent completed cohort: ' + rollout.schoolName + ' (' + rollout.period + ' ' + rollout.year + ')' : ''
  };
}

function buildDigitalResponseHash(values) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(values || []))).substring(0, 32);
}

function getExistingDigitalSyncKeys() {
  ensureDigitalFormSyncSheets();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormResponseSyncLog');
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const keys = {};
  for (let i = 1; i < data.length; i++) {
    const sourceWorkbookId = data[i][headers.indexOf('sourceWorkbookId')];
    const sourceSheet = data[i][headers.indexOf('sourceSheetName')];
    const sourceRow = data[i][headers.indexOf('sourceRowNumber')];
    const hash = data[i][headers.indexOf('responseHash')];
    keys[String(sourceWorkbookId) + '::' + String(sourceSheet) + '::' + String(sourceRow) + '::' + String(hash)] = true;
  }
  return keys;
}

function getDigitalSourceLink(workbookId, sheet, rowNumber) {
  return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(workbookId) + '/edit#gid=' + sheet.getSheetId() + '&range=' + rowNumber + ':' + rowNumber;
}

function analyzeDigitalFormResponse(mapping, sheet, headers, row, rowNumber, workbookId, datasets, config, site) {
  const timestampCol = headers.indexOf(mapping.timestampColumn || 'Timestamp');
  const nameCol = headers.indexOf(mapping.nameColumn || '');
  const responseTimestamp = timestampCol >= 0 ? row[timestampCol] : '';
  const rawName = nameCol >= 0 ? row[nameCol] : '';
  const responseDate = getDateOnly(parseSessionDate(responseTimestamp));
  const responseHash = buildDigitalResponseHash(row);
  const sourceLink = getDigitalSourceLink(workbookId, sheet, rowNumber);
  const normalizedName = normalizePersonNameForMatch(rawName);
  const sessionMatch = getDigitalSessionMatchesForResponse(responseDate, datasets, config, mapping);
  const sessions = sessionMatch.sessions || [];
  const rolloutIds = {};
  sessions.forEach(session => rolloutIds[String(session.rolloutId)] = true);
  let candidates = [];
  Object.keys(rolloutIds).forEach(rolloutId => {
    (datasets.participantsByRollout[rolloutId] || []).forEach(participant => {
      const sameSite = !site || String(participant.site || '').toLowerCase() === String(site || '').toLowerCase();
      if (sameSite && String(participant.status || '').toLowerCase() !== 'withdrawn') candidates.push(participant);
    });
  });
  const presentParticipantIds = {};
  if (config.preferPresentAttendance) {
    sessions.forEach(session => {
      Object.assign(presentParticipantIds, datasets.presentBySession[String(session.sessionId)] || {});
    });
  }
  let match = chooseBestParticipantMatch(rawName, candidates, datasets.aliases, presentParticipantIds);
  let participant = match.participant;
  let selectedSession = sessions[0] || null;
  let matchStrategy = sessionMatch.strategy || 'session_date';
  let strategyNotes = sessionMatch.notes || '';
  let fallbackUsed = false;

  if ((!participant || (match.score || 0) < config.reviewMatchThreshold) && config.enableRecentCompletedFallback) {
    const fallback = getRecentCompletedFallbackCandidates(responseDate, site, datasets, config);
    if (fallback.candidates.length) {
      const fallbackMatch = chooseBestParticipantMatch(rawName, fallback.candidates, datasets.aliases, {});
      if (!participant || fallbackMatch.score > (match.score || 0)) {
        candidates = fallback.candidates;
        match = fallbackMatch;
        participant = match.participant;
        selectedSession = null;
        matchStrategy = fallback.strategy;
        strategyNotes = fallback.notes;
        fallbackUsed = true;
      }
    }
  }

  const checklistKey = participant ? String(participant.participantId) + '|' + String(mapping.instrumentNumber) : '';
  const checklist = checklistKey ? datasets.checklistByParticipantInstrument[checklistKey] : null;
  let matchStatus = 'unmatched';
  let notes = match.reason || '';
  const canAutoMatch = matchStrategy !== 'recent_completed_cohort' || config.allowRecentCompletedAutoMatch;
  if (timestampCol === -1) notes = 'Timestamp column not found';
  if (!mapping.nameColumn || nameCol === -1) notes = 'Name column not found';
  else if (!sessions.length && !fallbackUsed) notes = strategyNotes || ('No DARTS session found for response date ' + (responseDate || '(blank)'));
  else if (!participant) notes = 'No participant candidate matched';
  else if (!checklist) notes = 'Matched participant, but checklist item was not found';
  else if (match.score >= config.autoMatchThreshold && canAutoMatch) matchStatus = 'auto_matched';
  else if (match.score >= config.reviewMatchThreshold) matchStatus = 'needs_review';
  else matchStatus = 'unmatched';
  if (strategyNotes) notes = notes ? notes + ' — ' + strategyNotes : strategyNotes;
  if (matchStrategy === 'recent_completed_cohort' && !config.allowRecentCompletedAutoMatch && matchStatus === 'needs_review') {
    notes += ' — Recent completed cohort fallback is configured for review-first matching.';
  }

  return {
    sourceWorkbookId: workbookId,
    sourceSheetName: sheet.getName(),
    sourceRowNumber: rowNumber,
    sourceTimestamp: responseTimestamp && !isNaN(new Date(responseTimestamp).getTime()) ? new Date(responseTimestamp).toISOString() : String(responseTimestamp || ''),
    responseHash: responseHash,
    rawName: rawName,
    normalizedName: normalizedName,
    instrumentNumber: mapping.instrumentNumber,
    instrumentName: mapping.instrumentName,
    candidateRolloutId: selectedSession ? selectedSession.rolloutId : (participant ? participant.rolloutId : ''),
    candidateSessionId: selectedSession ? selectedSession.sessionId : '',
    participantId: participant ? participant.participantId : '',
    participantName: participant ? participant.fullName : '',
    checklistId: checklist ? checklist.row[datasets.checklistSnapshot.headers.indexOf('checklistId')] : '',
    checklistRowIndex: checklist ? checklist.rowIndex : -1,
    matchStatus: matchStatus,
    confidenceScore: match.score || 0,
    matchedBy: matchStrategy,
    notes: notes,
    sourceLink: sourceLink
  };
}

function applyDigitalFormChecklistUpdates(matches, datasets, currentUserName) {
  const checklistSnapshot = datasets.checklistSnapshot;
  const cData = checklistSnapshot.data;
  const cHeaders = checklistSnapshot.headers;
  const statusCol = cHeaders.indexOf('status');
  const completedDateCol = cHeaders.indexOf('completedDate');
  const completedByCol = cHeaders.indexOf('completedBy');
  const notesCol = cHeaders.indexOf('notes');
  const dataLinkCol = cHeaders.indexOf('dataLink');
  const touchedParticipants = {};
  let updated = 0;
  matches.forEach(match => {
    if (match.matchStatus !== 'auto_matched' || match.checklistRowIndex < 1) return;
    const row = cData[match.checklistRowIndex];
    if (!row) return;
    if (String(row[statusCol] || '') === 'completed' && String(row[dataLinkCol] || '').indexOf(match.sourceLink) !== -1) {
      match.matchStatus = 'duplicate';
      return;
    }
    row[statusCol] = 'completed';
    row[completedDateCol] = match.sourceTimestamp || new Date().toISOString();
    row[completedByCol] = currentUserName || DIGITAL_FORM_SYNC_SYSTEM_USER;
    const existingNotes = String(row[notesCol] || '').trim();
    const syncNote = 'Auto-completed from Google Forms: ' + match.sourceSheetName + ' row ' + match.sourceRowNumber + ' (confidence ' + match.confidenceScore + '%).';
    row[notesCol] = existingNotes ? (existingNotes + '\n' + syncNote) : syncNote;
    if (dataLinkCol !== -1) row[dataLinkCol] = match.sourceLink;
    touchedParticipants[String(match.participantId)] = true;
    updated++;
  });
  if (updated > 0) {
    checklistSnapshot.sheet.getRange(1, 1, cData.length, cHeaders.length).setValues(cData);
    invalidateSheetSnapshot('Checklist');
    Object.keys(touchedParticipants).forEach(participantId => {
      updateParticipantCompletion(participantId);
      syncParticipantStatusFromChecklist(participantId);
    });
  }
  return updated;
}

function appendDigitalSyncLogs(matches, dryRun, actorName) {
  if (dryRun || !matches.length) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormResponseSyncLog');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
  const rows = matches.map(match => {
    const row = new Array(headers.length).fill('');
    const set = (header, value) => { const idx = headers.indexOf(header); if (idx !== -1) row[idx] = value; };
    set('syncId', generateUUID());
    set('sourceWorkbookId', match.sourceWorkbookId);
    set('sourceSheetName', match.sourceSheetName);
    set('sourceRowNumber', match.sourceRowNumber);
    set('sourceTimestamp', match.sourceTimestamp);
    set('responseHash', match.responseHash);
    set('rawName', match.rawName);
    set('normalizedName', match.normalizedName);
    set('instrumentNumber', match.instrumentNumber);
    set('instrumentName', match.instrumentName);
    set('candidateRolloutId', match.candidateRolloutId);
    set('candidateSessionId', match.candidateSessionId);
    set('participantId', match.participantId);
    set('checklistId', match.checklistId);
    set('matchStatus', match.matchStatus);
    set('confidenceScore', match.confidenceScore);
    set('matchedBy', match.matchedBy || actorName || DIGITAL_FORM_SYNC_SYSTEM_USER);
    set('matchedAt', new Date().toISOString());
    set('notes', match.notes);
    set('sourceLink', match.sourceLink);
    return row;
  });
  appendRowsAsPlainText(sheet, rows);
}

function appendDigitalReviewQueue(matches, dryRun) {
  if (dryRun) return 0;
  const reviewMatches = matches.filter(match => match.matchStatus === 'needs_review' || match.matchStatus === 'unmatched');
  if (!reviewMatches.length) return 0;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FormResponseMatchQueue');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
  const existing = {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    existing[String(data[i][headers.indexOf('sourceSheetName')]) + '::' + String(data[i][headers.indexOf('sourceRowNumber')]) + '::' + String(data[i][headers.indexOf('responseHash')])] = true;
  }
  const rows = [];
  reviewMatches.forEach(match => {
    const key = match.sourceSheetName + '::' + match.sourceRowNumber + '::' + match.responseHash;
    if (existing[key]) return;
    const row = new Array(headers.length).fill('');
    const set = (header, value) => { const idx = headers.indexOf(header); if (idx !== -1) row[idx] = value; };
    set('queueId', generateUUID());
    set('createdAt', new Date().toISOString());
    set('sourceWorkbookId', match.sourceWorkbookId);
    set('sourceSheetName', match.sourceSheetName);
    set('sourceRowNumber', match.sourceRowNumber);
    set('sourceTimestamp', match.sourceTimestamp);
    set('rawName', match.rawName);
    set('normalizedName', match.normalizedName);
    set('instrumentNumber', match.instrumentNumber);
    set('instrumentName', match.instrumentName);
    set('candidateRolloutId', match.candidateRolloutId);
    set('candidateSessionId', match.candidateSessionId);
    set('suggestedParticipantId', match.participantId);
    set('suggestedParticipantName', match.participantName);
    set('confidenceScore', match.confidenceScore);
    set('status', 'pending');
    set('notes', match.notes);
    set('sourceLink', match.sourceLink);
    set('responseHash', match.responseHash);
    rows.push(row);
  });
  appendRowsAsPlainText(sheet, rows);
  return rows.length;
}

function runDigitalFormSync(token, options) {
  const currentUser = token ? validateSession(token) : { fullName: DIGITAL_FORM_SYNC_SYSTEM_USER, role: 'admin', userId: DIGITAL_FORM_SYNC_SUPER_ADMIN_ID, username: DIGITAL_FORM_SYNC_SUPER_ADMIN_USERNAME };
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  ensureDigitalFormSyncSheets();
  const dryRun = !!(options && options.dryRun);
  const force = !!(options && options.force);
  const config = getDigitalFormSyncConfigInternal();
  const workbookConfigs = [
    { site: 'UGA', workbookId: config.masterWorkbookIds && config.masterWorkbookIds.UGA },
    { site: 'Missouri', workbookId: config.masterWorkbookIds && config.masterWorkbookIds.Missouri }
  ].filter(item => item.workbookId);
  if (!workbookConfigs.length && config.masterWorkbookId) workbookConfigs.push({ site: '', workbookId: config.masterWorkbookId });
  if (!workbookConfigs.length) return { success: false, message: 'At least one site master workbook is not configured' };
  const datasets = buildDigitalSyncDatasets();
  const existingKeys = getExistingDigitalSyncKeys();
  const watermarks = getDigitalFormSyncWatermarksInternal();
  const matches = [];
  const errors = [];
  workbookConfigs.forEach(workbookConfig => {
    const enabledMappings = getDigitalFormMappingsForSite(config, workbookConfig.site).filter(mapping => mapping.enabled);

    let workbook;
    try {
      workbook = openDigitalFormWorkbook(workbookConfig.workbookId);
    } catch (e) {
      errors.push((workbookConfig.site || 'Master') + ' workbook: ' + e.message);
      return;
    }
    enabledMappings.forEach(mapping => {
      const sheet = workbook.getSheetByName(mapping.sheetName);
      if (!sheet) { errors.push((workbookConfig.site || 'Master') + ' sheet not found: ' + mapping.sheetName); return; }
      if (sheet.getLastRow() < 2) return;
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
      const lastRow = sheet.getLastRow();
      const watermarkKey = (workbookConfig.site || workbookConfig.workbookId) + '::' + mapping.sheetName;
      const watermark = watermarks[watermarkKey] || {};
      const startRow = force || dryRun ? 2 : Math.max(2, (Number(watermark.lastProcessedRow) || 1) + 1);
      if (startRow > lastRow) return;
      const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, headers.length).getValues();
      values.forEach((row, offset) => {
        const rowNumber = startRow + offset;
        const analysis = analyzeDigitalFormResponse(mapping, sheet, headers, row, rowNumber, workbookConfig.workbookId, datasets, config, workbookConfig.site);
        analysis.sourceSite = workbookConfig.site || '';
        const existingKey = analysis.sourceWorkbookId + '::' + analysis.sourceSheetName + '::' + analysis.sourceRowNumber + '::' + analysis.responseHash;
        if (existingKeys[existingKey] && !force) return;
        matches.push(analysis);
      });
      if (!dryRun) {
        watermarks[watermarkKey] = { lastProcessedRow: lastRow, lastProcessedAt: new Date().toISOString() };
      }
    });
  });

  const updatedCount = dryRun ? 0 : applyDigitalFormChecklistUpdates(matches, datasets, currentUser.fullName || DIGITAL_FORM_SYNC_SYSTEM_USER);
  const queuedCount = appendDigitalReviewQueue(matches, dryRun);
  appendDigitalSyncLogs(matches, dryRun, currentUser.fullName || DIGITAL_FORM_SYNC_SYSTEM_USER);
  if (!dryRun) saveDigitalFormSyncWatermarksInternal(watermarks);

  return {
    success: true,
    dryRun: dryRun,
    processed: matches.length,
    autoMatched: matches.filter(match => match.matchStatus === 'auto_matched').length,
    needsReview: matches.filter(match => match.matchStatus === 'needs_review').length,
    unmatched: matches.filter(match => match.matchStatus === 'unmatched').length,
    duplicate: matches.filter(match => match.matchStatus === 'duplicate').length,
    checklistUpdated: updatedCount,
    queued: queuedCount,
    errors: errors,
    preview: matches.slice(0, 100)
  };
}

function runDigitalFormSyncDryRun(token) {
  return runDigitalFormSync(token, { dryRun: true, force: true });
}

function digitalFormSyncScheduledRun() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const config = getDigitalFormSyncConfigInternal();
    if (!config.enabled) return;
    runDigitalFormSync(null, { dryRun: false });
  } finally {
    lock.releaseLock();
  }
}

function getDigitalFormSyncTriggerSummary() {
  const triggers = ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === 'digitalFormSyncScheduledRun');
  return { enabled: triggers.length > 0, count: triggers.length };
}

function setDigitalFormSyncSchedule(token, enabled, everyMinutes) {
  const currentUser = validateSession(token);
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === 'digitalFormSyncScheduledRun') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  if (enabled) {
    const minutes = Math.max(5, Number(everyMinutes) || 10);
    ScriptApp.newTrigger('digitalFormSyncScheduledRun').timeBased().everyMinutes(minutes).create();
  }
  return { success: true, message: enabled ? 'Scheduled sync enabled' : 'Scheduled sync disabled', triggers: getDigitalFormSyncTriggerSummary() };
}

function saveParticipantAliasInternal(participantId, alias, createdBy, notes) {
  const normalized = normalizePersonNameForMatch(alias);
  if (!participantId || !normalized) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipantAliases');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(Boolean);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('participantId')]) === String(participantId) && String(data[i][headers.indexOf('normalizedAlias')]) === normalized) return;
  }
  const row = new Array(headers.length).fill('');
  const set = (header, value) => { const idx = headers.indexOf(header); if (idx !== -1) row[idx] = value; };
  set('aliasId', generateUUID());
  set('participantId', participantId);
  set('alias', alias);
  set('normalizedAlias', normalized);
  set('createdAt', new Date().toISOString());
  set('createdBy', createdBy || DIGITAL_FORM_SYNC_SYSTEM_USER);
  set('notes', notes || 'Created from Digital Forms Sync review');
  appendRowsAsPlainText(sheet, [row]);
}

function reviewDigitalFormMatch(token, queueId, action, participantId) {
  const currentUser = validateSession(token);
  if (!canAccessDigitalFormSync(currentUser)) return { success: false, message: 'Unauthorized' };
  ensureDigitalFormSyncSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('FormResponseMatchQueue');
  const queueData = queueSheet.getDataRange().getValues();
  const qHeaders = queueData[0] || [];
  let rowIndex = -1;
  let queueRow = null;
  for (let i = 1; i < queueData.length; i++) {
    if (String(queueData[i][qHeaders.indexOf('queueId')]) === String(queueId)) {
      rowIndex = i + 1;
      queueRow = queueData[i];
      break;
    }
  }
  if (!queueRow) return { success: false, message: 'Review item not found' };
  if (action === 'ignore') {
    queueSheet.getRange(rowIndex, qHeaders.indexOf('status') + 1).setValue('ignored');
    queueSheet.getRange(rowIndex, qHeaders.indexOf('reviewedBy') + 1).setValue(currentUser.fullName);
    queueSheet.getRange(rowIndex, qHeaders.indexOf('reviewedAt') + 1).setValue(new Date().toISOString());
    invalidateSheetSnapshot('FormResponseMatchQueue');
    return { success: true, message: 'Response ignored' };
  }
  const targetParticipantId = participantId || queueRow[qHeaders.indexOf('suggestedParticipantId')];
  if (!targetParticipantId) return { success: false, message: 'Participant ID is required to confirm this match' };
  const instrumentNumber = queueRow[qHeaders.indexOf('instrumentNumber')];
  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cHeaders = checklistSnapshot.headers;
  const cData = checklistSnapshot.data;
  const participantCol = cHeaders.indexOf('participantId');
  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  let checklistRowIndex = -1;
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][participantCol]) === String(targetParticipantId) && String(cData[i][instrumentCol]) === String(instrumentNumber)) {
      checklistRowIndex = i;
      break;
    }
  }
  if (checklistRowIndex === -1) return { success: false, message: 'Checklist item not found for selected participant' };
  const match = {
    matchStatus: 'auto_matched',
    checklistRowIndex: checklistRowIndex,
    participantId: targetParticipantId,
    sourceTimestamp: queueRow[qHeaders.indexOf('sourceTimestamp')],
    sourceSheetName: queueRow[qHeaders.indexOf('sourceSheetName')],
    sourceRowNumber: queueRow[qHeaders.indexOf('sourceRowNumber')],
    confidenceScore: queueRow[qHeaders.indexOf('confidenceScore')],
    sourceLink: queueRow[qHeaders.indexOf('sourceLink')]
  };
  applyDigitalFormChecklistUpdates([match], buildDigitalSyncDatasets(), currentUser.fullName);
  queueSheet.getRange(rowIndex, qHeaders.indexOf('status') + 1).setValue('resolved');
  queueSheet.getRange(rowIndex, qHeaders.indexOf('reviewedBy') + 1).setValue(currentUser.fullName);
  queueSheet.getRange(rowIndex, qHeaders.indexOf('reviewedAt') + 1).setValue(new Date().toISOString());
  invalidateSheetSnapshot('FormResponseMatchQueue');
  if (getDigitalFormSyncConfigInternal().saveAliasesOnReview) {
    saveParticipantAliasInternal(targetParticipantId, queueRow[qHeaders.indexOf('rawName')], currentUser.fullName, 'Confirmed from form response review');
  }
  return { success: true, message: 'Match confirmed and checklist updated' };
}

// ============================================
// DATABASE INITIALIZATION
// ============================================

/**
 * Initialize all database sheets
 * Run this once when setting up the application
 */
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create Users sheet
  createSheetIfNotExists(ss, 'Users', [
    'userId', 'username', 'passwordHash', 'fullName', 'role', 'site', 'status', 'createdAt', 'createdBy', 'lastLogin'
  ]);

  // Create StudyRollouts sheet
  createSheetIfNotExists(ss, 'StudyRollouts', [
    'rolloutId', 'site', 'schoolName', 'period', 'year', 'status', 'createdAt', 'createdBy', 'description'
  ]);

  // Create Participants sheet
  createSheetIfNotExists(ss, 'Participants', [
    'participantId', 'fullName', 'parent_guardian_names', 'parent_guardian_phone',
    'parent_guardian_address', 'parent_guardian_email', 'parent_guardian_dob',
    'site', 'rolloutId', 'schoolName', 'period', 'year',
    'enrollmentDate', 'enrolledBy', 'status', 'notes', 'completionPercentage'
  ]);

  // Create Checklist sheet
  createSheetIfNotExists(ss, 'Checklist', [
    'checklistId', 'participantId', 'instrumentNumber', 'instrumentName', 'category',
    'status', 'completedDate', 'completedBy', 'notes', 'dataLink'
  ]);

  // Create StudySessions sheet
  createSheetIfNotExists(ss, 'StudySessions', [
    'sessionId', 'rolloutId', 'sessionNumber', 'sessionDate', 'sessionName',
    'status', 'createdAt', 'createdBy'
  ]);
  const sessionsSheet = ss.getSheetByName('StudySessions');
  if (sessionsSheet) {
    const sessionDateColumn = 4;
    sessionsSheet.getRange(2, sessionDateColumn, sessionsSheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }

  // Create SessionAttendance sheet
  createSheetIfNotExists(ss, 'SessionAttendance', [
    'attendanceId', 'sessionId', 'participantId', 'status', 'markedAt', 'markedBy', 'notes'
  ]);

  ensureMediaRecordSheets();

  // Create ActivityLog sheet
  createSheetIfNotExists(ss, 'ActivityLog', [
    'logId', 'timestamp', 'userId', 'userName', 'action', 'targetType', 'targetId', 'details'
  ]);

  // Create Sessions sheet (for authentication)
  createSheetIfNotExists(ss, 'Sessions', [
    'sessionId', 'userId', 'token', 'createdAt', 'expiresAt', 'isActive'
  ]);

  // Create Config sheet
  createSheetIfNotExists(ss, 'Config', [
    'key', 'value', 'description', 'updatedAt'
  ]);

  ensureDigitalFormSyncSheets();

  Logger.log('Database initialized successfully!');
  return { success: true, message: 'Database initialized successfully!' };
}

/**
 * Create a sheet if it doesn't exist
 */
function createSheetIfNotExists(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('Created sheet: ' + sheetName);
  }

  return sheet;
}

/**
 * Create the initial admin user
 * Run this once after initializing the database
 */
function createInitialAdmin() {
  const username = 'admin';
  const password = generateRandomPassword();
  const passwordHash = hashPassword(password);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) {
    throw new Error('Users sheet not found. Run initializeDatabase() first.');
  }

  // Check if admin already exists
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      Logger.log('Admin user already exists!');
      return { success: false, message: 'Admin user already exists' };
    }
  }

  const userId = generateUUID();
  const timestamp = new Date().toISOString();

  usersSheet.appendRow([
    userId,
    username,
    passwordHash,
    'System Administrator',
    'admin',
    'All',
    'active',
    timestamp,
    'system',
    ''
  ]);

  Logger.log('Initial admin created!');
  Logger.log('Username: ' + username);
  Logger.log('Temporary Password: ' + password);
  Logger.log('IMPORTANT: Change this password after first login!');

  return {
    success: true,
    message: 'Admin created successfully',
    username: username,
    password: password
  };
}

/**
 * UTILITY: Generate a password hash
 * Run this function from the script editor to generate hashed passwords
 * for manual entry in the spreadsheet
 *
 * Usage: Change the password below and run this function.
 * Copy the hash from the execution log to your spreadsheet.
 */
function generatePasswordHash() {
  const password = 'admin'; // <-- Change this to your desired password
  const hash = hashPassword(password);
  Logger.log('Password: ' + password);
  Logger.log('Hash: ' + hash);
  Logger.log('Copy the hash above to the passwordHash column in the Users sheet');
  return hash;
}

// ============================================
// AUTHENTICATION FUNCTIONS
// ============================================

/**
 * Authenticate user with username and password
 */
function authenticateUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) {
    return { success: false, message: 'System not initialized' };
  }

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  // Debug logging
  Logger.log('Attempting login for username: ' + username);
  Logger.log('Headers found: ' + JSON.stringify(headers));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUsername = row[headers.indexOf('username')];
    const rowPasswordHash = row[headers.indexOf('passwordHash')];
    const rowStatus = row[headers.indexOf('status')];

    Logger.log('Checking row ' + i + ': username=' + rowUsername + ', status=' + rowStatus);

    if (rowUsername === username && rowStatus === 'active') {
      Logger.log('Username matched! Verifying password...');
      Logger.log('Stored hash: ' + rowPasswordHash);
      Logger.log('Input password hash: ' + hashPassword(password));

      if (verifyPassword(password, rowPasswordHash)) {
        const userId = row[headers.indexOf('userId')];

        // Update last login
        const lastLoginCol = headers.indexOf('lastLogin') + 1;
        usersSheet.getRange(i + 1, lastLoginCol).setValue(new Date().toISOString());

        // Log activity
        logActivity(userId, row[headers.indexOf('fullName')], 'LOGIN', 'user', userId, 'User logged in');

        return {
          success: true,
          token: rowUsername,
          user: {
            userId: userId,
            username: rowUsername,
            fullName: row[headers.indexOf('fullName')],
            role: row[headers.indexOf('role')],
            site: row[headers.indexOf('site')]
          }
        };
      } else {
        Logger.log('Password verification failed');
      }
    }
  }

  return { success: false, message: 'Invalid username or password' };
}

/**
 * Create a new session for a user
 */
function createSession(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sessionsSheet = ss.getSheetByName('Sessions');
  if (!sessionsSheet) {
    sessionsSheet = createSheetIfNotExists(ss, 'Sessions', [
      'sessionId', 'userId', 'token', 'createdAt', 'expiresAt', 'isActive'
    ]);
  }

  const sessionId = generateUUID();
  const token = generateSessionToken();
  const createdAt = new Date();
  const expiresAt = new Date(createdAt.getTime() + 60 * 60 * 1000); // 1 hour

  sessionsSheet.appendRow([
    sessionId,
    userId,
    token,
    createdAt.toISOString(),
    expiresAt.toISOString(),
    true
  ]);

  return { sessionId, token, expiresAt };
}

/**
 * Validate a session token
 */
function validateSession(token) {
  if (!token) return null;

  return getUserByUsername(token);
}

/**
 * Get current user from session token
 */
function getCurrentUser(token) {
  return validateSession(token);
}

/**
 * Logout user (invalidate session)
 */
function logoutUser(token) {
  return { success: true };
}

/**
 * Change user password
 */
function changePassword(token, currentPassword, newPassword) {
  const user = validateSession(token);
  if (!user) {
    return { success: false, message: 'Invalid session' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === user.userId) {
      const storedPassword = data[i][headers.indexOf('passwordHash')];

      if (!verifyPassword(currentPassword, storedPassword)) {
        return { success: false, message: 'Current password is incorrect' };
      }

      // Store plain text password since the Google Sheet is highly protected
      usersSheet.getRange(i + 1, headers.indexOf('passwordHash') + 1).setValue(newPassword);

      logActivity(user.userId, user.fullName, 'PASSWORD_CHANGE', 'user', user.userId, 'Password changed');

      return { success: true, message: 'Password changed successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

// ============================================
// USER MANAGEMENT FUNCTIONS
// ============================================

/**
 * Get user by ID
 */
function getUserById(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) return null;

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      return {
        userId: data[i][headers.indexOf('userId')],
        username: data[i][headers.indexOf('username')],
        fullName: data[i][headers.indexOf('fullName')],
        role: data[i][headers.indexOf('role')],
        site: data[i][headers.indexOf('site')],
        status: data[i][headers.indexOf('status')],
        createdAt: data[i][headers.indexOf('createdAt')],
        lastLogin: data[i][headers.indexOf('lastLogin')]
      };
    }
  }

  return null;
}

/**
 * Get user by username
 */
function getUserByUsername(username) {
  if (!username) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) return null;

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('username')] === username &&
        data[i][headers.indexOf('status')] === 'active') {
      return {
        userId: data[i][headers.indexOf('userId')],
        username: data[i][headers.indexOf('username')],
        fullName: data[i][headers.indexOf('fullName')],
        role: data[i][headers.indexOf('role')],
        site: data[i][headers.indexOf('site')],
        status: data[i][headers.indexOf('status')],
        createdAt: data[i][headers.indexOf('createdAt')],
        lastLogin: data[i][headers.indexOf('lastLogin')]
      };
    }
  }

  return null;
}

/**
 * Get all users (Admin only)
 */
function getAllUsers(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) return { success: false, message: 'Users sheet not found' };

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      userId: data[i][headers.indexOf('userId')],
      username: data[i][headers.indexOf('username')],
      fullName: data[i][headers.indexOf('fullName')],
      role: data[i][headers.indexOf('role')],
      site: data[i][headers.indexOf('site')],
      status: data[i][headers.indexOf('status')],
      createdAt: data[i][headers.indexOf('createdAt')],
      lastLogin: data[i][headers.indexOf('lastLogin')]
    });
  }

  return { success: true, users: users };
}

/**
 * Create a new user (Admin only)
 */
function createUser(token, userData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  // Basic validation
  if (!userData.role || !CONFIG.ROLES.includes(userData.role)) {
    return { success: false, message: 'Invalid role selected' };
  }

  if (!userData.site) {
    return { success: false, message: 'Site is required' };
  }

  const validSites = CONFIG.SITES.concat(['All']);
  if (validSites.indexOf(userData.site) === -1) {
    return { success: false, message: 'Invalid site selected' };
  }

  if (userData.role === 'facilitator' && userData.site === 'All') {
    return { success: false, message: 'Facilitators must be assigned to a single site' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');

  // Check if username already exists
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('username')] === userData.username) {
      return { success: false, message: 'Username already exists' };
    }
  }

  const userId = generateUUID();
  // Use provided password or generate one if not provided
  const password = userData.password || generateRandomPassword();
  const timestamp = new Date().toISOString();

  usersSheet.appendRow([
    userId,
    userData.username,
    password,  // Plain text password
    userData.fullName,
    userData.role,
    userData.site,
    'active',
    timestamp,
    currentUser.userId,
    ''
  ]);

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_USER', 'user', userId,
    'Created user: ' + userData.username);

  return {
    success: true,
    message: 'User created successfully',
    tempPassword: password,
    userId: userId
  };
}

/**
 * Update user (Admin only)
 */
function updateUser(token, userId, userData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      const existingRole = data[i][headers.indexOf('role')];
      const nextRole = userData.role || existingRole;

      if (userData.fullName) {
        usersSheet.getRange(i + 1, headers.indexOf('fullName') + 1).setValue(userData.fullName);
      }
      if (userData.role) {
        if (!CONFIG.ROLES.includes(userData.role)) {
          return { success: false, message: 'Invalid role selected' };
        }
        usersSheet.getRange(i + 1, headers.indexOf('role') + 1).setValue(userData.role);
      }
      if (userData.site) {
        const validSites = CONFIG.SITES.concat(['All']);
        if (validSites.indexOf(userData.site) === -1) {
          return { success: false, message: 'Invalid site selected' };
        }
        if (nextRole === 'facilitator' && userData.site === 'All') {
          return { success: false, message: 'Facilitators must be assigned to a single site' };
        }
        usersSheet.getRange(i + 1, headers.indexOf('site') + 1).setValue(userData.site);
      }
      if (userData.status) {
        usersSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(userData.status);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_USER', 'user', userId,
        'Updated user: ' + data[i][headers.indexOf('username')]);

      return { success: true, message: 'User updated successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

/**
 * Reset user password (Admin only)
 */
function resetUserPassword(token, userId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('userId')] === userId) {
      const newPassword = generateRandomPassword();
      // Store plain text password since the Google Sheet is highly protected
      usersSheet.getRange(i + 1, headers.indexOf('passwordHash') + 1).setValue(newPassword);

      logActivity(currentUser.userId, currentUser.fullName, 'RESET_PASSWORD', 'user', userId,
        'Reset password for: ' + data[i][headers.indexOf('username')]);

      return { success: true, tempPassword: newPassword };
    }
  }

  return { success: false, message: 'User not found' };
}

/**
 * Delete a user account (Admin only)
 */
function deleteUser(token, userId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  if (currentUser.userId === userId) {
    return { success: false, message: 'You cannot delete your own account' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName('Users');
  const sessionsSheet = ss.getSheetByName('Sessions');

  if (!usersSheet) {
    return { success: false, message: 'Users sheet not found' };
  }

  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const userIdCol = headers.indexOf('userId');
  const usernameCol = headers.indexOf('username');
  const roleCol = headers.indexOf('role');

  // Ensure at least one admin remains
  let adminCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][roleCol] === 'admin') {
      adminCount++;
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdCol] === userId) {
      const username = data[i][usernameCol];
      const role = data[i][roleCol];

      if (role === 'admin' && adminCount <= 1) {
        return { success: false, message: 'At least one admin user is required' };
      }

      usersSheet.deleteRow(i + 1);

      // Deactivate all sessions for this user
      if (sessionsSheet) {
        const sessionData = sessionsSheet.getDataRange().getValues();
        const sessionHeaders = sessionData[0];
        const sUserIdCol = sessionHeaders.indexOf('userId');
        const sIsActiveCol = sessionHeaders.indexOf('isActive');

        for (let j = 1; j < sessionData.length; j++) {
          if (sessionData[j][sUserIdCol] === userId) {
            sessionsSheet.getRange(j + 1, sIsActiveCol + 1).setValue(false);
          }
        }
      }

      logActivity(currentUser.userId, currentUser.fullName, 'DELETE_USER', 'user', userId,
        'Deleted user: ' + username);

      return { success: true, message: 'User deleted successfully' };
    }
  }

  return { success: false, message: 'User not found' };
}

// ============================================
// STUDY ROLLOUT FUNCTIONS
// ============================================

/**
 * Get all study cohorts
 */
function getAllRollouts(token, siteFilter) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return { success: false, message: 'Study cohorts sheet not found' };

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const cohorts = [];

  for (let i = 1; i < data.length; i++) {
    const cohort = {
      rolloutId: data[i][headers.indexOf('rolloutId')],
      site: data[i][headers.indexOf('site')],
      schoolName: data[i][headers.indexOf('schoolName')],
      period: data[i][headers.indexOf('period')],
      year: data[i][headers.indexOf('year')],
      status: data[i][headers.indexOf('status')],
      createdAt: data[i][headers.indexOf('createdAt')],
      description: data[i][headers.indexOf('description')]
    };

    // Apply site filter if specified
    if (!siteFilter || siteFilter === 'All' || cohort.site === siteFilter) {
      cohorts.push(cohort);
    }
  }

  return { success: true, cohorts: cohorts };
}

/**
 * Get cohorts by site
 */
function getRolloutsBySite(token, site) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && currentUser.site !== site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  return getAllRollouts(token, site);
}

/**
 * Create a new study cohort (Admin only)
 */
function createRollout(token, rolloutData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized - Admin access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  const rolloutId = generateUUID();
  const timestamp = new Date().toISOString();

  rolloutsSheet.appendRow([
    rolloutId,
    rolloutData.site,
    rolloutData.schoolName,
    rolloutData.period,
    rolloutData.year,
    'active',
    timestamp,
    currentUser.userId,
    rolloutData.description || ''
  ]);

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_ROLLOUT', 'cohort', rolloutId,
    'Created cohort: ' + rolloutData.schoolName + ' (' + rolloutData.period + ' ' + rolloutData.year + ')');

  return { success: true, message: 'Cohort created successfully', rolloutId: rolloutId };
}

/**
 * Update study cohort (Admin only)
 */
function updateRollout(token, rolloutId, rolloutData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('rolloutId')] === rolloutId) {
      if (rolloutData.schoolName) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('schoolName') + 1).setValue(rolloutData.schoolName);
      }
      if (rolloutData.period) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('period') + 1).setValue(rolloutData.period);
      }
      if (rolloutData.year) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('year') + 1).setValue(rolloutData.year);
      }
      if (rolloutData.status) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(rolloutData.status);
      }
      if (rolloutData.description !== undefined) {
        rolloutsSheet.getRange(i + 1, headers.indexOf('description') + 1).setValue(rolloutData.description);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_ROLLOUT', 'cohort', rolloutId,
        'Updated cohort');

      return { success: true, message: 'Cohort updated successfully' };
    }
  }

  return { success: false, message: 'Cohort not found' };
}

function getProtocolItems(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  const items = rolloutId ? getRolloutProtocolItemsInternal(rolloutId) : getGlobalProtocolItems();
  return { success: true, items: items };
}

function saveGlobalProtocolItems(token, items) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  if (!Array.isArray(items) || items.length === 0) return { success: false, message: 'At least one protocol item is required' };
  const normalized = items.map((item, idx) => ({
    number: idx + 1,
    name: String(item.name || '').trim(),
    category: String(item.category || 'lesson').trim()
  })).filter(i => i.name);
  if (!normalized.length) return { success: false, message: 'Invalid protocol list' };
  upsertConfigValue('PROTOCOL_ITEMS', JSON.stringify(normalized), 'Global protocol items for new checklist assignments');
  return { success: true, message: 'Global protocol items updated', items: normalized };
}

function saveRolloutProtocolItems(token, rolloutId, itemNumbers) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') return { success: false, message: 'Unauthorized' };
  if (!rolloutId) return { success: false, message: 'rolloutId required' };
  const globalItems = getGlobalProtocolItems();
  const globalMap = {};
  globalItems.forEach(i => globalMap[String(i.number)] = true);
  const filtered = (itemNumbers || []).map(n => String(n)).filter(n => globalMap[n]);
  if (filtered.length === 0) return { success: false, message: 'At least one protocol item must remain enabled for a rollout' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valueCol = headers.indexOf('value');
  let overrides = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === 'ROLLOUT_PROTOCOL_OVERRIDES') {
      try { overrides = JSON.parse(data[i][valueCol] || '{}') || {}; } catch (e) {}
    }
  }
  overrides[rolloutId] = filtered;
  upsertConfigValue('ROLLOUT_PROTOCOL_OVERRIDES', JSON.stringify(overrides), 'Per-rollout enabled protocol item numbers');
  const syncSummary = syncRolloutChecklistProtocolItems(rolloutId, filtered);
  return { success: true, message: 'Rollout protocol items updated', sync: syncSummary };
}

function syncRolloutChecklistProtocolItems(rolloutId, enabledNumbers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  if (!participantsSheet || !checklistSheet) return { added: 0, deleted: 0 };

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cHeaders = checklistSnapshot.headers;
  const cData = checklistSnapshot.data;
  const enabledMap = {};
  enabledNumbers.forEach(n => enabledMap[String(n)] = true);
  const enabledItems = getGlobalProtocolItems().filter(i => enabledMap[String(i.number)]);

  const rolloutParticipantIds = {};
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][pHeaders.indexOf('rolloutId')]) === String(rolloutId)) {
      rolloutParticipantIds[String(pData[i][pHeaders.indexOf('participantId')])] = true;
    }
  }

  let deleted = 0;
  for (let i = cData.length - 1; i >= 1; i--) {
    const participantId = String(cData[i][cHeaders.indexOf('participantId')]);
    if (!rolloutParticipantIds[participantId]) continue;
    const number = String(cData[i][cHeaders.indexOf('instrumentNumber')]);
    if (!enabledMap[number]) {
      checklistSheet.deleteRow(i + 1);
      invalidateSheetSnapshot('Checklist');
      deleted++;
    }
  }

  const freshData = checklistSheet.getRange(1, 1, checklistSheet.getLastRow(), cHeaders.length).getValues();
  const existing = {};
  for (let i = 1; i < freshData.length; i++) {
    const participantId = String(freshData[i][cHeaders.indexOf('participantId')]);
    const number = String(freshData[i][cHeaders.indexOf('instrumentNumber')]);
    existing[participantId + '|' + number] = true;
  }

  let added = 0;
  const rowsToAppend = [];
  Object.keys(rolloutParticipantIds).forEach(participantId => {
    enabledItems.forEach(item => {
      const key = participantId + '|' + String(item.number);
      if (!existing[key]) {
        const row = new Array(cHeaders.length).fill('');
        const setChecklistValue = (header, value) => {
          const idx = cHeaders.indexOf(header);
          if (idx !== -1) row[idx] = value;
        };
        setChecklistValue('checklistId', generateUUID());
        setChecklistValue('participantId', participantId);
        setChecklistValue('instrumentNumber', item.number);
        setChecklistValue('instrumentName', item.name);
        setChecklistValue('category', item.category);
        setChecklistValue('status', 'not_started');
        rowsToAppend.push(row);
        added++;
      }
    });
  });
  appendRowsAsPlainText(checklistSheet, rowsToAppend);
  Object.keys(rolloutParticipantIds).forEach(participantId => updateParticipantCompletion(participantId));

  return { added: added, deleted: deleted };
}

/**
 * Get cohort by ID
 */
function getRolloutById(rolloutId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return null;

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('rolloutId')] === rolloutId) {
      return {
        rolloutId: data[i][headers.indexOf('rolloutId')],
        site: data[i][headers.indexOf('site')],
        schoolName: data[i][headers.indexOf('schoolName')],
        period: data[i][headers.indexOf('period')],
        year: data[i][headers.indexOf('year')],
        status: data[i][headers.indexOf('status')],
        description: data[i][headers.indexOf('description')]
      };
    }
  }

  return null;
}

/**
 * Delete a cohort and all related data (Admin only)
 * This will cascade delete:
 * - All participants in the cohort
 * - All checklist items for those participants
 * - All sessions for the cohort
 * - All attendance records for those sessions
 */
function deleteRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized - Admin access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const attendanceSheet = ss.getSheetByName('SessionAttendance');

  // Verify all sheets exist
  if (!rolloutsSheet || !participantsSheet || !checklistSheet || !sessionsSheet || !attendanceSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Get all participants in this cohort
  const participantData = participantsSheet.getDataRange().getValues();
  const participantHeaders = participantData[0];
  const participantIds = [];

  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][participantHeaders.indexOf('rolloutId')] === rolloutId) {
      participantIds.push(participantData[i][participantHeaders.indexOf('participantId')]);
    }
  }

  // Get all sessions in this cohort
  const sessionData = sessionsSheet.getDataRange().getValues();
  const sessionHeaders = sessionData[0];
  const sessionIds = [];

  for (let i = 1; i < sessionData.length; i++) {
    if (sessionData[i][sessionHeaders.indexOf('rolloutId')] === rolloutId) {
      sessionIds.push(sessionData[i][sessionHeaders.indexOf('sessionId')]);
    }
  }

  // COUNT items before deletion for logging
  let deletedChecklistCount = 0;
  let deletedAttendanceCount = 0;

  // DELETE ORDER (from most dependent to least dependent):

  // 1. Delete all attendance records for the sessions
  const attendanceData = attendanceSheet.getDataRange().getValues();
  const attendanceHeaders = attendanceData[0];

  for (let i = attendanceData.length - 1; i >= 1; i--) {
    if (sessionIds.includes(attendanceData[i][attendanceHeaders.indexOf('sessionId')])) {
      attendanceSheet.deleteRow(i + 1);
      deletedAttendanceCount++;
    }
  }

  // 2. Delete all sessions for this cohort
  for (let i = sessionData.length - 1; i >= 1; i--) {
    if (sessionData[i][sessionHeaders.indexOf('rolloutId')] === rolloutId) {
      sessionsSheet.deleteRow(i + 1);
    }
  }

  // 3. Delete all checklist items for participants in this cohort
  // IMPORTANT: Must get fresh data after any deletions
  const checklistData = checklistSheet.getDataRange().getValues();
  const checklistHeaders = checklistData[0];

  Logger.log('Deleting checklists for participantIds: ' + participantIds.join(', '));
  Logger.log('Total checklist rows: ' + (checklistData.length - 1));

  for (let i = checklistData.length - 1; i >= 1; i--) {
    const rowParticipantId = checklistData[i][checklistHeaders.indexOf('participantId')];
    if (participantIds.includes(rowParticipantId)) {
      Logger.log('Deleting checklist row ' + (i + 1) + ' for participant ' + rowParticipantId);
      checklistSheet.deleteRow(i + 1);
      deletedChecklistCount++;
    }
  }

  Logger.log('Deleted ' + deletedChecklistCount + ' checklist items');

  // 4. Delete all participants in this cohort
  for (let i = participantData.length - 1; i >= 1; i--) {
    if (participantData[i][participantHeaders.indexOf('rolloutId')] === rolloutId) {
      participantsSheet.deleteRow(i + 1);
    }
  }

  // 5. Delete the cohort itself
  const rolloutData = rolloutsSheet.getDataRange().getValues();
  const rolloutHeaders = rolloutData[0];

  for (let i = 1; i < rolloutData.length; i++) {
    if (rolloutData[i][rolloutHeaders.indexOf('rolloutId')] === rolloutId) {
      rolloutsSheet.deleteRow(i + 1);
      break;
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'DELETE_ROLLOUT', 'cohort', rolloutId,
    'Deleted cohort and all related data (' + participantIds.length + ' participants, ' +
    deletedChecklistCount + ' checklist items, ' + sessionIds.length + ' sessions, ' +
    deletedAttendanceCount + ' attendance records)');

  return {
    success: true,
    message: 'Cohort deleted successfully',
    deletedParticipants: participantIds.length,
    deletedChecklistItems: deletedChecklistCount,
    deletedSessions: sessionIds.length,
    deletedAttendance: deletedAttendanceCount
  };
}

// ============================================
// STUDY SESSION FUNCTIONS
// ============================================

/**
 * Get all sessions for a specific cohort
 */
function getSessionsByRollout(token, rolloutId) {
  try {
    const currentUser = validateSession(token);
    if (!currentUser) {
      return { success: false, message: 'Unauthorized' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sessionsSheet = ss.getSheetByName('StudySessions');

    if (!sessionsSheet) {
      return { success: false, message: 'StudySessions sheet not found. Please run initializeDatabase() from the Apps Script editor to create required sheets.' };
    }

    const data = sessionsSheet.getDataRange().getValues();
    const headers = data[0];
    const sessions = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('rolloutId')] === rolloutId) {
        sessions.push({
          sessionId: data[i][headers.indexOf('sessionId')],
          rolloutId: data[i][headers.indexOf('rolloutId')],
          sessionNumber: data[i][headers.indexOf('sessionNumber')],
          sessionDate: normalizeSessionDateValue(data[i][headers.indexOf('sessionDate')]),
          sessionName: data[i][headers.indexOf('sessionName')],
          status: data[i][headers.indexOf('status')],
          createdAt: data[i][headers.indexOf('createdAt')],
          createdBy: data[i][headers.indexOf('createdBy')],
          goproLink: headers.indexOf('goproLink') !== -1 ? data[i][headers.indexOf('goproLink')] : '',
          tascamLink: headers.indexOf('tascamLink') !== -1 ? data[i][headers.indexOf('tascamLink')] : '',
          meetingOwlLink: headers.indexOf('meetingOwlLink') !== -1 ? data[i][headers.indexOf('meetingOwlLink')] : '',
          recordingLinksJson: headers.indexOf('recordingLinksJson') !== -1 ? data[i][headers.indexOf('recordingLinksJson')] : '',
          fieldNotesLink: headers.indexOf('fieldNotesLink') !== -1 ? data[i][headers.indexOf('fieldNotesLink')] : ''
        });
      }
    }

    // Sort by session number
    sessions.sort((a, b) => a.sessionNumber - b.sessionNumber);

    return { success: true, sessions: sessions };
  } catch (error) {
    Logger.log('getSessionsByRollout ERROR: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return { success: false, message: 'Error loading sessions: ' + error.toString() };
  }
}

function buildStudySessionRow(headers, data) {
  const values = {
    sessionId: data.sessionId || '',
    rolloutId: data.rolloutId || '',
    sessionNumber: data.sessionNumber || '',
    sessionDate: data.sessionDate || '',
    sessionName: data.sessionName || '',
    status: data.status || 'scheduled',
    createdAt: data.createdAt || '',
    createdBy: data.createdBy || '',
    goproLink: data.goproLink || '',
    tascamLink: data.tascamLink || '',
    meetingOwlLink: data.meetingOwlLink || '',
    recordingLinksJson: data.recordingLinksJson || ''
  };

  return headers.map(header => {
    const key = String(header || '').trim();
    return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : '';
  });
}

/**
 * Create a new study session
 */
function createSession(token, sessionData) {
  const currentUser = validateSession(token);
  if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'facilitator')) {
    return { success: false, message: 'Unauthorized - Admin or Facilitator access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const headers = sessionsSheet.getRange(1, 1, 1, sessionsSheet.getLastColumn()).getValues()[0];
  const sessionDateCol = headers.indexOf('sessionDate') + 1;

  const sessionId = generateUUID();
  const timestamp = new Date().toISOString();
  const sessionDateText = normalizeSessionDateValue(sessionData.sessionDate);

  const nextRow = sessionsSheet.getLastRow() + 1;
  const row = buildStudySessionRow(headers, {
    sessionId: sessionId,
    rolloutId: sessionData.rolloutId,
    sessionNumber: sessionData.sessionNumber,
    sessionDate: sessionDateText,
    sessionName: sessionData.sessionName || '',
    status: 'scheduled',
    createdAt: timestamp,
    createdBy: currentUser.userId
  });
  sessionsSheet.getRange(nextRow, 1, 1, headers.length).setValues([row]);
  if (sessionDateCol > 0) {
    setPlainTextCell(sessionsSheet, nextRow, sessionDateCol, sessionDateText);
  }

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_SESSION', 'session', sessionId,
    'Created session #' + sessionData.sessionNumber + ' for cohort ' + sessionData.rolloutId);

  return { success: true, message: 'Session created successfully', sessionId: sessionId };
}

/**
 * Batch create multiple sessions for a cohort
 */
function batchCreateSessions(token, rolloutId, sessionsData) {
  const currentUser = validateSession(token);
  if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'facilitator')) {
    return { success: false, message: 'Unauthorized - Admin or Facilitator access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const headers = sessionsSheet.getRange(1, 1, 1, sessionsSheet.getLastColumn()).getValues()[0];
  const sessionDateCol = headers.indexOf('sessionDate') + 1;
  const timestamp = new Date().toISOString();
  const createdSessions = [];

  const startRow = sessionsSheet.getLastRow() + 1;
  const rows = sessionsData.map(sessionData => {
    const sessionId = generateUUID();
    const sessionDateText = normalizeSessionDateValue(sessionData.sessionDate);
    createdSessions.push({
      sessionId: sessionId,
      sessionNumber: sessionData.sessionNumber
    });
    return buildStudySessionRow(headers, {
      sessionId: sessionId,
      rolloutId: rolloutId,
      sessionNumber: sessionData.sessionNumber,
      sessionDate: sessionDateText,
      sessionName: sessionData.sessionName || '',
      status: 'scheduled',
      createdAt: timestamp,
      createdBy: currentUser.userId
    });
  });

  if (rows.length > 0) {
    sessionsSheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
    if (sessionDateCol > 0) {
      const dateRange = sessionsSheet.getRange(startRow, sessionDateCol, rows.length, 1);
      dateRange.setNumberFormat('@');
      dateRange.setValues(rows.map(row => [row[sessionDateCol - 1]]));
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'CREATE_SESSION', 'cohort', rolloutId,
    'Created ' + sessionsData.length + ' sessions for cohort');

  return {
    success: true,
    message: sessionsData.length + ' sessions created successfully',
    sessions: createdSessions
  };
}

/**
 * Update a study session
 */
function updateSession(token, sessionId, sessionData) {
  const currentUser = validateSession(token);
  if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'facilitator')) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const data = sessionsSheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('sessionId')] === sessionId) {
      if (sessionData.sessionDate) {
        const sessionDateText = normalizeSessionDateValue(sessionData.sessionDate);
        setPlainTextCell(
          sessionsSheet,
          i + 1,
          headers.indexOf('sessionDate') + 1,
          sessionDateText
        );
      }
      if (sessionData.sessionName !== undefined) {
        sessionsSheet.getRange(i + 1, headers.indexOf('sessionName') + 1).setValue(sessionData.sessionName);
      }
      if (sessionData.status) {
        sessionsSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(sessionData.status);
      }

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_SESSION', 'session', sessionId,
        'Updated session');

      return { success: true, message: 'Session updated successfully' };
    }
  }

  return { success: false, message: 'Session not found' };
}

/**
 * Delete a study session
 */
function deleteSession(token, sessionId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized - Admin access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const attendanceSheet = ss.getSheetByName('SessionAttendance');

  // Delete the session
  const sessionData = sessionsSheet.getDataRange().getValues();
  const sessionHeaders = sessionData[0];

  for (let i = 1; i < sessionData.length; i++) {
    if (sessionData[i][sessionHeaders.indexOf('sessionId')] === sessionId) {
      sessionsSheet.deleteRow(i + 1);
      break;
    }
  }

  // Delete all attendance records for this session
  const attendanceData = attendanceSheet.getDataRange().getValues();
  const attendanceHeaders = attendanceData[0];

  for (let i = attendanceData.length - 1; i >= 1; i--) {
    if (attendanceData[i][attendanceHeaders.indexOf('sessionId')] === sessionId) {
      attendanceSheet.deleteRow(i + 1);
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'DELETE_SESSION', 'session', sessionId,
    'Deleted session and related attendance records');

  return { success: true, message: 'Session deleted successfully' };
}

// ============================================
// ATTENDANCE FUNCTIONS
// ============================================

/**
 * Mark attendance for a participant at a session
 */
function markAttendance(token, attendanceData) {
  const currentUser = validateSession(token);
  if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'facilitator')) {
    return { success: false, message: 'Unauthorized - Admin or Facilitator access required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('SessionAttendance');
  const data = attendanceSheet.getDataRange().getValues();
  const headers = data[0];

  // Check if attendance already exists
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('sessionId')] === attendanceData.sessionId &&
        data[i][headers.indexOf('participantId')] === attendanceData.participantId) {
      existingRow = i;
      break;
    }
  }

  const timestamp = new Date().toISOString();

  if (existingRow > 0) {
    // Update existing attendance
    attendanceSheet.getRange(existingRow + 1, headers.indexOf('status') + 1).setValue(attendanceData.status);
    attendanceSheet.getRange(existingRow + 1, headers.indexOf('markedAt') + 1).setValue(timestamp);
    attendanceSheet.getRange(existingRow + 1, headers.indexOf('markedBy') + 1).setValue(currentUser.userId);
    attendanceSheet.getRange(existingRow + 1, headers.indexOf('notes') + 1).setValue(attendanceData.notes || '');

    logActivity(currentUser.userId, currentUser.fullName, 'MARK_ATTENDANCE', 'attendance',
      data[existingRow][headers.indexOf('attendanceId')],
      'Updated attendance: ' + attendanceData.participantId + ' -> ' + attendanceData.status);

    return { success: true, message: 'Attendance updated successfully' };
  } else {
    // Create new attendance record
    const attendanceId = generateUUID();

    attendanceSheet.appendRow([
      attendanceId,
      attendanceData.sessionId,
      attendanceData.participantId,
      attendanceData.status,
      timestamp,
      currentUser.userId,
      attendanceData.notes || ''
    ]);

    logActivity(currentUser.userId, currentUser.fullName, 'MARK_ATTENDANCE', 'attendance', attendanceId,
      'Marked attendance: ' + attendanceData.participantId + ' -> ' + attendanceData.status);

    return { success: true, message: 'Attendance marked successfully', attendanceId: attendanceId };
  }
}

/**
 * Bulk mark attendance for multiple participants at once
 */
function bulkMarkAttendance(token, sessionId, attendanceRecords) {
  const currentUser = validateSession(token);
  if (!currentUser || (currentUser.role !== 'admin' && currentUser.role !== 'facilitator')) {
    return { success: false, message: 'Unauthorized - Admin or Facilitator access required' };
  }

  let successCount = 0;
  let errors = [];

  for (let record of attendanceRecords) {
    const result = markAttendance(token, {
      sessionId: sessionId,
      participantId: record.participantId,
      status: record.status,
      notes: record.notes || ''
    });

    if (result.success) {
      successCount++;
    } else {
      errors.push({ participantId: record.participantId, error: result.message });
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'BULK_MARK_ATTENDANCE', 'session', sessionId,
    'Bulk marked attendance for ' + successCount + ' participants');

  return {
    success: true,
    message: successCount + ' attendance records processed',
    successCount: successCount,
    errors: errors
  };
}

/**
 * Get attendance for a specific session
 */
function getAttendanceBySession(token, sessionId) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('SessionAttendance');

  if (!attendanceSheet) {
    return { success: false, message: 'SessionAttendance sheet not found' };
  }

  const data = attendanceSheet.getDataRange().getValues();
  const headers = data[0];
  const attendanceRecords = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('sessionId')] === sessionId) {
      attendanceRecords.push({
        attendanceId: data[i][headers.indexOf('attendanceId')],
        sessionId: data[i][headers.indexOf('sessionId')],
        participantId: data[i][headers.indexOf('participantId')],
        status: data[i][headers.indexOf('status')],
        markedAt: data[i][headers.indexOf('markedAt')],
        markedBy: data[i][headers.indexOf('markedBy')],
        notes: data[i][headers.indexOf('notes')]
      });
    }
  }

  return { success: true, attendance: attendanceRecords };
}

/**
 * Get attendance records for a specific participant
 */
function getAttendanceByParticipant(token, participantId) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheetByName('SessionAttendance');
  const sessionsSheet = ss.getSheetByName('StudySessions');

  if (!attendanceSheet || !sessionsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const attendanceData = attendanceSheet.getDataRange().getValues();
  const attendanceHeaders = attendanceData[0];
  const sessionData = sessionsSheet.getDataRange().getValues();
  const sessionHeaders = sessionData[0];

  const attendanceRecords = [];

  for (let i = 1; i < attendanceData.length; i++) {
    if (attendanceData[i][attendanceHeaders.indexOf('participantId')] === participantId) {
      const sessionId = attendanceData[i][attendanceHeaders.indexOf('sessionId')];

      // Find the session details
      let sessionInfo = null;
      for (let j = 1; j < sessionData.length; j++) {
        if (sessionData[j][sessionHeaders.indexOf('sessionId')] === sessionId) {
          sessionInfo = {
            sessionNumber: sessionData[j][sessionHeaders.indexOf('sessionNumber')],
            sessionDate: normalizeSessionDateValue(sessionData[j][sessionHeaders.indexOf('sessionDate')]),
            sessionName: sessionData[j][sessionHeaders.indexOf('sessionName')]
          };
          break;
        }
      }

      attendanceRecords.push({
        attendanceId: attendanceData[i][attendanceHeaders.indexOf('attendanceId')],
        sessionId: sessionId,
        participantId: attendanceData[i][attendanceHeaders.indexOf('participantId')],
        status: attendanceData[i][attendanceHeaders.indexOf('status')],
        markedAt: attendanceData[i][attendanceHeaders.indexOf('markedAt')],
        markedBy: attendanceData[i][attendanceHeaders.indexOf('markedBy')],
        notes: attendanceData[i][attendanceHeaders.indexOf('notes')],
        sessionInfo: sessionInfo
      });
    }
  }

  // Sort by session number
  attendanceRecords.sort((a, b) => {
    if (a.sessionInfo && b.sessionInfo) {
      return a.sessionInfo.sessionNumber - b.sessionInfo.sessionNumber;
    }
    return 0;
  });

  return { success: true, attendance: attendanceRecords };
}

/**
 * Get attendance statistics for a cohort
 */
function getAttendanceStatsByRollout(token, rolloutId) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const attendanceSheet = ss.getSheetByName('SessionAttendance');
  const participantsSheet = ss.getSheetByName('Participants');

  // Check if required sheets exist
  if (!sessionsSheet || !attendanceSheet || !participantsSheet) {
    return {
      success: false,
      message: 'Required sheets not found. Please run initializeDatabase() from the Apps Script editor to create all required sheets (StudySessions, SessionAttendance).'
    };
  }

  // Get all sessions for this cohort
  const sessionData = sessionsSheet.getDataRange().getValues();
  const sessionHeaders = sessionData[0];
  const sessions = [];

  for (let i = 1; i < sessionData.length; i++) {
    if (sessionData[i][sessionHeaders.indexOf('rolloutId')] === rolloutId) {
      sessions.push({
        sessionId: sessionData[i][sessionHeaders.indexOf('sessionId')],
        sessionNumber: sessionData[i][sessionHeaders.indexOf('sessionNumber')],
        sessionDate: normalizeSessionDateValue(sessionData[i][sessionHeaders.indexOf('sessionDate')]),
        sessionName: sessionData[i][sessionHeaders.indexOf('sessionName')]
      });
    }
  }

  // Get all participants for this cohort
  const participantData = participantsSheet.getDataRange().getValues();
  const participantHeaders = participantData[0];
  const participants = [];

  for (let i = 1; i < participantData.length; i++) {
    if (participantData[i][participantHeaders.indexOf('rolloutId')] === rolloutId) {
      participants.push({
        participantId: participantData[i][participantHeaders.indexOf('participantId')],
        fullName: participantData[i][participantHeaders.indexOf('fullName')]
      });
    }
  }

  // Get all attendance records
  const attendanceData = attendanceSheet.getDataRange().getValues();
  const attendanceHeaders = attendanceData[0];
  const attendanceMap = {};

  for (let i = 1; i < attendanceData.length; i++) {
    const sessionId = attendanceData[i][attendanceHeaders.indexOf('sessionId')];
    const participantId = attendanceData[i][attendanceHeaders.indexOf('participantId')];
    const status = attendanceData[i][attendanceHeaders.indexOf('status')];

    const key = sessionId + '_' + participantId;
    attendanceMap[key] = status;
  }

  // Calculate statistics
  const stats = {
    totalSessions: sessions.length,
    totalParticipants: participants.length,
    sessionStats: [],
    participantStats: []
  };

  // Per-session statistics
  for (let session of sessions) {
    let present = 0;
    let absent = 0;
    let excused = 0;
    let notMarked = 0;
    const presentParticipants = [];
    const absentParticipants = [];
    const excusedParticipants = [];
    const notMarkedParticipants = [];
    const nonPresentParticipants = [];

    for (let participant of participants) {
      const key = session.sessionId + '_' + participant.participantId;
      const status = attendanceMap[key];

      if (status === 'present') {
        present++;
        presentParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName
        });
      } else if (status === 'absent') {
        absent++;
        absentParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName
        });
        nonPresentParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName,
          status: 'absent'
        });
      } else if (status === 'excused') {
        excused++;
        excusedParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName
        });
        nonPresentParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName,
          status: 'excused'
        });
      } else {
        notMarked++;
        notMarkedParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName
        });
        nonPresentParticipants.push({
          participantId: participant.participantId,
          fullName: participant.fullName,
          status: 'not_marked'
        });
      }
    }

    stats.sessionStats.push({
      sessionId: session.sessionId,
      sessionNumber: session.sessionNumber,
      sessionDate: session.sessionDate,
      sessionName: session.sessionName,
      present: present,
      absent: absent,
      excused: excused,
      notMarked: notMarked,
      attendanceRate: participants.length > 0 ? Math.round((present / participants.length) * 100) : 0,
      presentParticipants: presentParticipants,
      absentParticipants: absentParticipants,
      excusedParticipants: excusedParticipants,
      notMarkedParticipants: notMarkedParticipants,
      nonPresentParticipants: nonPresentParticipants
    });
  }

  // Per-participant statistics
  for (let participant of participants) {
    let present = 0;
    let absent = 0;
    let excused = 0;
    let notMarked = 0;

    for (let session of sessions) {
      const key = session.sessionId + '_' + participant.participantId;
      const status = attendanceMap[key];

      if (status === 'present') present++;
      else if (status === 'absent') absent++;
      else if (status === 'excused') excused++;
      else notMarked++;
    }

    stats.participantStats.push({
      participantId: participant.participantId,
      fullName: participant.fullName,
      present: present,
      absent: absent,
      excused: excused,
      notMarked: notMarked,
      attendanceRate: sessions.length > 0 ? Math.round((present / sessions.length) * 100) : 0
    });
  }

  return { success: true, stats: stats };
}

// ============================================
// PARTICIPANT FUNCTIONS
// ============================================

const PARTICIPANT_PARENT_HEADERS = [
  'parent_guardian_names',
  'parent_guardian_phone',
  'parent_guardian_address',
  'parent_guardian_email',
  'parent_guardian_dob'
];

const ENROLLMENT_SYSTEM_HEADERS = [
  'participantId', 'site', 'rolloutId', 'schoolName', 'period', 'year',
  'enrollmentDate', 'enrolledBy', 'status', 'notes', 'completionPercentage'
];

const DEFAULT_ENROLLMENT_FIELDS = [
  { key: 'fullName', label: 'Participant Full Name', type: 'text', required: true, core: true, placeholder: "Enter participant's full name" },
  { key: 'parent_guardian_names', label: 'Parent/Guardian Name(s)', type: 'text', required: false, core: true, placeholder: 'Enter parent/guardian names (comma-separated if multiple)' },
  { key: 'parent_guardian_phone', label: 'Parent/Guardian Phone Number', type: 'phone', required: false, core: true, placeholder: '(706) 555-1234' },
  { key: 'parent_guardian_address', label: 'Parent/Guardian Residential Address', type: 'textarea', required: false, core: true, placeholder: 'Enter parent/guardian address' },
  { key: 'parent_guardian_email', label: 'Parent/Guardian Email Address', type: 'email', required: false, core: true, placeholder: 'parent@example.com' },
  { key: 'parent_guardian_dob', label: 'Parent/Guardian Date of Birth', type: 'date', required: false, core: true, placeholder: '' }
];

function slugifyEnrollmentFieldKey(label) {
  return 'custom_' + String(label || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .substring(0, 50);
}

function normalizeEnrollmentField(field, index, usedKeys) {
  const typeWhitelist = ['text', 'textarea', 'number', 'date', 'email', 'phone', 'select'];
  const label = String((field && field.label) || '').trim();
  if (!label) return null;

  let key = String((field && field.key) || '').trim();
  if (!key) key = slugifyEnrollmentFieldKey(label);
  key = key.replace(/[^a-zA-Z0-9_]/g, '_');
  if (!key) key = 'custom_field_' + (index + 1);
  if (ENROLLMENT_SYSTEM_HEADERS.indexOf(key) !== -1 || key === 'cohortName' || key === 'rolloutName') {
    key = 'custom_' + key;
  }

  const lowerUsed = usedKeys || {};
  const baseKey = key;
  let suffix = 2;
  while (lowerUsed[key.toLowerCase()]) {
    key = baseKey + '_' + suffix;
    suffix++;
  }
  lowerUsed[key.toLowerCase()] = true;

  const type = typeWhitelist.indexOf(String((field && field.type) || 'text')) !== -1
    ? String(field.type)
    : 'text';

  return {
    key: key,
    label: label,
    type: type,
    required: !!(field && field.required),
    core: !!(field && field.core),
    placeholder: String((field && field.placeholder) || ''),
    options: Array.isArray(field && field.options)
      ? field.options.map(option => String(option || '').trim()).filter(Boolean)
      : []
  };
}

function getEnrollmentFieldsInternal() {
  const used = {};
  const defaults = DEFAULT_ENROLLMENT_FIELDS.map((field, idx) => normalizeEnrollmentField(field, idx, used)).filter(Boolean);
  const defaultByKey = {};
  defaults.forEach(field => defaultByKey[field.key] = field);

  const configMap = getConfigMap();
  const raw = configMap.ENROLLMENT_FIELDS;
  if (!raw) return defaults;

  try {
    const parsed = JSON.parse(raw || '[]');
    if (!Array.isArray(parsed) || parsed.length === 0) return defaults;

    const normalizedUsed = {};
    const normalized = [];
    parsed.forEach((field, idx) => {
      const merged = defaultByKey[field && field.key]
        ? Object.assign({}, defaultByKey[field.key], field, { key: field.key, core: true })
        : field;
      const normalizedField = normalizeEnrollmentField(merged, idx, normalizedUsed);
      if (normalizedField) normalized.push(normalizedField);
    });

    if (!normalized.some(field => field.key === 'fullName')) {
      normalized.unshift(Object.assign({}, defaultByKey.fullName));
    }
    normalized.forEach(field => {
      if (field.key === 'fullName') {
        field.required = true;
        field.core = true;
        field.type = 'text';
      }
    });
    return normalized;
  } catch (e) {
    return defaults;
  }
}

function getEnrollmentFields(token) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Unauthorized' };
  return { success: true, fields: getEnrollmentFieldsInternal() };
}

function saveEnrollmentFields(token, fields) {
  const currentUser = validateSession(token);
  if (!canAccessEnrollmentFieldManager(currentUser)) return { success: false, message: 'Unauthorized' };
  if (!Array.isArray(fields) || fields.length === 0) return { success: false, message: 'At least one enrollment field is required' };

  const used = {};
  const normalized = fields.map((field, idx) => normalizeEnrollmentField(field, idx, used)).filter(Boolean);
  if (!normalized.some(field => field.key === 'fullName')) {
    normalized.unshift(Object.assign({}, DEFAULT_ENROLLMENT_FIELDS[0]));
  }
  normalized.forEach(field => {
    if (field.key === 'fullName') {
      field.required = true;
      field.core = true;
      field.type = 'text';
    }
  });
  if (normalized.length === 0) return { success: false, message: 'At least one valid enrollment field is required' };

  upsertConfigValue('ENROLLMENT_FIELDS', JSON.stringify(normalized), 'Participant enrollment form fields');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  if (participantsSheet) ensureParticipantColumns(participantsSheet);

  return { success: true, fields: normalized, message: 'Enrollment fields saved' };
}

function getEnrollmentFieldValueMapFromData(data) {
  const values = Object.assign({}, (data && data.enrollmentFields) || {});
  if (data && data.fullName !== undefined) values.fullName = data.fullName;
  if (data && data.parentGuardianNames !== undefined) values.parent_guardian_names = data.parentGuardianNames;
  if (data && data.parentGuardianPhone !== undefined) values.parent_guardian_phone = data.parentGuardianPhone;
  if (data && data.parentGuardianAddress !== undefined) values.parent_guardian_address = data.parentGuardianAddress;
  if (data && data.parentGuardianEmail !== undefined) values.parent_guardian_email = data.parentGuardianEmail;
  if (data && data.parentGuardianDob !== undefined) values.parent_guardian_dob = data.parentGuardianDob;
  return values;
}

function validateEnrollmentFieldValues(fields, values) {
  const errors = [];
  fields.forEach(field => {
    const value = String(values[field.key] || '').trim();
    if (field.required && !value) {
      errors.push(field.label + ' is required');
      return;
    }
    if (!value) return;
    if (field.type === 'email' && !isValidEmail(value)) errors.push(field.label + ' is invalid');
    if (field.type === 'date' && !isValidDateValue(value)) errors.push(field.label + ' is invalid');
    if (field.type === 'phone') {
      const digits = normalizePhoneDigits(value);
      if (digits.length < 10 || digits.length > 15) errors.push(field.label + ' must include 10 to 15 digits');
    }
  });
  return errors;
}

function ensureParticipantColumns(participantsSheet) {
  const requiredHeaders = [
    'participantId', 'fullName',
    'parent_guardian_names', 'parent_guardian_phone', 'parent_guardian_address',
    'parent_guardian_email', 'parent_guardian_dob',
    'site', 'rolloutId', 'schoolName', 'period', 'year',
    'enrollmentDate', 'enrolledBy', 'status', 'notes', 'completionPercentage'
  ];

  getEnrollmentFieldsInternal().forEach(field => {
    if (requiredHeaders.indexOf(field.key) === -1 && ENROLLMENT_SYSTEM_HEADERS.indexOf(field.key) === -1) {
      requiredHeaders.push(field.key);
    }
  });

  const headerRange = participantsSheet.getRange(1, 1, 1, participantsSheet.getLastColumn() || 1);
  const headers = headerRange.getValues()[0].filter(Boolean);
  let updated = false;

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      updated = true;
    }
  });

  if (updated) {
    participantsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return headers;
}

function ensureChecklistColumns(checklistSheet) {
  const requiredHeaders = [
    'checklistId', 'participantId', 'instrumentNumber', 'instrumentName', 'category',
    'status', 'completedDate', 'completedBy', 'notes', 'dataLink'
  ];

  const headerRange = checklistSheet.getRange(1, 1, 1, checklistSheet.getLastColumn() || 1);
  const headers = headerRange.getValues()[0].filter(Boolean);
  let updated = false;

  requiredHeaders.forEach(header => {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      updated = true;
    }
  });

  if (updated) {
    checklistSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return headers;
}

function getHeaderValue(row, headers, headerName) {
  const index = headers.indexOf(headerName);
  return index === -1 ? '' : row[index];
}

function setCellAsPlainText(sheet, rowIndex, columnIndex, value) {
  const range = sheet.getRange(rowIndex, columnIndex, 1, 1);
  range.setNumberFormat('@');
  range.setValue(value);
  invalidateSheetSnapshot(sheet.getName());
}

function appendRowsAsPlainText(sheet, rows) {
  if (!sheet || !Array.isArray(rows) || rows.length === 0) return 0;
  const startRow = sheet.getLastRow() + 1;
  const columnCount = rows[0].length;
  const range = sheet.getRange(startRow, 1, rows.length, columnCount);
  range.setNumberFormat('@');
  range.setValues(rows);
  invalidateSheetSnapshot(sheet.getName());
  return startRow;
}

function appendParticipantRowAsText(participantsSheet, rowValues) {
  return appendRowsAsPlainText(participantsSheet, [rowValues]);
}

function normalizePhoneDigits(phone) {
  return (phone || '').toString().replace(/\D/g, '');
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidDateValue(value) {
  if (!value) return true;
  const date = new Date(value);
  return !isNaN(date.getTime());
}

function buildParticipantRow(headers, data) {
  const row = new Array(headers.length).fill('');
  const setValue = (header, value) => {
    const idx = headers.indexOf(header);
    if (idx !== -1) {
      row[idx] = value;
    }
  };

  setValue('participantId', data.participantId);
  setValue('fullName', data.fullName);
  setValue('parent_guardian_names', data.parentGuardianNames || '');
  setValue('parent_guardian_phone', data.parentGuardianPhone || '');
  setValue('parent_guardian_address', data.parentGuardianAddress || '');
  setValue('parent_guardian_email', data.parentGuardianEmail || '');
  setValue('parent_guardian_dob', data.parentGuardianDob || '');
  setValue('site', data.site);
  setValue('rolloutId', data.rolloutId);
  setValue('schoolName', data.schoolName);
  setValue('period', data.period);
  setValue('year', data.year);
  setValue('enrollmentDate', data.enrollmentDate);
  setValue('enrolledBy', data.enrolledBy);
  setValue('status', data.status);
  setValue('notes', data.notes || '');
  setValue('completionPercentage', data.completionPercentage || 0);

  const enrollmentValues = getEnrollmentFieldValueMapFromData(data);
  getEnrollmentFieldsInternal().forEach(field => {
    if (enrollmentValues[field.key] !== undefined) {
      setValue(field.key, enrollmentValues[field.key]);
    }
  });

  return row;
}

/**
 * Get all participants with optional filters
 */
function getAllParticipants(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');

  if (!participantsSheet) return { success: false, message: 'Participants sheet not found' };

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const headers = participantSnapshot.headers;
  const data = participantSnapshot.data;
  const participants = [];
  const filteredRows = [];

  filters = filters || {};

  const participantIdCol = headers.indexOf('participantId');
  const fullNameCol = headers.indexOf('fullName');
  const siteCol = headers.indexOf('site');
  const rolloutIdCol = headers.indexOf('rolloutId');
  const schoolNameCol = headers.indexOf('schoolName');
  const statusCol = headers.indexOf('status');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const participantId = row[participantIdCol];
    const fullName = row[fullNameCol] || '';
    const site = row[siteCol];
    const rolloutId = row[rolloutIdCol];
    const schoolName = row[schoolNameCol] || '';
    const status = row[statusCol];

    let include = true;
    if (filters.site && filters.site !== 'All' && site !== filters.site) include = false;
    if (filters.rolloutId && filters.rolloutId !== 'All' && rolloutId !== filters.rolloutId) include = false;
    if (filters.status && status !== filters.status) include = false;
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      const matchesSearch =
        String(fullName).toLowerCase().includes(searchLower) ||
        String(participantId).toLowerCase().includes(searchLower) ||
        String(schoolName).toLowerCase().includes(searchLower);
      if (!matchesSearch) include = false;
    }
    if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && site !== currentUser.site) include = false;
    if (include) filteredRows.push(row);
  }

  if (filteredRows.length === 0) {
    return { success: true, participants: [] };
  }

  const completionMap = buildLiveCompletionMap(filteredRows.map(row => row[participantIdCol]).filter(Boolean));

  for (let i = 0; i < filteredRows.length; i++) {
    const row = filteredRows[i];
    const participantId = row[participantIdCol];
    const participant = {
      participantId: participantId,
      fullName: row[fullNameCol],
      site: row[siteCol],
      rolloutId: row[rolloutIdCol],
      schoolName: row[schoolNameCol],
      period: getHeaderValue(row, headers, 'period'),
      year: getHeaderValue(row, headers, 'year'),
      enrollmentDate: getHeaderValue(row, headers, 'enrollmentDate'),
      enrolledBy: getHeaderValue(row, headers, 'enrolledBy'),
      status: row[statusCol],
      notes: getHeaderValue(row, headers, 'notes'),
      completionPercentage: (completionMap[participantId] || {}).percentage || 0
    };

    if (currentUser.role !== 'viewer') {
      participant.parentGuardianNames = getHeaderValue(row, headers, 'parent_guardian_names');
      participant.parentGuardianPhone = getHeaderValue(row, headers, 'parent_guardian_phone');
      participant.parentGuardianAddress = getHeaderValue(row, headers, 'parent_guardian_address');
      participant.parentGuardianEmail = getHeaderValue(row, headers, 'parent_guardian_email');
      participant.parentGuardianDob = getHeaderValue(row, headers, 'parent_guardian_dob');
      participant.enrollmentFields = {};
      getEnrollmentFieldsInternal().forEach(field => {
        participant.enrollmentFields[field.key] = getHeaderValue(row, headers, field.key);
      });
    }

    participants.push(participant);
  }

  return { success: true, participants: participants };
}

/**
 * Get participant by ID with checklist
 */
function getParticipantById(token, participantId) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Get participant data
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  let participant = null;

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('participantId')] === participantId) {
      participant = {
        participantId: getHeaderValue(pData[i], pHeaders, 'participantId'),
        fullName: getHeaderValue(pData[i], pHeaders, 'fullName'),
        site: getHeaderValue(pData[i], pHeaders, 'site'),
        rolloutId: getHeaderValue(pData[i], pHeaders, 'rolloutId'),
        schoolName: getHeaderValue(pData[i], pHeaders, 'schoolName'),
        period: getHeaderValue(pData[i], pHeaders, 'period'),
        year: getHeaderValue(pData[i], pHeaders, 'year'),
        enrollmentDate: getHeaderValue(pData[i], pHeaders, 'enrollmentDate'),
        enrolledBy: getHeaderValue(pData[i], pHeaders, 'enrolledBy'),
        status: getHeaderValue(pData[i], pHeaders, 'status'),
        notes: getHeaderValue(pData[i], pHeaders, 'notes'),
        completionPercentage: getHeaderValue(pData[i], pHeaders, 'completionPercentage') || 0
      };

      if (currentUser.role !== 'viewer') {
        participant.parentGuardianNames = getHeaderValue(pData[i], pHeaders, 'parent_guardian_names');
        participant.parentGuardianPhone = getHeaderValue(pData[i], pHeaders, 'parent_guardian_phone');
        participant.parentGuardianAddress = getHeaderValue(pData[i], pHeaders, 'parent_guardian_address');
        participant.parentGuardianEmail = getHeaderValue(pData[i], pHeaders, 'parent_guardian_email');
        participant.parentGuardianDob = getHeaderValue(pData[i], pHeaders, 'parent_guardian_dob');
        participant.enrollmentFields = {};
        getEnrollmentFieldsInternal().forEach(field => {
          participant.enrollmentFields[field.key] = getHeaderValue(pData[i], pHeaders, field.key);
        });
      }
      break;
    }
  }

  if (!participant) {
    return { success: false, message: 'Participant not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && participant.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  // Get checklist items
  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cHeaders = checklistSnapshot.headers;
  const cData = checklistSnapshot.data;
  const checklist = [];

  for (let i = 1; i < cData.length; i++) {
    if (cData[i][cHeaders.indexOf('participantId')] === participantId) {
      checklist.push({
        checklistId: cData[i][cHeaders.indexOf('checklistId')],
        participantId: cData[i][cHeaders.indexOf('participantId')],
        instrumentNumber: cData[i][cHeaders.indexOf('instrumentNumber')],
        instrumentName: cData[i][cHeaders.indexOf('instrumentName')],
        category: cData[i][cHeaders.indexOf('category')],
        status: cData[i][cHeaders.indexOf('status')],
        completedDate: cData[i][cHeaders.indexOf('completedDate')],
        completedBy: cData[i][cHeaders.indexOf('completedBy')],
        notes: cData[i][cHeaders.indexOf('notes')],
        dataLink: cData[i][cHeaders.indexOf('dataLink')]
      });
    }
  }

  // Sort checklist by instrument number
  checklist.sort((a, b) => a.instrumentNumber - b.instrumentNumber);

  return { success: true, participant: participant, checklist: checklist };
}

/**
 * Enroll a new participant (Admin and Facilitator)
 */
function enrollParticipant(token, participantData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const enrollmentFields = getEnrollmentFieldsInternal();
  const enrollmentValues = getEnrollmentFieldValueMapFromData(participantData || {});
  enrollmentValues.fullName = String(enrollmentValues.fullName || '').trim();
  enrollmentValues.parent_guardian_email = String(enrollmentValues.parent_guardian_email || '').trim();
  enrollmentValues.parent_guardian_dob = String(enrollmentValues.parent_guardian_dob || '').trim();
  enrollmentValues.parent_guardian_phone = String(enrollmentValues.parent_guardian_phone || '').trim();

  const enrollmentErrors = validateEnrollmentFieldValues(enrollmentFields, enrollmentValues);
  if (enrollmentErrors.length) {
    return { success: false, message: enrollmentErrors.join('; ') };
  }
  if (!participantData || !participantData.rolloutId) {
    return { success: false, message: 'Study cohort is required' };
  }

  const parentEmail = enrollmentValues.parent_guardian_email;
  const parentDob = enrollmentValues.parent_guardian_dob;
  const parentPhone = enrollmentValues.parent_guardian_phone;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const participantHeaders = ensureParticipantColumns(participantsSheet);
  const checklistHeaders = ensureChecklistColumns(checklistSheet);

  // Get cohort info
  const cohort = getRolloutById(participantData.rolloutId);
  if (!cohort) {
    return { success: false, message: 'Invalid cohort selected' };
  }

  if (currentUser.role === 'facilitator') {
    if (!currentUser.site || currentUser.site === 'All') {
      return { success: false, message: 'Facilitators must be assigned to a single site before enrolling participants' };
    }

    if (cohort.site !== currentUser.site) {
      return { success: false, message: 'Unauthorized to enroll participants for this site' };
    }
  }

  // Generate participant ID
  const participantId = generateParticipantId(cohort.site);
  const timestamp = new Date().toISOString();

  // Add participant
  const row = buildParticipantRow(participantHeaders, {
    participantId: participantId,
    fullName: enrollmentValues.fullName,
    enrollmentFields: enrollmentValues,
    parentGuardianNames: enrollmentValues.parent_guardian_names,
    parentGuardianPhone: parentPhone,
    parentGuardianAddress: enrollmentValues.parent_guardian_address,
    parentGuardianEmail: parentEmail,
    parentGuardianDob: parentDob,
    site: cohort.site,
    rolloutId: participantData.rolloutId,
    schoolName: cohort.schoolName,
    period: cohort.period,
    year: cohort.year,
    enrollmentDate: timestamp,
    enrolledBy: currentUser.userId,
    status: 'active',
    notes: participantData.notes || '',
    completionPercentage: 0
  });
  appendParticipantRowAsText(participantsSheet, row);

  // Create checklist items for rollout-configured protocol items in one bulk write.
  const checklistRows = getRolloutProtocolItemsInternal(participantData.rolloutId).map(instrument => {
    const checklistId = generateUUID();
    const checklistRow = new Array(checklistHeaders.length).fill('');
    const setChecklistValue = (header, value) => {
      const idx = checklistHeaders.indexOf(header);
      if (idx !== -1) checklistRow[idx] = value;
    };
    setChecklistValue('checklistId', checklistId);
    setChecklistValue('participantId', participantId);
    setChecklistValue('instrumentNumber', instrument.number);
    setChecklistValue('instrumentName', instrument.name);
    setChecklistValue('category', instrument.category);
    setChecklistValue('status', 'not_started');
    return checklistRow;
  });
  appendRowsAsPlainText(checklistSheet, checklistRows);

  logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT', 'participant', participantId,
    'Enrolled: ' + enrollmentValues.fullName);

  return {
    success: true,
    message: 'Participant enrolled successfully',
    participantId: participantId
  };
}

/**
 * Update participant info
 */
function updateParticipant(token, participantId, participantData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const canChangeStatus = currentUser.role === 'admin';

  if (participantData.status && !canChangeStatus) {
    return { success: false, message: 'Only administrators can change participant status' };
  }

  if (participantData.status && CONFIG.PARTICIPANT_STATUSES.indexOf(participantData.status) === -1) {
    return { success: false, message: 'Invalid status selected' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const headers = participantSnapshot.headers;
  const data = participantSnapshot.data;

  const enrollmentFields = getEnrollmentFieldsInternal();
  const enrollmentValues = getEnrollmentFieldValueMapFromData(participantData || {});
  enrollmentValues.fullName = String(enrollmentValues.fullName || '').trim();
  enrollmentValues.parent_guardian_email = String(enrollmentValues.parent_guardian_email || '').trim();
  enrollmentValues.parent_guardian_dob = String(enrollmentValues.parent_guardian_dob || '').trim();
  enrollmentValues.parent_guardian_phone = String(enrollmentValues.parent_guardian_phone || '').trim();

  const enrollmentErrors = validateEnrollmentFieldValues(enrollmentFields, enrollmentValues);
  if (enrollmentErrors.length) {
    return { success: false, message: enrollmentErrors.join('; ') };
  }

  const parentEmail = enrollmentValues.parent_guardian_email;
  const parentDob = enrollmentValues.parent_guardian_dob;
  const parentPhone = enrollmentValues.parent_guardian_phone;

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('participantId')] === participantId) {
      if (currentUser.role === 'facilitator' && currentUser.site !== 'All' &&
          data[i][headers.indexOf('site')] !== currentUser.site) {
        return { success: false, message: 'Unauthorized for this site' };
      }

      const rowValues = data[i].slice(0, headers.length);
      while (rowValues.length < headers.length) rowValues.push('');
      const setRowValue = (header, value) => {
        const idx = headers.indexOf(header);
        if (idx !== -1) rowValues[idx] = value;
      };

      if (participantData.fullName) setRowValue('fullName', participantData.fullName);
      if (participantData.parentGuardianNames !== undefined) setRowValue('parent_guardian_names', participantData.parentGuardianNames);
      if (participantData.parentGuardianPhone !== undefined) setRowValue('parent_guardian_phone', parentPhone);
      if (participantData.parentGuardianAddress !== undefined) setRowValue('parent_guardian_address', participantData.parentGuardianAddress);
      if (participantData.parentGuardianEmail !== undefined) setRowValue('parent_guardian_email', parentEmail);
      if (participantData.parentGuardianDob !== undefined) setRowValue('parent_guardian_dob', parentDob);
      enrollmentFields.forEach(field => {
        if (enrollmentValues[field.key] !== undefined) setRowValue(field.key, enrollmentValues[field.key]);
      });
      if (participantData.status) setRowValue('status', participantData.status);
      if (participantData.notes !== undefined) setRowValue('notes', participantData.notes);

      const targetRange = participantsSheet.getRange(i + 1, 1, 1, headers.length);
      targetRange.setNumberFormat('@');
      targetRange.setValues([rowValues]);
      invalidateSheetSnapshot('Participants');

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_PARTICIPANT', 'participant', participantId,
        'Updated participant info');

      return { success: true, message: 'Participant updated successfully' };
    }
  }

  return { success: false, message: 'Participant not found' };
}

/**
 * Delete a participant and their checklist entries (Admin only)
 */
function deleteParticipant(token, participantId) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role !== 'admin') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet) {
    return { success: false, message: 'Participants sheet not found' };
  }

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const participantIdCol = pHeaders.indexOf('participantId');
  let participantName = '';
  let participantSite = '';
  let rowIndex = -1;

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][participantIdCol] === participantId) {
      participantName = pData[i][pHeaders.indexOf('fullName')];
      participantSite = pData[i][pHeaders.indexOf('site')];
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    return { success: false, message: 'Participant not found' };
  }

  participantsSheet.deleteRow(rowIndex);
  invalidateSheetSnapshot('Participants');

  if (checklistSheet) {
    const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
    const cData = checklistSnapshot.data;
    const cHeaders = checklistSnapshot.headers;
    const checklistParticipantCol = cHeaders.indexOf('participantId');

    for (let i = cData.length - 1; i >= 1; i--) {
      if (cData[i][checklistParticipantCol] === participantId) {
        checklistSheet.deleteRow(i + 1);
      }
    }
  }

  logActivity(currentUser.userId, currentUser.fullName, 'DELETE_PARTICIPANT', 'participant', participantId,
    'Deleted participant: ' + participantName + ' (' + participantSite + ')');

  return { success: true, message: 'Participant deleted successfully' };
}

/**
 * Generate participant ID with site prefix
 */
function generateParticipantId(site) {
  const prefix = site === 'UGA' ? 'UGA' : 'MIZ';
  const timestamp = Date.now().toString(36).toUpperCase();
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + '-' + timestamp + random;
}

/**
 * Generate a CSV template for batch participant enrollment
 */
function generateEnrollmentTemplate(token) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  if (!rolloutsSheet) {
    return { success: false, message: 'Study cohorts sheet not found' };
  }

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const cohorts = [];
  const isFacilitator = currentUser.role === 'facilitator' && currentUser.site && currentUser.site !== 'All';

  for (let i = 1; i < data.length; i++) {
    const site = data[i][headers.indexOf('site')];
    if (isFacilitator && site !== currentUser.site) continue;
    cohorts.push({
      rolloutId: data[i][headers.indexOf('rolloutId')],
      site: site,
      schoolName: data[i][headers.indexOf('schoolName')],
      period: data[i][headers.indexOf('period')],
      year: data[i][headers.indexOf('year')],
      status: data[i][headers.indexOf('status')]
    });
  }

  // Build dropdown-enabled template as an Excel file
  const tempSs = SpreadsheetApp.create('Participant Enrollment Template');
  const templateSheet = tempSs.getSheets()[0];
  templateSheet.setName('Template');

  // Headers are generated from the live enrollment-field configuration.
  const enrollmentFields = getEnrollmentFieldsInternal();
  const headerValues = enrollmentFields.map(field => field.key).concat(['cohortName']);
  templateSheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]);
  templateSheet.getRange(1, 1, 1, headerValues.length)
    .setFontWeight('bold')
    .setBackground('#f1f5f9');
  enrollmentFields.forEach((field, index) => {
    templateSheet.getRange(1, index + 1).setNote(
      field.label + (field.required ? ' (required)' : ' (optional)') +
      (field.type ? ' — ' + field.type : '')
    );
  });
  templateSheet.getRange(1, headerValues.length).setNote('Choose the cohort/rollout for each participant (required).');

  // Helper sheet with cohort list
  const helperSheet = tempSs.insertSheet('Cohorts');
  if (cohorts.length > 0) {
    helperSheet.getRange(1, 1, cohorts.length, 1).setValues(
      cohorts.map(r => [`${r.schoolName} (${r.period} ${r.year})`])
    );
  }
  helperSheet.hideSheet();

  // Data validation for cohort dropdown (apply to reasonable range)
  const lastRow = Math.max(2, cohorts.length + 5);
  const cohortColumnIndex = headerValues.indexOf('cohortName') + 1;
  if (cohorts.length > 0) {
    const validationRange = helperSheet.getRange(1, 1, cohorts.length, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(validationRange, true)
      .setAllowInvalid(false)
      .build();
    templateSheet.getRange(2, cohortColumnIndex, lastRow, 1).setDataValidation(rule);
  }

  // Auto-size
  templateSheet.autoResizeColumns(1, headerValues.length);

  // Ensure writes complete before export
  SpreadsheetApp.flush();

  // Export as Excel using Drive export endpoint to avoid format issues
  const exportUrl = `https://docs.google.com/spreadsheets/d/${tempSs.getId()}/export?format=xlsx`;
  const response = UrlFetchApp.fetch(exportUrl, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    return { success: false, message: 'Failed to generate template file' };
  }

  const blob = response.getBlob().setName('participant_enrollment_template.xlsx');
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);

  return {
    success: true,
    file: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType(),
    filename: blob.getName(),
    rolloutCount: cohorts.length
  };
}

/**
 * Import participants from CSV (batch enrollment/upsert)
 */
function importParticipantsCSV(token, fileData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  if (!fileData || !fileData.csv) {
    return { success: false, message: 'No CSV content provided' };
  }

  let rows;
  try {
    rows = Utilities.parseCsv(fileData.csv);
  } catch (error) {
    return { success: false, message: 'Unable to parse CSV: ' + error.message };
  }

  if (!rows || rows.length === 0) {
    return { success: false, message: 'CSV is empty' };
  }

  const headers = rows[0].map(h => h.trim());
  const enrollmentFields = getEnrollmentFieldsInternal();
  const hasCohortName = headers.indexOf('cohortName') !== -1;
  const hasRolloutName = headers.indexOf('rolloutName') !== -1;
  const missing = [];
  enrollmentFields.forEach(field => {
    if (field.required && headers.indexOf(field.key) === -1) {
      missing.push(field.key);
    }
  });
  if (!hasCohortName && !hasRolloutName) {
    missing.push('cohortName');
  }
  if (missing.length > 0) {
    return { success: false, message: 'Missing required columns: ' + missing.join(', ') };
  }

  const getValue = (row, name) => {
    const index = headers.indexOf(name);
    return index === -1 ? '' : (row[index] || '');
  };
  const cohortColumn = hasCohortName ? 'cohortName' : 'rolloutName';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!participantsSheet || !checklistSheet || !rolloutsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Build cohort map for quick lookups
  const rolloutData = rolloutsSheet.getDataRange().getValues();
  const rolloutHeaders = rolloutData[0];
  const rolloutMap = {};
  const isFacilitator = currentUser.role === 'facilitator' && currentUser.site && currentUser.site !== 'All';

  for (let i = 1; i < rolloutData.length; i++) {
    const rolloutId = rolloutData[i][rolloutHeaders.indexOf('rolloutId')];
    const site = rolloutData[i][rolloutHeaders.indexOf('site')];
    if (isFacilitator && site !== currentUser.site) continue;

    const label = `${rolloutData[i][rolloutHeaders.indexOf('schoolName')]} (${rolloutData[i][rolloutHeaders.indexOf('period')]} ${rolloutData[i][rolloutHeaders.indexOf('year')]})`;
    rolloutMap[label.toLowerCase()] = {
      rolloutId: rolloutId,
      site: site,
      schoolName: rolloutData[i][rolloutHeaders.indexOf('schoolName')],
      period: rolloutData[i][rolloutHeaders.indexOf('period')],
      year: rolloutData[i][rolloutHeaders.indexOf('year')],
      status: rolloutData[i][rolloutHeaders.indexOf('status')],
      label: label
    };
  }

  if (Object.keys(rolloutMap).length === 0) {
    return { success: false, message: 'No accessible cohorts found for this user' };
  }

  const participantHeaders = ensureParticipantColumns(participantsSheet);
  const checklistHeaders = ensureChecklistColumns(checklistSheet);
  const participantRowsToAppend = [];
  const checklistRowsToAppend = [];

  const summary = {
    processed: 0,
    created: 0,
    updated: 0,
    skipped: 0,
    errors: []
  };

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rowNumber = r + 1; // 1-based CSV row

    const enrollmentValues = {};
    enrollmentFields.forEach(field => {
      enrollmentValues[field.key] = (getValue(row, field.key) || '').trim();
    });
    const fullName = String(enrollmentValues.fullName || '').trim();
    const rolloutName = (getValue(row, cohortColumn) || '').trim();
    const parentGuardianNames = String(enrollmentValues.parent_guardian_names || '').trim();
    const parentGuardianPhone = String(enrollmentValues.parent_guardian_phone || '').trim();
    const parentGuardianAddress = String(enrollmentValues.parent_guardian_address || '').trim();
    const parentGuardianEmail = String(enrollmentValues.parent_guardian_email || '').trim();
    const parentGuardianDob = String(enrollmentValues.parent_guardian_dob || '').trim();

    if (!fullName && !rolloutName) {
      summary.skipped++;
      continue;
    }

    summary.processed++;

    const rowErrors = validateEnrollmentFieldValues(enrollmentFields, enrollmentValues);
    if (!rolloutName) {
      rowErrors.push('cohortName is required');
    }
    if (rowErrors.length) {
      summary.errors.push({ row: rowNumber, message: rowErrors.join('; ') });
      summary.skipped++;
      continue;
    }
    const cohort = rolloutMap[rolloutName.toLowerCase()];
    if (!cohort) {
      summary.errors.push({ row: rowNumber, message: 'Cohort not found or not accessible: ' + rolloutName });
      summary.skipped++;
      continue;
    }

    // Create new participant (participantId auto-generated)
    if (isFacilitator && cohort.site !== currentUser.site) {
      summary.errors.push({ row: rowNumber, message: 'Unauthorized to enroll for site ' + cohort.site });
      summary.skipped++;
      continue;
    }

    const newParticipantId = generateParticipantId(cohort.site);
    const timestamp = new Date().toISOString();

    const rowValues = buildParticipantRow(participantHeaders, {
      participantId: newParticipantId,
      fullName: fullName,
      enrollmentFields: enrollmentValues,
      parentGuardianNames: parentGuardianNames,
      parentGuardianPhone: parentGuardianPhone,
      parentGuardianAddress: parentGuardianAddress,
      parentGuardianEmail: parentGuardianEmail,
      parentGuardianDob: parentGuardianDob,
      site: cohort.site,
      rolloutId: cohort.rolloutId,
      schoolName: cohort.schoolName,
      period: cohort.period,
      year: cohort.year,
      enrollmentDate: timestamp,
      enrolledBy: currentUser.userId,
      status: 'active',
      notes: '',
      completionPercentage: 0
    });
    participantRowsToAppend.push(rowValues);

    getRolloutProtocolItemsInternal(cohort.rolloutId).forEach(instrument => {
      const checklistId = generateUUID();
      const checklistRow = new Array(checklistHeaders.length).fill('');
      const setChecklistValue = (header, value) => {
        const idx = checklistHeaders.indexOf(header);
        if (idx !== -1) checklistRow[idx] = value;
      };
      setChecklistValue('checklistId', checklistId);
      setChecklistValue('participantId', newParticipantId);
      setChecklistValue('instrumentNumber', instrument.number);
      setChecklistValue('instrumentName', instrument.name);
      setChecklistValue('category', instrument.category);
      setChecklistValue('status', 'not_started');
      checklistRowsToAppend.push(checklistRow);
    });

    logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT_BULK', 'participant', newParticipantId,
      'Bulk enrolled: ' + fullName + ' into ' + cohort.schoolName);

    summary.created++;
  }

  appendRowsAsPlainText(participantsSheet, participantRowsToAppend);
  appendRowsAsPlainText(checklistSheet, checklistRowsToAppend);

  return {
    success: true,
    summary: summary
  };
}

// ============================================
// CHECKLIST FUNCTIONS
// ============================================

/**
 * Update checklist item status
 */
function updateChecklistItem(token, checklistId, updateData) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checklistSheet = ss.getSheetByName('Checklist');
  const participantsSheet = ss.getSheetByName('Participants');

  if (!checklistSheet || !participantsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const headers = checklistSnapshot.headers;
  const data = checklistSnapshot.data;

  // Build participant site map for authorization checks from the cached participant snapshot.
  const participantSiteMap = {};
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  for (let i = 1; i < pData.length; i++) {
    participantSiteMap[pData[i][pHeaders.indexOf('participantId')]] = pData[i][pHeaders.indexOf('site')];
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('checklistId')] === checklistId) {
      const rowValues = data[i].slice(0, headers.length);
      while (rowValues.length < headers.length) rowValues.push('');
      const getRowValue = header => rowValues[headers.indexOf(header)];
      const setRowValue = (header, value) => {
        const idx = headers.indexOf(header);
        if (idx !== -1) rowValues[idx] = value;
      };
      const participantId = getRowValue('participantId');
      const instrumentName = getRowValue('instrumentName');
      const participantSite = participantSiteMap[participantId];

      if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
        if (!participantSite) {
          return { success: false, message: 'Unable to verify participant site' };
        }
        if (participantSite !== currentUser.site) {
          return { success: false, message: 'Unauthorized for this site' };
        }
      }

      const updateDetails = [];

      // Update status
      if (updateData.status) {
        setRowValue('status', updateData.status);
        updateDetails.push('status -> ' + updateData.status);

        // If marking as completed, set the date and user
        if (updateData.status === 'completed') {
          setRowValue('completedDate', new Date().toISOString());
          setRowValue('completedBy', currentUser.fullName);
        } else {
          // Clear completion info if status changed to not_started or missing
          setRowValue('completedDate', '');
          setRowValue('completedBy', '');
        }
      }

      // Update notes
      if (updateData.notes !== undefined) {
        setRowValue('notes', updateData.notes);
        updateDetails.push('notes updated');
      }

      if (updateData.dataLink !== undefined) {
        setRowValue('dataLink', updateData.dataLink);
        updateDetails.push(updateData.dataLink ? 'data link saved' : 'data link removed');
      }

      checklistSheet.getRange(i + 1, 1, 1, headers.length).setValues([rowValues]);
      invalidateSheetSnapshot('Checklist');

      // Update participant's completion percentage
      updateParticipantCompletion(participantId);
      syncParticipantStatusFromChecklist(participantId);

      logActivity(currentUser.userId, currentUser.fullName, 'UPDATE_CHECKLIST', 'checklist', checklistId,
        'Updated: ' + instrumentName + (updateDetails.length ? ' (' + updateDetails.join(', ') + ')' : ''));

      return { success: true, message: 'Checklist item updated' };
    }
  }

  return { success: false, message: 'Checklist item not found' };
}

function syncParticipantStatusFromChecklist(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  if (!participantsSheet || !checklistSheet) return;

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  let participantRow = -1;
  let currentStatus = '';
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][pHeaders.indexOf('participantId')]) === String(participantId)) {
      participantRow = i + 1;
      currentStatus = String(pData[i][pHeaders.indexOf('status')] || '');
      break;
    }
  }
  if (participantRow === -1) return;
  if (currentStatus === 'withdrawn') return; // never auto-overwrite withdrawn

  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cHeaders = checklistSnapshot.headers;
  const cData = checklistSnapshot.data;
  let hasPostTestCompleted = false;
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][cHeaders.indexOf('participantId')]) !== String(participantId)) continue;
    const name = String(cData[i][cHeaders.indexOf('instrumentName')] || '').toLowerCase();
    const status = String(cData[i][cHeaders.indexOf('status')] || '');
    if (name.indexOf('post-test') !== -1 || name.indexOf('post test') !== -1) {
      if (status === 'completed') {
        hasPostTestCompleted = true;
        break;
      }
    }
  }

  const targetStatus = hasPostTestCompleted ? 'completed' : 'active';
  if (currentStatus !== targetStatus) {
    setCellAsPlainText(participantsSheet, participantRow, pHeaders.indexOf('status') + 1, targetStatus);
  }
}

function bulkUpdateParticipantStatus(token, rolloutId, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }
  const cohort = getRolloutById(rolloutId);
  if (!cohort) return { success: false, message: 'Cohort not found' };
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && cohort.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }
  if (!Array.isArray(updates) || updates.length === 0) return { success: false, message: 'No updates provided' };

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  const participantStatusCol = pHeaders.indexOf('status');
  const participantIdCol = pHeaders.indexOf('participantId');
  const rolloutCol = pHeaders.indexOf('rolloutId');

  const rowByParticipantId = {};
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][rolloutCol]) === String(rolloutId)) {
      rowByParticipantId[String(pData[i][participantIdCol])] = i + 1;
    }
  }

  let successCount = 0;
  let errorCount = 0;
  let hasChanges = false;
  updates.forEach(update => {
    if (!update || !update.participantId || !update.status) { errorCount++; return; }
    if (CONFIG.PARTICIPANT_STATUSES.indexOf(update.status) === -1) { errorCount++; return; }
    const row = rowByParticipantId[String(update.participantId)];
    if (!row) { errorCount++; return; }
    pData[row - 1][participantStatusCol] = update.status;
    hasChanges = true;
    successCount++;
  });
  if (hasChanges) {
    participantSnapshot.sheet.getRange(1, 1, pData.length, pHeaders.length).setValues(pData);
    invalidateSheetSnapshot('Participants');
  }

  logActivity(currentUser.userId, currentUser.fullName, 'BULK_UPDATE_PARTICIPANT_STATUS', 'cohort', rolloutId,
    'Updated statuses for ' + successCount + ' participants');
  return {
    success: errorCount === 0,
    message: 'Updated ' + successCount + ' participants' + (errorCount > 0 ? (' (' + errorCount + ' errors)') : '')
  };
}

/**
 * Bulk update checklist items
 */
function bulkUpdateChecklist(token, participantId, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < updates.length; i++) {
    const update = updates[i];
    const result = updateChecklistItem(token, update.checklistId, {
      status: update.status,
      notes: update.notes
    });

    if (result.success) {
      successCount++;
    } else {
      errorCount++;
      if (result.message && result.message.toLowerCase().includes('unauthorized')) {
        return { success: false, message: result.message };
      }
    }
  }

  return {
    success: true,
    message: `Updated ${successCount} items, ${errorCount} errors`
  };
}

/**
 * Get checklist statuses for a specific instrument within a cohort
 */
function getInstrumentChecklistForRollout(token, rolloutId, instrumentNumber) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const cohort = getRolloutById(rolloutId);
  if (!cohort) {
    return { success: false, message: 'Cohort not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && cohort.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = getGlobalProtocolItems().find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const checklistData = checklistSheet.getDataRange().getValues();
  const cHeaders = checklistData[0];

  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const statusCol = cHeaders.indexOf('status');
  const checklistIdCol = cHeaders.indexOf('checklistId');

  const checklistMap = {};
  for (let i = 1; i < checklistData.length; i++) {
    if (String(checklistData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = checklistData[i][checklistParticipantCol];
      checklistMap[pid] = {
        checklistId: checklistData[i][checklistIdCol],
        status: checklistData[i][statusCol],
        completedDate: checklistData[i][cHeaders.indexOf('completedDate')],
        completedBy: checklistData[i][cHeaders.indexOf('completedBy')]
      };
    }
  }

  const participants = [];
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] !== rolloutId) continue;

    const participantId = pData[i][pHeaders.indexOf('participantId')];
    const checklistInfo = checklistMap[participantId] || {};

    participants.push({
      participantId: participantId,
      fullName: pData[i][pHeaders.indexOf('fullName')],
      site: pData[i][pHeaders.indexOf('site')],
      schoolName: pData[i][pHeaders.indexOf('schoolName')],
      status: checklistInfo.status || 'not_started',
      checklistId: checklistInfo.checklistId || '',
      completedDate: checklistInfo.completedDate || '',
      completedBy: checklistInfo.completedBy || ''
    });
  }

  return {
    success: true,
    instrumentName: instrument.name,
    rolloutName: cohort.schoolName,
    site: cohort.site,
    participants: participants
  };
}

/**
 * Get checklist data links for a specific instrument within a cohort
 */
function getInstrumentLinksForRollout(token, rolloutId, instrumentNumber) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const cohort = getRolloutById(rolloutId);
  if (!cohort) {
    return { success: false, message: 'Cohort not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && cohort.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = getGlobalProtocolItems().find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const checklistData = checklistSheet.getDataRange().getValues();
  const cHeaders = checklistData[0];

  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const dataLinkCol = cHeaders.indexOf('dataLink');

  const checklistMap = {};
  for (let i = 1; i < checklistData.length; i++) {
    if (String(checklistData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = checklistData[i][checklistParticipantCol];
      checklistMap[pid] = {
        dataLink: checklistData[i][dataLinkCol] || ''
      };
    }
  }

  const participants = [];
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] !== rolloutId) continue;

    const participantId = pData[i][pHeaders.indexOf('participantId')];
    const checklistInfo = checklistMap[participantId] || {};

    participants.push({
      participantId: participantId,
      fullName: pData[i][pHeaders.indexOf('fullName')],
      site: pData[i][pHeaders.indexOf('site')],
      schoolName: pData[i][pHeaders.indexOf('schoolName')],
      dataLink: checklistInfo.dataLink || ''
    });
  }

  return {
    success: true,
    instrumentName: instrument.name,
    rolloutName: cohort.schoolName,
    site: cohort.site,
    participants: participants
  };
}

/**
 * Bulk update an instrument across participants in a cohort
 */
function bulkUpdateInstrumentStatus(token, rolloutId, instrumentNumber, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const cohort = getRolloutById(rolloutId);
  if (!cohort) {
    return { success: false, message: 'Cohort not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && cohort.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = getGlobalProtocolItems().find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  if (!Array.isArray(updates) || updates.length === 0) {
    return { success: false, message: 'No updates provided' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const participantsInRollout = {};

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] === rolloutId) {
      participantsInRollout[pData[i][pHeaders.indexOf('participantId')]] = true;
    }
  }

  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cData = checklistSnapshot.data;
  const cHeaders = checklistSnapshot.headers;
  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const statusCol = cHeaders.indexOf('status');
  const completedDateCol = cHeaders.indexOf('completedDate');
  const completedByCol = cHeaders.indexOf('completedBy');

  const checklistMap = {};
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = cData[i][checklistParticipantCol];
      checklistMap[pid] = i + 1;
    }
  }

  let successCount = 0;
  let errorCount = 0;
  const touchedParticipants = {};

  let hasChanges = false;
  updates.forEach(update => {
    if (!participantsInRollout[update.participantId]) {
      errorCount++;
      return;
    }

    if (CONFIG.CHECKLIST_STATUSES.indexOf(update.status) === -1) {
      errorCount++;
      return;
    }

    const row = checklistMap[update.participantId];
    if (!row) {
      errorCount++;
      return;
    }
    cData[row - 1][statusCol] = update.status;
    if (update.status === 'completed') {
      cData[row - 1][completedDateCol] = new Date().toISOString();
      cData[row - 1][completedByCol] = currentUser.fullName;
    } else {
      cData[row - 1][completedDateCol] = '';
      cData[row - 1][completedByCol] = '';
    }
    hasChanges = true;
    touchedParticipants[update.participantId] = true;
    successCount++;
  });

  if (hasChanges) {
    checklistSnapshot.sheet.getRange(1, 1, cData.length, cHeaders.length).setValues(cData);
    invalidateSheetSnapshot('Checklist');
  }

  Object.keys(touchedParticipants).forEach(pid => {
    updateParticipantCompletion(pid);
    syncParticipantStatusFromChecklist(pid);
  });

  logActivity(currentUser.userId, currentUser.fullName, 'BULK_UPDATE_INSTRUMENT', 'checklist', instrumentNumber,
    'Updated ' + successCount + ' ' + instrument.name + ' records for cohort ' + cohort.schoolName);

  return {
    success: errorCount === 0,
    message: 'Updated ' + successCount + ' participants' + (errorCount > 0 ? (' (' + errorCount + ' errors)') : '')
  };
}

/**
 * Bulk update data links for an instrument across participants in a cohort
 */
function bulkUpdateInstrumentLinks(token, rolloutId, instrumentNumber, updates) {
  const currentUser = validateSession(token);
  if (!currentUser || currentUser.role === 'viewer') {
    return { success: false, message: 'Unauthorized' };
  }

  const cohort = getRolloutById(rolloutId);
  if (!cohort) {
    return { success: false, message: 'Cohort not found' };
  }

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && cohort.site !== currentUser.site) {
    return { success: false, message: 'Unauthorized for this site' };
  }

  const instrument = getGlobalProtocolItems().find(inst => String(inst.number) === String(instrumentNumber));
  if (!instrument) {
    return { success: false, message: 'Instrument not found' };
  }

  if (!Array.isArray(updates) || updates.length === 0) {
    return { success: false, message: 'No updates provided' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const participantsInRollout = {};

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('rolloutId')] === rolloutId) {
      participantsInRollout[pData[i][pHeaders.indexOf('participantId')]] = true;
    }
  }

  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cData = checklistSnapshot.data;
  const cHeaders = checklistSnapshot.headers;
  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const dataLinkCol = cHeaders.indexOf('dataLink');

  const checklistMap = {};
  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][instrumentCol]) === String(instrumentNumber)) {
      const pid = cData[i][checklistParticipantCol];
      checklistMap[pid] = i + 1;
    }
  }

  let successCount = 0;
  let errorCount = 0;

  let hasChanges = false;
  updates.forEach(update => {
    if (!participantsInRollout[update.participantId]) {
      errorCount++;
      return;
    }

    if (update.dataLink === undefined) {
      errorCount++;
      return;
    }

    const row = checklistMap[update.participantId];
    if (!row) {
      errorCount++;
      return;
    }
    cData[row - 1][dataLinkCol] = update.dataLink;
    hasChanges = true;
    successCount++;
  });
  if (hasChanges) {
    checklistSnapshot.sheet.getRange(1, 1, cData.length, cHeaders.length).setValues(cData);
  }

  logActivity(currentUser.userId, currentUser.fullName, 'BULK_UPDATE_INSTRUMENT_LINK', 'checklist', instrumentNumber,
    'Updated ' + successCount + ' ' + instrument.name + ' links for cohort ' + cohort.schoolName);

  return {
    success: errorCount === 0,
    message: 'Updated ' + successCount + ' participants' + (errorCount > 0 ? (' (' + errorCount + ' errors)') : '')
  };
}

/**
 * Update participant's completion percentage
 */
function updateParticipantCompletion(participantId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  // Count completed items
  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];

  let total = 0;
  let completed = 0;

  for (let i = 1; i < cData.length; i++) {
    if (cData[i][cHeaders.indexOf('participantId')] === participantId) {
      total++;
      if (cData[i][cHeaders.indexOf('status')] === 'completed') {
        completed++;
      }
    }
  }

  const percentage = total > 0 ? Math.round((completed / total) * 100) : 0;

  // Update participant record
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;

  for (let i = 1; i < pData.length; i++) {
    if (pData[i][pHeaders.indexOf('participantId')] === participantId) {
      setCellAsPlainText(
        participantsSheet,
        i + 1,
        pHeaders.indexOf('completionPercentage') + 1,
        percentage
      );
      break;
    }
  }

  return percentage;
}

function buildLiveCompletionMap(participantIds) {
  const map = {};
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  const cHeaders = checklistSnapshot.headers;
  const cData = checklistSnapshot.data;
  if (pData.length < 2 || cData.length < 2) return map;

  const requestedIds = new Set((participantIds || []).map(id => String(id)));
  const filterEnabled = requestedIds.size > 0;
  const participantIdCol = pHeaders.indexOf('participantId');
  const participantRolloutCol = pHeaders.indexOf('rolloutId');
  const participantRolloutMap = {};
  const rolloutAllowedNumbersMap = {};

  for (let i = 1; i < pData.length; i++) {
    const pid = String(pData[i][participantIdCol] || '');
    if (!pid || (filterEnabled && !requestedIds.has(pid))) continue;
    const rolloutId = String(pData[i][participantRolloutCol] || '');
    participantRolloutMap[pid] = rolloutId;
    if (!rolloutAllowedNumbersMap[rolloutId]) {
      const allowed = {};
      getRolloutProtocolItemsInternal(rolloutId).forEach(item => { allowed[String(item.number)] = true; });
      rolloutAllowedNumbersMap[rolloutId] = allowed;
    }
  }

  const checklistParticipantCol = cHeaders.indexOf('participantId');
  const statusCol = cHeaders.indexOf('status');
  const numberCol = cHeaders.indexOf('instrumentNumber');
  for (let i = 1; i < cData.length; i++) {
    const pid = String(cData[i][checklistParticipantCol] || '');
    if (!pid || (filterEnabled && !requestedIds.has(pid))) continue;
    const rolloutId = participantRolloutMap[pid] || '';
    const allowedMap = rolloutAllowedNumbersMap[rolloutId] || {};
    const itemNumber = String(cData[i][numberCol]);
    if (!allowedMap[itemNumber]) continue;
    if (!map[pid]) map[pid] = { total: 0, completed: 0, percentage: 0 };
    map[pid].total++;
    if (cData[i][statusCol] === 'completed') map[pid].completed++;
  }
  Object.keys(map).forEach(pid => {
    const entry = map[pid];
    entry.percentage = entry.total > 0 ? Math.round((entry.completed / entry.total) * 100) : 0;
  });
  return map;
}

// ============================================
// DASHBOARD & STATISTICS FUNCTIONS
// ============================================

/**
 * Get dashboard statistics
 */
function getDashboardStats(token, siteFilter, rolloutFilter) {
  const currentUser = validateSession(token);
  // Allow public access for basic stats, but filter for authenticated users

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  const stats = {
    totalParticipants: 0,
    activeParticipants: 0,
    completedParticipants: 0,
    withdrawnParticipants: 0,
    ugaParticipants: 0,
    missouriParticipants: 0,
    activeRollouts: 0,
    overallCompletion: 0,
    instrumentStats: [],
    recentActivity: [],
    currentRollout: null
  };

  if (!participantsSheet) return stats;

  // Get participants data
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;

  let totalCompletion = 0;
  let filteredParticipantIds = [];

  const dashboardSiteCol = pHeaders.indexOf('site');
  const dashboardStatusCol = pHeaders.indexOf('status');
  const dashboardRolloutIdCol = pHeaders.indexOf('rolloutId');
  const dashboardParticipantIdCol = pHeaders.indexOf('participantId');

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][dashboardSiteCol];
    const status = pData[i][dashboardStatusCol];
    const rolloutId = pData[i][dashboardRolloutIdCol];
    const participantId = pData[i][dashboardParticipantIdCol];

    // Apply site filter
    if (siteFilter && siteFilter !== 'All' && site !== siteFilter) {
      continue;
    }

    // Apply cohort filter
    if (rolloutFilter && rolloutFilter !== 'All' && rolloutId !== rolloutFilter) {
      continue;
    }

    filteredParticipantIds.push(participantId);
    stats.totalParticipants++;

    if (status === 'active') stats.activeParticipants++;
    if (status === 'completed') stats.completedParticipants++;
    if (status === 'withdrawn') stats.withdrawnParticipants++;

    if (site === 'UGA') stats.ugaParticipants++;
    if (site === 'Missouri') stats.missouriParticipants++;
  }

  const filteredParticipantIdSet = new Set(filteredParticipantIds.map(id => String(id)));
  const liveCompletionMap = filteredParticipantIds.length > 0
    ? buildLiveCompletionMap(filteredParticipantIds)
    : {};
  filteredParticipantIds.forEach(pid => {
    totalCompletion += (liveCompletionMap[pid] || {}).percentage || 0;
  });

  stats.overallCompletion = stats.totalParticipants > 0
    ? Math.round(totalCompletion / stats.totalParticipants)
    : 0;

  // Get cohorts data
  if (rolloutsSheet) {
    const rolloutSnapshot = getSheetSnapshot('StudyRollouts');
    const rData = rolloutSnapshot.data;
    const rHeaders = rolloutSnapshot.headers;

    for (let i = 1; i < rData.length; i++) {
      const rSite = rData[i][rHeaders.indexOf('site')];
      const rStatus = rData[i][rHeaders.indexOf('status')];
      const rId = rData[i][rHeaders.indexOf('rolloutId')];

      // If a specific cohort is selected, get its details
      if (rolloutFilter && rolloutFilter !== 'All' && rId === rolloutFilter) {
        stats.currentRollout = {
          rolloutId: rId,
          site: rSite,
          schoolName: rData[i][rHeaders.indexOf('schoolName')],
          period: rData[i][rHeaders.indexOf('period')],
          year: rData[i][rHeaders.indexOf('year')],
          status: rStatus
        };
      }

      if (siteFilter && siteFilter !== 'All' && rSite !== siteFilter) {
        continue;
      }

      if (rStatus === 'active') stats.activeRollouts++;
    }
  }

  // Get instrument completion stats (only for filtered participants)
  if (checklistSheet) {
    const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
    const cData = checklistSnapshot.data;
    const cHeaders = checklistSnapshot.headers;

    const protocolItems = getProtocolItemsForFilter(rolloutFilter);
    const instrumentCounts = {};
    protocolItems.forEach(inst => {
      instrumentCounts[inst.number] = { total: 0, completed: 0, name: inst.name };
    });

    const instrumentNumberCol = cHeaders.indexOf('instrumentNumber');
    const instrumentStatusCol = cHeaders.indexOf('status');
    const checklistParticipantCol = cHeaders.indexOf('participantId');

    for (let i = 1; i < cData.length; i++) {
      const instNum = cData[i][instrumentNumberCol];
      const instStatus = cData[i][instrumentStatusCol];
      const checklistParticipantId = cData[i][checklistParticipantCol];

      // Only count checklist items for filtered participants
      if (!filteredParticipantIdSet.has(String(checklistParticipantId))) {
        continue;
      }

      if (instrumentCounts[instNum]) {
        instrumentCounts[instNum].total++;
        if (instStatus === 'completed') {
          instrumentCounts[instNum].completed++;
        }
      }
    }

    stats.instrumentStats = protocolItems.map(inst => ({
      number: inst.number,
      name: inst.name,
      category: inst.category,
      total: instrumentCounts[inst.number].total,
      completed: instrumentCounts[inst.number].completed,
      percentage: instrumentCounts[inst.number].total > 0
        ? Math.round((instrumentCounts[inst.number].completed / instrumentCounts[inst.number].total) * 100)
        : 0
    }));

    const instrumentCompleted = stats.instrumentStats.reduce((sum, inst) => sum + inst.completed, 0);
    const instrumentTotal = stats.instrumentStats.reduce((sum, inst) => sum + inst.total, 0);
    if (instrumentTotal > 0) {
      stats.overallCompletion = Math.round((instrumentCompleted / instrumentTotal) * 100);
    }
  }

  return stats;
}

/**
 * Get analytics overview across cohorts and instruments
 */
function getAnalyticsOverview(token, siteFilter, rolloutFilter) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const effectiveSite = (currentUser.role === 'facilitator' && currentUser.site !== 'All')
    ? currentUser.site
    : (siteFilter || 'All');
  const effectiveRollout = rolloutFilter || 'All';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet) {
    return {
      success: true,
      overview: {
        totalParticipants: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        averageCompletion: 0,
        activeRollouts: 0,
        rolloutCount: 0,
        instrumentsCompleted: 0,
        instrumentsTotal: 0,
        siteFilter: effectiveSite,
        rolloutFilter: effectiveRollout
      },
      cohorts: [],
      instrumentStats: []
    };
  }

  const rolloutEntries = {};
  if (rolloutsSheet) {
    const rolloutSnapshot = getSheetSnapshot('StudyRollouts');
    const rData = rolloutSnapshot.data;
    const rHeaders = rolloutSnapshot.headers;

    for (let i = 1; i < rData.length; i++) {
      const rolloutId = rData[i][rHeaders.indexOf('rolloutId')];
      const rolloutSite = rData[i][rHeaders.indexOf('site')];

      if (effectiveSite !== 'All' && rolloutSite !== effectiveSite) continue;
      if (effectiveRollout !== 'All' && rolloutId !== effectiveRollout) continue;

      rolloutEntries[rolloutId] = {
        rolloutId: rolloutId,
        site: rolloutSite,
        schoolName: rData[i][rHeaders.indexOf('schoolName')],
        period: rData[i][rHeaders.indexOf('period')],
        year: rData[i][rHeaders.indexOf('year')],
        status: rData[i][rHeaders.indexOf('status')],
        participantCount: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        completionSum: 0,
        averageCompletion: 0
      };
    }
  }

  const overview = {
    totalParticipants: 0,
    activeParticipants: 0,
    completedParticipants: 0,
    withdrawnParticipants: 0,
    averageCompletion: 0,
    activeRollouts: 0,
    rolloutCount: Object.keys(rolloutEntries).length,
    instrumentsCompleted: 0,
    instrumentsTotal: 0,
    siteFilter: effectiveSite,
    rolloutFilter: effectiveRollout
  };

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const includedParticipantIds = [];
  const participantRolloutMap = {};
  const participantNameMap = {};
  const analyticsSiteCol = pHeaders.indexOf('site');
  const analyticsRolloutIdCol = pHeaders.indexOf('rolloutId');
  const analyticsStatusCol = pHeaders.indexOf('status');
  const analyticsParticipantIdCol = pHeaders.indexOf('participantId');
  const analyticsFullNameCol = pHeaders.indexOf('fullName');
  const analyticsSchoolNameCol = pHeaders.indexOf('schoolName');
  const analyticsPeriodCol = pHeaders.indexOf('period');
  const analyticsYearCol = pHeaders.indexOf('year');
  let completionSum = 0;

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][analyticsSiteCol];
    const rolloutId = pData[i][analyticsRolloutIdCol];
    const status = pData[i][analyticsStatusCol];
    const participantId = pData[i][analyticsParticipantIdCol];

    if (effectiveSite !== 'All' && site !== effectiveSite) continue;
    if (effectiveRollout !== 'All' && rolloutId !== effectiveRollout) continue;

    includedParticipantIds.push(participantId);
    participantRolloutMap[participantId] = rolloutId || 'Unassigned';
    participantNameMap[participantId] = pData[i][analyticsFullNameCol] || 'Unknown Participant';
    overview.totalParticipants++;
    overview.activeParticipants += status === 'active' ? 1 : 0;
    overview.completedParticipants += status === 'completed' ? 1 : 0;
    overview.withdrawnParticipants += status === 'withdrawn' ? 1 : 0;

    const key = rolloutId || 'Unassigned';
    if (!rolloutEntries[key]) {
      rolloutEntries[key] = {
        rolloutId: rolloutId || 'Unassigned',
        site: site,
        schoolName: rolloutId ? pData[i][analyticsSchoolNameCol] : 'No cohort assigned',
        period: pData[i][analyticsPeriodCol],
        year: pData[i][analyticsYearCol],
        status: rolloutId ? 'active' : 'unassigned',
        participantCount: 0,
        activeParticipants: 0,
        completedParticipants: 0,
        withdrawnParticipants: 0,
        completionSum: 0,
        averageCompletion: 0
      };
    }

    const rolloutStats = rolloutEntries[key];
    rolloutStats.participantCount++;
    rolloutStats.activeParticipants += status === 'active' ? 1 : 0;
    rolloutStats.completedParticipants += status === 'completed' ? 1 : 0;
    rolloutStats.withdrawnParticipants += status === 'withdrawn' ? 1 : 0;
    // completionSum populated from live checklist map below
  }

  const includedParticipantIdSet = new Set(includedParticipantIds.map(id => String(id)));
  const liveCompletionMap = includedParticipantIds.length > 0
    ? buildLiveCompletionMap(includedParticipantIds)
    : {};
  includedParticipantIds.forEach(participantId => {
    const livePct = (liveCompletionMap[participantId] || {}).percentage || 0;
    completionSum += livePct;
    const rolloutKey = participantRolloutMap[participantId] || 'Unassigned';
    if (rolloutEntries[rolloutKey]) {
      rolloutEntries[rolloutKey].completionSum += livePct;
    }
  });

overview.averageCompletion = overview.totalParticipants > 0
    ? Math.round(completionSum / overview.totalParticipants)
    : 0;

  Object.keys(rolloutEntries).forEach(key => {
    const entry = rolloutEntries[key];
    if (entry.participantCount > 0) {
      entry.averageCompletion = Math.round(entry.completionSum / entry.participantCount);
    }
    if (entry.status === 'active') {
      overview.activeRollouts++;
    }
  });

  // Instrument progress for included participants
  let instrumentStats = [];
  if (checklistSheet && includedParticipantIds.length > 0) {
    const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
    const cData = checklistSnapshot.data;
    const cHeaders = checklistSnapshot.headers;
    const counts = {};
    const pendingByInstrument = {};
    const pendingSets = {};

    const protocolItems = getProtocolItemsForFilter(effectiveRollout);
    protocolItems.forEach(inst => {
      counts[inst.number] = { total: 0, completed: 0, name: inst.name, category: inst.category };
      pendingByInstrument[inst.number] = [];
      pendingSets[inst.number] = new Set();
    });

    const checklistParticipantCol = cHeaders.indexOf('participantId');
    const checklistInstrumentNumberCol = cHeaders.indexOf('instrumentNumber');
    const checklistStatusCol = cHeaders.indexOf('status');

    for (let i = 1; i < cData.length; i++) {
      const participantId = cData[i][checklistParticipantCol];
      if (!includedParticipantIdSet.has(String(participantId))) continue;

      const instrumentNumber = cData[i][checklistInstrumentNumberCol];
      const status = cData[i][checklistStatusCol];

      if (counts[instrumentNumber]) {
        counts[instrumentNumber].total++;
        if (status === 'completed') {
          counts[instrumentNumber].completed++;
        } else {
          if (!pendingSets[instrumentNumber].has(participantId)) {
            pendingSets[instrumentNumber].add(participantId);
            pendingByInstrument[instrumentNumber].push({
              participantId: participantId,
              name: participantNameMap[participantId] || 'Unknown Participant',
              status: status || 'not_started'
            });
          }
        }
      }
    }

    Object.keys(pendingByInstrument).forEach(key => {
      pendingByInstrument[key].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    });

    instrumentStats = protocolItems.map(inst => ({
      number: inst.number,
      name: inst.name,
      category: inst.category,
      total: counts[inst.number].total,
      completed: counts[inst.number].completed,
      percentage: counts[inst.number].total > 0
        ? Math.round((counts[inst.number].completed / counts[inst.number].total) * 100)
        : 0,
      pendingParticipants: pendingByInstrument[inst.number]
    }));

    overview.instrumentsCompleted = instrumentStats.reduce((sum, inst) => sum + inst.completed, 0);
    overview.instrumentsTotal = instrumentStats.reduce((sum, inst) => sum + inst.total, 0);
    if (overview.instrumentsTotal > 0) {
      overview.averageCompletion = Math.round((overview.instrumentsCompleted / overview.instrumentsTotal) * 100);
    }
  }

  const cohorts = Object.values(rolloutEntries).sort((a, b) => {
    if (a.site === b.site) {
      return (b.year || '').toString().localeCompare((a.year || '').toString());
    }
    return a.site.localeCompare(b.site);
  });

  overview.rolloutCount = cohorts.length;

  return {
    success: true,
    overview: overview,
    cohorts: cohorts,
    instrumentStats: instrumentStats
  };
}

/**
 * Get public dashboard statistics (no authentication required)
 */
function getPublicDashboardStats(siteFilter, rolloutFilter) {
  return getDashboardStats(null, siteFilter || 'All', rolloutFilter || 'All');
}

/**
 * Get public landing page snapshot metrics for the research dashboard.
 */
function getPublicLandingSnapshot(siteFilter, rolloutFilter) {
  const effectiveSite = siteFilter || 'All';
  const effectiveCohort = rolloutFilter || 'All';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const attendanceSheet = ss.getSheetByName('SessionAttendance');
  const checklistSheet = ss.getSheetByName('Checklist');

  const cohortsById = {};
  const cohorts = [];

  if (rolloutsSheet) {
    const rolloutSnapshot = getSheetSnapshot('StudyRollouts');
    const rData = rolloutSnapshot.data;
    const rHeaders = rolloutSnapshot.headers;
    for (let i = 1; i < rData.length; i++) {
      const rolloutId = rData[i][rHeaders.indexOf('rolloutId')];
      const site = rData[i][rHeaders.indexOf('site')];
      if (effectiveSite !== 'All' && site !== effectiveSite) continue;
      const cohort = {
        rolloutId: rolloutId,
        site: site,
        schoolName: rData[i][rHeaders.indexOf('schoolName')],
        period: rData[i][rHeaders.indexOf('period')],
        year: rData[i][rHeaders.indexOf('year')],
        status: rData[i][rHeaders.indexOf('status')]
      };
      cohorts.push(cohort);
      cohortsById[rolloutId] = cohort;
    }
  }

  const currentCohort = effectiveCohort !== 'All' ? cohortsById[effectiveCohort] : null;
  const contextLine = buildLandingContextLine(currentCohort, effectiveSite, effectiveCohort);

  const participants = [];
  const participantsBySite = { UGA: [], Missouri: [] };
  const participantsByCohort = {};
  let totalEnrolled = 0;
  let activeEnrollments = 0;

  if (participantsSheet) {
    const pData = participantsSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    for (let i = 1; i < pData.length; i++) {
      const participant = {
        participantId: pData[i][pHeaders.indexOf('participantId')],
        site: pData[i][pHeaders.indexOf('site')],
        rolloutId: pData[i][pHeaders.indexOf('rolloutId')],
        status: pData[i][pHeaders.indexOf('status')],
        completionPercentage: Number(pData[i][pHeaders.indexOf('completionPercentage')]) || 0
      };

      participantsBySite[participant.site] = participantsBySite[participant.site] || [];
      participantsBySite[participant.site].push(participant);

      if (!participantsByCohort[participant.rolloutId]) {
        participantsByCohort[participant.rolloutId] = [];
      }
      participantsByCohort[participant.rolloutId].push(participant);

      if (effectiveSite !== 'All' && participant.site !== effectiveSite) continue;
      if (effectiveCohort !== 'All' && participant.rolloutId !== effectiveCohort) continue;

      participants.push(participant);
      totalEnrolled++;
      if (participant.status === 'active') {
        activeEnrollments++;
      }
    }
  }

  const sessionLookup = {};
  const sessionsByCohort = {};

  if (sessionsSheet) {
    const sData = sessionsSheet.getDataRange().getValues();
    const sHeaders = sData[0];
    for (let i = 1; i < sData.length; i++) {
      const rolloutId = sData[i][sHeaders.indexOf('rolloutId')];
      if (effectiveSite !== 'All') {
        const cohort = cohortsById[rolloutId];
        if (!cohort) continue;
      }
      if (effectiveCohort !== 'All' && rolloutId !== effectiveCohort) continue;

      const session = {
        sessionId: sData[i][sHeaders.indexOf('sessionId')],
        rolloutId: rolloutId,
        sessionNumber: sData[i][sHeaders.indexOf('sessionNumber')],
        sessionDate: normalizeSessionDateValue(sData[i][sHeaders.indexOf('sessionDate')]),
        sessionName: sData[i][sHeaders.indexOf('sessionName')],
        status: sData[i][sHeaders.indexOf('status')]
      };
      sessionsByCohort[rolloutId] = sessionsByCohort[rolloutId] || [];
      sessionsByCohort[rolloutId].push(session);
      sessionLookup[session.sessionId] = session;
    }
  }

  const attendanceMap = {};
  if (attendanceSheet) {
    const aData = attendanceSheet.getDataRange().getValues();
    const aHeaders = aData[0];
    for (let i = 1; i < aData.length; i++) {
      const sessionId = aData[i][aHeaders.indexOf('sessionId')];
      if (!sessionLookup[sessionId]) continue;
      const participantId = aData[i][aHeaders.indexOf('participantId')];
      if (!attendanceMap[sessionId]) {
        attendanceMap[sessionId] = {};
      }
      attendanceMap[sessionId][participantId] = aData[i][aHeaders.indexOf('status')];
    }
  }

  const attendanceSummary = buildAttendanceSummary(
    participants,
    sessionsByCohort,
    attendanceMap,
    effectiveCohort
  );

  const protocolSummary = buildProtocolSummary(
    participants,
    checklistSheet
  );

  const attentionItems = buildAttentionItems(
    attendanceSummary,
    protocolSummary,
    participants,
    effectiveCohort
  );

  const comparison = buildComparisonSummary(
    effectiveSite,
    effectiveCohort,
    participantsBySite,
    participantsByCohort,
    sessionsByCohort,
    attendanceMap,
    checklistSheet,
    cohorts
  );

  const operations = buildOperationsSummary(
    effectiveSite,
    effectiveCohort,
    cohorts,
    participantsBySite,
    participantsByCohort,
    sessionsByCohort,
    attendanceMap,
    checklistSheet
  );

  const aboutSummary = buildAboutSummary(
    effectiveSite,
    effectiveCohort,
    totalEnrolled,
    cohorts,
    participantsBySite
  );

  return {
    success: true,
    snapshot: {
      contextLine: contextLine,
      meta: {
        generatedAt: new Date().toISOString(),
        exportReady: totalEnrolled > 0
      },
      kpis: {
        totalEnrolled: totalEnrolled,
        activeEnrollments: activeEnrollments,
        averageAttendance: attendanceSummary.averageAttendance,
        participationStatus: attendanceSummary.participationStatus,
        atRiskParticipants: attendanceSummary.atRiskParticipants,
        protocolCompletionRate: protocolSummary.completionRate,
        sessionsCompleted: attendanceSummary.sessionsCompleted,
        totalSessions: attendanceSummary.totalSessions
      },
      attendance: attendanceSummary,
      protocol: protocolSummary,
      attention: attentionItems,
      comparison: comparison,
      operations: operations,
      about: aboutSummary
    }
  };
}


function buildOperationsSummary(
  effectiveSite,
  effectiveCohort,
  cohorts,
  participantsBySite,
  participantsByCohort,
  sessionsByCohort,
  attendanceMap,
  checklistSheet
) {
  const today = getDateOnly(new Date());
  const visibleCohorts = cohorts.filter(cohort => effectiveCohort === 'All' || String(cohort.rolloutId) === String(effectiveCohort));
  const lifecycleCounts = { upcoming: 0, inProgress: 0, completed: 0, unscheduled: 0 };

  const cohortOperations = visibleCohorts.map(cohort => {
    const participants = participantsByCohort[cohort.rolloutId] || [];
    const sessions = (sessionsByCohort[cohort.rolloutId] || []).slice().sort(sortSessionsByNumberAndDate);
    const lifecycle = resolveCohortLifecycle(cohort, sessions, today);
    lifecycleCounts[lifecycle.key] = (lifecycleCounts[lifecycle.key] || 0) + 1;

    const attendanceSummary = buildAttendanceSummary(participants, { [cohort.rolloutId]: sessions }, attendanceMap, 'All');
    const protocolSummary = buildProtocolSummary(participants, checklistSheet);
    const sessionProtocolMap = buildSessionProtocolCompletionMap(sessions, participants, checklistSheet);
    const sessionProgress = sessions.map(session => buildSessionOperationsRow(
      session,
      participants,
      attendanceMap,
      today,
      protocolSummary,
      sessionProtocolMap[session.sessionId]
    ));
    const nextSession = sessionProgress.find(session => session.timing === 'today' || session.timing === 'upcoming') || null;
    const completedOrLastSessions = sessionProgress
      .filter(session => session.sessionDate)
      .slice()
      .sort((a, b) => compareDateStrings(b.sessionDate, a.sessionDate));
    const lastSession = completedOrLastSessions[0] || null;

    return {
      rolloutId: cohort.rolloutId,
      site: cohort.site,
      label: formatCohortDisplayLabel(cohort),
      schoolName: cohort.schoolName,
      period: cohort.period,
      year: cohort.year,
      sourceStatus: cohort.status,
      lifecycle: lifecycle,
      participantCount: participants.length,
      activeParticipants: participants.filter(p => p.status === 'active').length,
      completedParticipants: participants.filter(p => p.status === 'completed').length,
      withdrawnParticipants: participants.filter(p => p.status === 'withdrawn').length,
      averageAttendance: attendanceSummary.averageAttendance,
      protocolCompletion: protocolSummary.completionRate,
      sessionsCompleted: sessionProgress.filter(session => session.timing === 'completed').length,
      totalSessions: sessionProgress.length,
      nextSession: nextSession,
      lastSession: lastSession,
      lastSessionDate: lastSession ? lastSession.sessionDate : '',
      sessions: sessionProgress
    };
  });

  const sites = ['UGA', 'Missouri'].map(site => {
    const siteCohorts = cohortOperations.filter(cohort => cohort.site === site);
    const participants = (participantsBySite[site] || []).filter(participant => effectiveCohort === 'All' || String(participant.rolloutId) === String(effectiveCohort));
    const attendanceSummary = buildAttendanceSummary(participants, sessionsByCohort, attendanceMap, 'All');
    const protocolSummary = buildProtocolSummary(participants, checklistSheet);
    return {
      site: site,
      cohortCount: siteCohorts.length,
      participantCount: participants.length,
      activeCohorts: siteCohorts.filter(cohort => cohort.lifecycle.key === 'inProgress').length,
      upcomingCohorts: siteCohorts.filter(cohort => cohort.lifecycle.key === 'upcoming').length,
      completedCohorts: siteCohorts.filter(cohort => cohort.lifecycle.key === 'completed').length,
      averageAttendance: attendanceSummary.averageAttendance,
      protocolCompletion: protocolSummary.completionRate,
      cohorts: siteCohorts
    };
  }).filter(site => effectiveSite === 'All' || site.site === effectiveSite);

  const upcomingCohorts = cohortOperations
    .filter(cohort => cohort.lifecycle.key === 'upcoming')
    .sort((a, b) => compareDateStrings((a.nextSession || {}).sessionDate, (b.nextSession || {}).sessionDate))
    .slice(0, 6);

  const activeCohorts = cohortOperations
    .filter(cohort => cohort.lifecycle.key === 'inProgress')
    .sort((a, b) => a.site.localeCompare(b.site) || a.label.localeCompare(b.label));

  const recentCompletedCohorts = activeCohorts.length > 0 ? [] : cohortOperations
    .filter(cohort => cohort.lifecycle.key === 'completed')
    .sort((a, b) => compareDateStrings(b.lastSessionDate, a.lastSessionDate))
    .slice(0, 1);

  return {
    sites: sites,
    cohorts: cohortOperations,
    lifecycleCounts: lifecycleCounts,
    activeCohorts: activeCohorts,
    recentCompletedCohorts: recentCompletedCohorts,
    upcomingCohorts: upcomingCohorts,
    totalCohorts: cohortOperations.length
  };
}

function buildSessionOperationsRow(session, participants, attendanceMap, today, protocolSummary, sessionProtocolSummary) {
  let present = 0;
  let absent = 0;
  let excused = 0;
  let notMarked = 0;
  const records = attendanceMap[session.sessionId] || {};

  participants.forEach(participant => {
    const status = records[participant.participantId];
    if (status === 'present') present++;
    else if (status === 'absent') absent++;
    else if (status === 'excused') excused++;
    else notMarked++;
  });

  const participantCount = participants.length;
  const attendanceRate = participantCount > 0 ? Math.round((present / participantCount) * 100) : 0;
  const markedCount = present + absent + excused;
  const markedRate = participantCount > 0 ? Math.round((markedCount / participantCount) * 100) : 0;
  const sessionDateOnly = getDateOnly(parseSessionDate(session.sessionDate));
  const timing = resolveSessionTiming(session, sessionDateOnly, today);

  const protocolDaySummary = sessionProtocolSummary || { completed: 0, total: 0, completionRate: 0 };

  return {
    sessionId: session.sessionId,
    sessionNumber: session.sessionNumber,
    sessionName: session.sessionName || ('Session ' + (session.sessionNumber || '')),
    sessionDate: session.sessionDate,
    sourceStatus: session.status,
    timing: timing,
    present: present,
    absent: absent,
    excused: excused,
    notMarked: notMarked,
    markedCount: markedCount,
    participantCount: participantCount,
    attendanceRate: attendanceRate,
    markedRate: markedRate,
    protocolCompletion: protocolSummary.completionRate,
    protocolDayCompleted: protocolDaySummary.completed || 0,
    protocolDayTotal: protocolDaySummary.total || 0,
    protocolDayCompletionRate: protocolDaySummary.completionRate || 0,
    protocolDayItems: protocolDaySummary.items || []
  };
}

function buildSessionProtocolCompletionMap(sessions, participants, checklistSheet) {
  const summaries = {};
  (sessions || []).forEach(session => {
    summaries[session.sessionId] = { completed: 0, total: 0, completionRate: 0, items: [] };
  });

  if (!checklistSheet || !participants || participants.length === 0 || !sessions || sessions.length === 0) {
    return summaries;
  }

  const participantIds = new Set(participants.map(p => String(p.participantId)));
  const participantRolloutMap = {};
  const rolloutAllowedNumbersMap = {};
  participants.forEach(participant => {
    const participantId = String(participant.participantId);
    const rolloutId = String(participant.rolloutId || '');
    participantRolloutMap[participantId] = rolloutId;
    if (!rolloutAllowedNumbersMap[rolloutId]) {
      const allowed = {};
      getRolloutProtocolItemsInternal(rolloutId).forEach(item => {
        allowed[String(item.number)] = true;
      });
      rolloutAllowedNumbersMap[rolloutId] = allowed;
    }
  });

  const sessionIdsByDate = {};
  sessions.forEach(session => {
    const sessionDateOnly = getDateOnly(parseSessionDate(session.sessionDate));
    if (!sessionDateOnly) return;
    sessionIdsByDate[sessionDateOnly] = sessionIdsByDate[sessionDateOnly] || [];
    sessionIdsByDate[sessionDateOnly].push(session.sessionId);
  });

  const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
  const cData = checklistSnapshot.data;
  const cHeaders = checklistSnapshot.headers;
  const participantCol = cHeaders.indexOf('participantId');
  const instrumentCol = cHeaders.indexOf('instrumentNumber');
  const instrumentNameCol = cHeaders.indexOf('instrumentName');
  const statusCol = cHeaders.indexOf('status');
  const completedDateCol = cHeaders.indexOf('completedDate');
  const totalsByInstrument = {};
  const namesByInstrument = {};
  const completedBySession = {};

  Object.keys(summaries).forEach(sessionId => {
    completedBySession[sessionId] = {};
  });

  for (let i = 1; i < cData.length; i++) {
    const participantId = String(cData[i][participantCol]);
    if (!participantIds.has(participantId)) continue;

    const rolloutId = participantRolloutMap[participantId] || '';
    const allowedMap = rolloutAllowedNumbersMap[rolloutId] || {};
    const instrumentNumber = String(cData[i][instrumentCol]);
    if (!allowedMap[instrumentNumber]) continue;

    const instrumentName = cData[i][instrumentNameCol] || ('Item ' + instrumentNumber);
    namesByInstrument[instrumentNumber] = instrumentName;
    totalsByInstrument[instrumentNumber] = (totalsByInstrument[instrumentNumber] || 0) + 1;

    const status = cData[i][statusCol];
    if (status !== 'completed') continue;

    const completedDateOnly = getDateOnly(new Date(cData[i][completedDateCol]));
    const matchingSessionIds = sessionIdsByDate[completedDateOnly] || [];
    matchingSessionIds.forEach(sessionId => {
      if (!completedBySession[sessionId]) return;
      completedBySession[sessionId][instrumentNumber] = (completedBySession[sessionId][instrumentNumber] || 0) + 1;
    });
  }

  Object.keys(summaries).forEach(sessionId => {
    const completedMap = completedBySession[sessionId] || {};
    const items = Object.keys(completedMap)
      .map(instrumentNumber => {
        const total = totalsByInstrument[instrumentNumber] || participants.length;
        const completed = completedMap[instrumentNumber] || 0;
        return {
          number: Number(instrumentNumber),
          name: namesByInstrument[instrumentNumber] || ('Item ' + instrumentNumber),
          completed: completed,
          total: total,
          percentage: total > 0 ? Math.round((completed / total) * 100) : 0
        };
      })
      .filter(item => item.completed > 0)
      .sort((a, b) => a.number - b.number);

    summaries[sessionId].items = items;
    summaries[sessionId].completed = items.reduce((sum, item) => sum + item.completed, 0);
    summaries[sessionId].total = items.reduce((sum, item) => sum + item.total, 0);
    summaries[sessionId].completionRate = summaries[sessionId].total > 0
      ? Math.round((summaries[sessionId].completed / summaries[sessionId].total) * 100)
      : 0;
  });

  return summaries;
}

function resolveCohortLifecycle(cohort, sessions, today) {
  if (!sessions || sessions.length === 0) {
    return { key: 'unscheduled', label: 'Unscheduled', tone: 'neutral', detail: 'No sessions scheduled' };
  }

  const datedSessions = sessions
    .map(session => getDateOnly(parseSessionDate(session.sessionDate)))
    .filter(Boolean);

  const completedByStatus = sessions.filter(session => String(session.status || '').toLowerCase() === 'completed').length;
  if (completedByStatus === sessions.length) {
    return { key: 'completed', label: 'Completed', tone: 'success', detail: 'All sessions marked complete' };
  }

  if (datedSessions.length === sessions.length) {
    const firstDate = datedSessions[0];
    const lastDate = datedSessions[datedSessions.length - 1];
    if (firstDate > today) {
      return { key: 'upcoming', label: 'Upcoming', tone: 'info', detail: 'Starts ' + formatDashboardDate(firstDate) };
    }
    if (lastDate < today) {
      return { key: 'completed', label: 'Completed', tone: 'success', detail: 'Ended ' + formatDashboardDate(lastDate) };
    }
    return { key: 'inProgress', label: 'In Progress', tone: 'warning', detail: 'Session window is active' };
  }

  if (String(cohort.status || '').toLowerCase() === 'active') {
    return { key: 'inProgress', label: 'In Progress', tone: 'warning', detail: 'Active cohort' };
  }

  return { key: 'upcoming', label: 'Upcoming', tone: 'info', detail: 'Schedule partially pending' };
}

function resolveSessionTiming(session, sessionDateOnly, today) {
  if (String(session.status || '').toLowerCase() === 'completed') return 'completed';
  if (!sessionDateOnly) return 'unscheduled';
  if (sessionDateOnly < today) return 'completed';
  if (sessionDateOnly === today) return 'today';
  return 'upcoming';
}

function sortSessionsByNumberAndDate(a, b) {
  const aNumber = Number(a.sessionNumber) || 0;
  const bNumber = Number(b.sessionNumber) || 0;
  if (aNumber !== bNumber) return aNumber - bNumber;
  return compareDateStrings(a.sessionDate, b.sessionDate);
}

function compareDateStrings(a, b) {
  const aDate = getDateOnly(parseSessionDate(a));
  const bDate = getDateOnly(parseSessionDate(b));
  if (!aDate && !bDate) return 0;
  if (!aDate) return 1;
  if (!bDate) return -1;
  return aDate.localeCompare(bDate);
}

function getDateOnly(date) {
  if (!date || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDashboardDate(dateString) {
  if (!dateString) return 'TBD';
  const parsed = parseSessionDate(dateString);
  if (!parsed) return dateString;
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'MMM d, yyyy');
}

function buildLandingContextLine(currentCohort, effectiveSite, effectiveCohort) {
  if (currentCohort) {
    const label = formatCohortDisplayLabel(currentCohort);
    return `Showing data for: ${label} • ${currentCohort.site}`;
  }

  const cohortLabel = effectiveCohort === 'All' ? 'All Study Cohorts' : 'Selected Study Cohort';
  const siteLabel = effectiveSite === 'All' ? 'All Sites' : effectiveSite;
  return `Showing data for: ${cohortLabel} • ${siteLabel}`;
}

function formatCohortDisplayLabel(cohort) {
  const periodYear = [cohort.period, cohort.year].filter(Boolean).join(' ');
  return `${cohort.schoolName || 'Study Cohort'} – ${periodYear} Cohort`;
}

function buildAttendanceSummary(participants, sessionsByCohort, attendanceMap, effectiveCohort) {
  const attendanceRates = [];
  let atRisk70 = 0;
  let atRisk60 = 0;
  let missingAttendance = 0;
  let participantsWithSessions = 0;

  const distribution = { high: 0, mid: 0, low: 0 };
  const totalSessions = Object.values(sessionsByCohort).reduce((sum, sessions) => sum + sessions.length, 0);
  const sessionsCompleted = countCompletedSessions(sessionsByCohort);

  participants.forEach(participant => {
    const sessions = sessionsByCohort[participant.rolloutId] || [];
    if (sessions.length === 0) return;

    participantsWithSessions++;
    let presentCount = 0;

    sessions.forEach(session => {
      const status = attendanceMap[session.sessionId]
        ? attendanceMap[session.sessionId][participant.participantId]
        : null;
      if (status === 'present') {
        presentCount++;
      }
      if (!status) {
        missingAttendance++;
      }
    });

    const attendanceRate = Math.round((presentCount / sessions.length) * 100);
    attendanceRates.push(attendanceRate);

    if (attendanceRate < 70) atRisk70++;
    if (attendanceRate < 60) atRisk60++;

    if (attendanceRate >= 90) distribution.high++;
    else if (attendanceRate >= 70) distribution.mid++;
    else distribution.low++;
  });

  const averageAttendance = attendanceRates.length > 0
    ? Math.round(attendanceRates.reduce((sum, rate) => sum + rate, 0) / attendanceRates.length)
    : 0;

  const participationStatus = resolveParticipationStatus(averageAttendance, participantsWithSessions, atRisk70);
  const trend = buildAttendanceTrend(participants, sessionsByCohort, attendanceMap, effectiveCohort);
  const totalDistribution = distribution.high + distribution.mid + distribution.low;

  return {
    averageAttendance: averageAttendance,
    participationStatus: participationStatus,
    atRiskParticipants: atRisk70,
    atRiskParticipantsCritical: atRisk60,
    distribution: {
      high: distribution.high,
      mid: distribution.mid,
      low: distribution.low,
      total: totalDistribution
    },
    totalSessions: totalSessions,
    sessionsCompleted: sessionsCompleted,
    trend: trend,
    trendAvailable: effectiveCohort !== 'All' && trend.length > 0,
    missingAttendanceRecords: missingAttendance,
    participantsWithSessions: participantsWithSessions
  };
}

function resolveParticipationStatus(averageAttendance, participantsWithSessions, atRisk70) {
  if (participantsWithSessions === 0) {
    return { label: 'No Data', tone: 'neutral', subtitle: 'Attendance data pending' };
  }
  const atRiskShare = participantsWithSessions > 0 ? (atRisk70 / participantsWithSessions) * 100 : 0;
  if (averageAttendance >= 85 && atRiskShare < 10) {
    return { label: 'Good', tone: 'success', subtitle: `Avg attendance: ${averageAttendance}%` };
  }
  if (averageAttendance >= 70 && averageAttendance < 85 || (atRiskShare >= 10 && atRiskShare <= 25)) {
    return { label: 'Watch', tone: 'warning', subtitle: `Avg attendance: ${averageAttendance}%` };
  }
  return { label: 'At Risk', tone: 'danger', subtitle: `Avg attendance: ${averageAttendance}%` };
}

function buildAttendanceTrend(participants, sessionsByCohort, attendanceMap, effectiveCohort) {
  if (effectiveCohort === 'All') return [];
  const sessions = sessionsByCohort[effectiveCohort] || [];
  const participantsForCohort = participants.filter(participant => participant.rolloutId === effectiveCohort);

  sessions.sort((a, b) => (a.sessionNumber || 0) - (b.sessionNumber || 0));

  return sessions.map(session => {
    let present = 0;
    let absent = 0;
    let excused = 0;
    let notMarked = 0;

    participantsForCohort.forEach(participant => {
      const status = attendanceMap[session.sessionId]
        ? attendanceMap[session.sessionId][participant.participantId]
        : null;
      if (status === 'present') present++;
      else if (status === 'absent') absent++;
      else if (status === 'excused') excused++;
      else notMarked++;
    });

    const total = present + absent + excused + notMarked;
    const attendanceRate = total > 0 ? Math.round((present / total) * 100) : 0;
    const label = session.sessionName || `Session ${session.sessionNumber || ''}`.trim();

    return {
      label: label,
      present: present,
      absent: absent,
      excused: excused,
      notMarked: notMarked,
      attendanceRate: attendanceRate
    };
  });
}

function countCompletedSessions(sessionsByCohort) {
  const today = new Date();
  let completed = 0;

  Object.values(sessionsByCohort).forEach(sessions => {
    sessions.forEach(session => {
      if (session.status === 'completed') {
        completed++;
        return;
      }
      const sessionDate = parseSessionDate(session.sessionDate);
      if (sessionDate && sessionDate <= today) {
        completed++;
      }
    });
  });

  return completed;
}

function parseDateOnlyString(value) {
  const match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);
}

function parseSessionDate(dateValue) {
  if (!dateValue) return null;
  const normalized = normalizeSessionDateValue(dateValue);
  if (!normalized) return null;
  const dateOnly = parseDateOnlyString(normalized);
  if (dateOnly) return dateOnly;
  const parsed = new Date(normalized);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function buildProtocolSummary(participants, checklistSheet) {
  const participantIds = new Set(participants.map(p => p.participantId));
  const participantRolloutMap = {};
  participants.forEach(p => { participantRolloutMap[String(p.participantId)] = String(p.rolloutId || ''); });
  const instrumentCounts = {};
  let totalCompleted = 0;
  const rolloutAllowedNumbersMap = {};
  const enabledItemsByNumber = {};
  participants.forEach(p => {
    const rolloutId = String(p.rolloutId || '');
    if (!rolloutAllowedNumbersMap[rolloutId]) {
      const allowed = {};
      getRolloutProtocolItemsInternal(rolloutId).forEach(item => {
        allowed[String(item.number)] = true;
        enabledItemsByNumber[String(item.number)] = item;
      });
      rolloutAllowedNumbersMap[rolloutId] = allowed;
    }
  });

  if (checklistSheet && participantIds.size > 0) {
    const checklistSnapshot = getSheetSnapshot('Checklist', { ensureFn: ensureChecklistColumns });
    const cData = checklistSnapshot.data;
    const cHeaders = checklistSnapshot.headers;
    for (let i = 1; i < cData.length; i++) {
      const participantId = String(cData[i][cHeaders.indexOf('participantId')]);
      if (!participantIds.has(participantId)) continue;
      const instNumber = String(cData[i][cHeaders.indexOf('instrumentNumber')]);
      const status = cData[i][cHeaders.indexOf('status')];
      const rolloutId = participantRolloutMap[participantId] || '';
      const allowedMap = rolloutAllowedNumbersMap[rolloutId] || {};
      if (!allowedMap[instNumber]) continue;
      if (!instrumentCounts[instNumber]) {
        instrumentCounts[instNumber] = {
          name: cData[i][cHeaders.indexOf('instrumentName')] || (enabledItemsByNumber[instNumber] || {}).name || ('Item ' + instNumber),
          total: 0,
          completed: 0
        };
      }

      instrumentCounts[instNumber].total++;
      if (status === 'completed') {
        instrumentCounts[instNumber].completed++;
        totalCompleted++;
      }
    }
  }

  const totalRequired = Object.values(instrumentCounts).reduce((sum, item) => sum + item.total, 0);
  const requiredItems = participants.length > 0 ? Math.round(totalRequired / participants.length) : 0;
  const completionRate = totalRequired > 0 ? Math.round((totalCompleted / totalRequired) * 100) : 0;

  const instrumentStats = Object.keys(instrumentCounts).map(instNumber => {
    const counts = instrumentCounts[instNumber];
    const percentage = counts && counts.total > 0 ? Math.round((counts.completed / counts.total) * 100) : 0;
    return {
      number: Number(instNumber),
      name: counts.name,
      total: counts.total,
      completed: counts.completed,
      percentage: percentage
    };
  }).sort((a, b) => a.number - b.number);

  const lowestItems = instrumentStats
    .filter(item => item.total > 0)
    .sort((a, b) => a.percentage - b.percentage)
    .slice(0, 3);

  const needsReviewCount = instrumentStats.filter(item => item.total > 0 && item.percentage < 50).length;

  return {
    requiredItems: requiredItems,
    totalRequired: totalRequired,
    completedItems: totalCompleted,
    completionRate: completionRate,
    instrumentStats: instrumentStats,
    lowestItems: lowestItems,
    needsReviewCount: needsReviewCount
  };
}

function buildAttentionItems(attendanceSummary, protocolSummary, participants, effectiveCohort) {
  const items = [];

  if (participants.length === 0) {
    return items;
  }

  if (attendanceSummary.atRiskParticipantsCritical > 0) {
    items.push({
      severity: 'critical',
      message: `${attendanceSummary.atRiskParticipantsCritical} participants below 60% attendance (At Risk).`,
      action: 'Review attendance'
    });
  }

  if (attendanceSummary.atRiskParticipants > 0) {
    items.push({
      severity: 'warning',
      message: `${attendanceSummary.atRiskParticipants} participants below 70% attendance.`,
      action: 'View attendance report'
    });
  }

  if (effectiveCohort !== 'All') {
    const lowSessions = attendanceSummary.trend.filter(session => session.attendanceRate > 0 && session.attendanceRate < 60);
    if (lowSessions.length > 0) {
      const sessionNames = lowSessions.slice(0, 2).map(session => `${session.label} (${session.attendanceRate}%)`);
      items.push({
        severity: 'warning',
        message: `Low attendance detected: ${sessionNames.join(', ')}.`,
        action: 'Review session attendance'
      });
    }
  }

  if (attendanceSummary.missingAttendanceRecords > 0) {
    items.push({
      severity: 'info',
      message: `${attendanceSummary.missingAttendanceRecords} attendance records are missing for selected sessions.`,
      action: 'Complete attendance'
    });
  }

  if (attendanceSummary.totalSessions === 0 && participants.length > 0) {
    items.push({
      severity: 'info',
      message: 'No sessions scheduled for the selected cohort. Add sessions to enable attendance tracking.',
      action: 'Schedule sessions'
    });
  }

  if (protocolSummary.needsReviewCount > 0) {
    items.push({
      severity: 'warning',
      message: `${protocolSummary.needsReviewCount} protocol items are below 50% completion.`,
      action: 'Review protocol checklist'
    });
  }

  return items.slice(0, 7);
}

function buildComparisonSummary(
  effectiveSite,
  effectiveCohort,
  participantsBySite,
  participantsByCohort,
  sessionsByCohort,
  attendanceMap,
  checklistSheet,
  cohorts
) {
  if (effectiveSite === 'All' && effectiveCohort === 'All') {
    const siteEntries = ['UGA', 'Missouri'].map(site => {
      const participants = participantsBySite[site] || [];
      const summary = buildAttendanceSummary(participants, sessionsByCohort, attendanceMap, 'All');
      const protocol = buildProtocolSummary(participants, checklistSheet);
      return {
        label: site,
        participants: participants.length,
        averageAttendance: summary.averageAttendance,
        protocolCompletion: protocol.completionRate
      };
    }).filter(entry => entry.participants > 0);

    return {
      title: 'Site Comparison',
      entries: siteEntries,
      emptyMessage: 'No participant data available for site comparison.'
    };
  }

  if (effectiveSite !== 'All' && effectiveCohort !== 'All') {
    return {
      title: 'Cohort Comparison',
      entries: [],
      emptyMessage: 'Comparison is unavailable for a single cohort selection.'
    };
  }

  if (effectiveSite !== 'All') {
    const cohortEntries = cohorts.map(cohort => {
      const participants = participantsByCohort[cohort.rolloutId] || [];
      const summary = buildAttendanceSummary(participants, sessionsByCohort, attendanceMap, 'All');
      const protocol = buildProtocolSummary(participants, checklistSheet);
      return {
        label: formatCohortDisplayLabel(cohort),
        participants: participants.length,
        averageAttendance: summary.averageAttendance,
        protocolCompletion: protocol.completionRate
      };
    }).filter(entry => entry.participants > 0);

    return {
      title: 'Cohort Comparison',
      entries: cohortEntries,
      emptyMessage: 'No cohorts with participant data found for this site.'
    };
  }

  return {
    title: 'Comparison',
    entries: [],
    emptyMessage: 'Comparison is unavailable for the selected cohort filter.'
  };
}

function buildAboutSummary(effectiveSite, effectiveCohort, totalEnrolled, cohorts, participantsBySite) {
  const siteCount = effectiveSite === 'All'
    ? Object.keys(participantsBySite).filter(site => participantsBySite[site] && participantsBySite[site].length > 0).length
    : 1;
  const cohortCount = effectiveCohort === 'All' ? cohorts.length : 1;

  return {
    siteCount: siteCount,
    cohortCount: cohortCount,
    protocolItems: CONFIG.INSTRUMENTS.length,
    platform: 'Nintendo Switch',
    totalParticipants: totalEnrolled
  };
}

/**
 * Get public cohorts list (no authentication required)
 */
function getPublicRollouts(siteFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!rolloutsSheet) return { success: true, cohorts: [] };

  const data = rolloutsSheet.getDataRange().getValues();
  const headers = data[0];
  const cohorts = [];

  for (let i = 1; i < data.length; i++) {
    const site = data[i][headers.indexOf('site')];
    const status = data[i][headers.indexOf('status')];

    // Apply site filter
    if (siteFilter && siteFilter !== 'All' && site !== siteFilter) {
      continue;
    }

    // Only return active cohorts for public view
    if (status === 'active') {
      cohorts.push({
        rolloutId: data[i][headers.indexOf('rolloutId')],
        site: site,
        schoolName: data[i][headers.indexOf('schoolName')],
        period: data[i][headers.indexOf('period')],
        year: data[i][headers.indexOf('year')],
        status: status
      });
    }
  }

  return { success: true, cohorts: cohorts };
}

/**
 * Get site comparison data
 */
function getSiteComparison(token) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  const ugaStats = getDashboardStats(token, 'UGA', 'All');
  const missouriStats = getDashboardStats(token, 'Missouri', 'All');

  return {
    success: true,
    comparison: {
      UGA: ugaStats,
      Missouri: missouriStats
    }
  };
}

// ============================================
// ACTIVITY LOG FUNCTIONS
// ============================================

/**
 * Log an activity
 */
function logActivity(userId, userName, action, targetType, targetId, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ActivityLog');

  if (!logSheet) return;

  const logId = generateUUID();
  const timestamp = new Date().toISOString();

  logSheet.appendRow([
    logId,
    timestamp,
    userId,
    userName,
    action,
    targetType,
    targetId,
    details
  ]);
}

/**
 * Get recent activity logs
 */
function getRecentActivity(token, limit) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  limit = limit || 50;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ActivityLog');

  if (!logSheet) return { success: true, activities: [] };

  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  const activities = [];

  // Get last 'limit' entries (reverse order)
  for (let i = Math.min(data.length - 1, limit); i >= 1; i--) {
    activities.push({
      logId: data[i][headers.indexOf('logId')],
      timestamp: data[i][headers.indexOf('timestamp')],
      userId: data[i][headers.indexOf('userId')],
      userName: data[i][headers.indexOf('userName')],
      action: data[i][headers.indexOf('action')],
      targetType: data[i][headers.indexOf('targetType')],
      targetId: data[i][headers.indexOf('targetId')],
      details: data[i][headers.indexOf('details')]
    });
  }

  // Sort by timestamp descending
  activities.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  return { success: true, activities: activities.slice(0, limit) };
}

// ============================================
// EXPORT FUNCTIONS
// ============================================

/**
 * Export participants data to CSV format
 */
function exportParticipantsCSV(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  // Enforce facilitator site scope
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    filters = filters || {};
    filters.site = currentUser.site;
  }

  const result = getAllParticipants(token, filters);
  if (!result.success) return result;

  const participants = result.participants;

  // Create CSV content
  const includeEnrollmentFields = currentUser.role !== 'viewer';
  const enrollmentFields = getEnrollmentFieldsInternal().filter(field => field.key !== 'fullName');
  const headers = ['Participant ID', 'Full Name'];

  if (includeEnrollmentFields) {
    enrollmentFields.forEach(field => headers.push(field.label));
  }

  headers.push('Site', 'School', 'Period', 'Year', 'Enrollment Date', 'Status', 'Completion %', 'Notes');

  let csv = headers.join(',') + '\n';

  participants.forEach(p => {
    const row = [
      p.participantId,
      '"' + (p.fullName || '').replace(/"/g, '""') + '"'
    ];

    if (includeEnrollmentFields) {
      enrollmentFields.forEach(field => {
        const value = p.enrollmentFields && p.enrollmentFields[field.key] !== undefined ? p.enrollmentFields[field.key] : '';
        row.push('"' + String(value || '').replace(/"/g, '""') + '"');
      });
    }

    row.push(
      p.site,
      '"' + (p.schoolName || '').replace(/"/g, '""') + '"',
      p.period,
      p.year,
      p.enrollmentDate,
      p.status,
      p.completionPercentage,
      '"' + (p.notes || '').replace(/"/g, '""') + '"'
    );
    csv += row.join(',') + '\n';
  });

  return { success: true, csv: csv, filename: 'participants_export_' + new Date().toISOString().split('T')[0] + '.csv' };
}

/**
 * Export full checklist data
 */
function exportChecklistCSV(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  // Enforce facilitator site scope
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    filters = filters || {};
    filters.site = currentUser.site;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');

  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  // Build participant map scoped to filters
  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pData = participantSnapshot.data;
  const pHeaders = participantSnapshot.headers;
  const participantMap = {};

  for (let i = 1; i < pData.length; i++) {
    const pid = pData[i][pHeaders.indexOf('participantId')];
    const site = pData[i][pHeaders.indexOf('site')];
    const rolloutId = pData[i][pHeaders.indexOf('rolloutId')];

    if (filters && filters.site && filters.site !== 'All' && site !== filters.site) {
      continue;
    }
    if (filters && filters.rolloutId && filters.rolloutId !== 'All' && rolloutId !== filters.rolloutId) {
      continue;
    }

    participantMap[pid] = {
      fullName: pData[i][pHeaders.indexOf('fullName')],
      schoolName: pData[i][pHeaders.indexOf('schoolName')],
      completed: 0,
      total: 0
    };
  }

  // Aggregate checklist data by participant
  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];

  for (let i = 1; i < cData.length; i++) {
    const pid = cData[i][cHeaders.indexOf('participantId')];
    if (!participantMap[pid]) continue;

    participantMap[pid].total++;
    if (cData[i][cHeaders.indexOf('status')] === 'completed') {
      participantMap[pid].completed++;
    }
  }

  const headers = ['Participant Name', 'School', 'Completed Instruments', 'Incomplete Instruments', 'Completion %'];
  let csv = headers.join(',') + '\n';

  Object.keys(participantMap).forEach(pid => {
    const p = participantMap[pid];
    const incomplete = Math.max(p.total - p.completed, 0);
    const percentage = p.total > 0 ? Math.round((p.completed / p.total) * 100) : 0;

    const row = [
      '"' + (p.fullName || '').replace(/"/g, '""') + '"',
      '"' + (p.schoolName || '').replace(/"/g, '""') + '"',
      p.completed,
      incomplete,
      percentage
    ];
    csv += row.join(',') + '\n';
  });

  return { success: true, csv: csv, filename: 'checklist_export_' + new Date().toISOString().split('T')[0] + '.csv' };
}

/**
 * Export attendance data to CSV format
 */
function exportAttendanceCSV(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  // Enforce facilitator site scope
  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    filters = filters || {};
    filters.site = currentUser.site;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const sessionsSheet = ss.getSheetByName('StudySessions');
  const attendanceSheet = ss.getSheetByName('SessionAttendance');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!participantsSheet || !sessionsSheet || !attendanceSheet || !rolloutsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const rolloutsData = rolloutsSheet.getDataRange().getValues();
  const rolloutsHeaders = rolloutsData[0];
  const rolloutEntries = [];

  for (let i = 1; i < rolloutsData.length; i++) {
    const rolloutId = rolloutsData[i][rolloutsHeaders.indexOf('rolloutId')];
    const rolloutSite = rolloutsData[i][rolloutsHeaders.indexOf('site')];

    if (filters && filters.site && filters.site !== 'All' && rolloutSite !== filters.site) continue;
    if (filters && filters.rolloutId && filters.rolloutId !== 'All' && rolloutId !== filters.rolloutId) continue;

    rolloutEntries.push({
      rolloutId: rolloutId,
      site: rolloutSite,
      schoolName: rolloutsData[i][rolloutsHeaders.indexOf('schoolName')],
      period: rolloutsData[i][rolloutsHeaders.indexOf('period')],
      year: rolloutsData[i][rolloutsHeaders.indexOf('year')]
    });
  }

  const participantsData = participantsSheet.getDataRange().getValues();
  const participantsHeaders = participantsData[0];
  const participantsByRollout = {};

  for (let i = 1; i < participantsData.length; i++) {
    const rolloutId = participantsData[i][participantsHeaders.indexOf('rolloutId')];
    if (!rolloutEntries.find(entry => entry.rolloutId === rolloutId)) continue;

    const participant = {
      participantId: participantsData[i][participantsHeaders.indexOf('participantId')],
      fullName: participantsData[i][participantsHeaders.indexOf('fullName')],
      site: participantsData[i][participantsHeaders.indexOf('site')],
      rolloutId: rolloutId
    };

    if (!participantsByRollout[rolloutId]) {
      participantsByRollout[rolloutId] = [];
    }
    participantsByRollout[rolloutId].push(participant);
  }

  const sessionsData = sessionsSheet.getDataRange().getValues();
  const sessionsHeaders = sessionsData[0];
  const sessionsByRollout = {};

  for (let i = 1; i < sessionsData.length; i++) {
    const rolloutId = sessionsData[i][sessionsHeaders.indexOf('rolloutId')];
    if (!rolloutEntries.find(entry => entry.rolloutId === rolloutId)) continue;

    const session = {
      sessionId: sessionsData[i][sessionsHeaders.indexOf('sessionId')],
      sessionNumber: sessionsData[i][sessionsHeaders.indexOf('sessionNumber')],
      sessionDate: normalizeSessionDateValue(sessionsData[i][sessionsHeaders.indexOf('sessionDate')]),
      sessionName: sessionsData[i][sessionsHeaders.indexOf('sessionName')]
    };

    if (!sessionsByRollout[rolloutId]) {
      sessionsByRollout[rolloutId] = [];
    }
    sessionsByRollout[rolloutId].push(session);
  }

  const attendanceData = attendanceSheet.getDataRange().getValues();
  const attendanceHeaders = attendanceData[0];
  const attendanceMap = {};

  for (let i = 1; i < attendanceData.length; i++) {
    const sessionId = attendanceData[i][attendanceHeaders.indexOf('sessionId')];
    const participantId = attendanceData[i][attendanceHeaders.indexOf('participantId')];
    attendanceMap[sessionId + '_' + participantId] = attendanceData[i][attendanceHeaders.indexOf('status')] || 'not_marked';
  }

  const rolloutLookup = {};
  rolloutEntries.forEach(entry => {
    rolloutLookup[entry.rolloutId] = entry;
  });

  const maxSessionCount = rolloutEntries.reduce((max, cohort) => {
    const sessions = sessionsByRollout[cohort.rolloutId] || [];
    const maxSessionNumber = sessions.reduce((count, session) => {
      return Math.max(count, session.sessionNumber || 0);
    }, 0);
    return Math.max(max, maxSessionNumber);
  }, 0);

  const headerRow = [
    'Participant Name',
    'Participant ID',
    'Site',
    'Cohort'
  ];

  for (let i = 1; i <= maxSessionCount; i++) {
    headerRow.push(`Session ${i}`);
  }

  headerRow.push(
    'Total Sessions',
    'Sessions Attended',
    'Sessions Missed',
    'Attendance %',
    'Missed Sessions'
  );

  const rows = [];
  const attendanceRates = [];
  const bandCounts = {
    high: 0,
    mid: 0,
    low: 0
  };

  const sessionSummaryRows = [];
  const rowsByRollout = {};
  rolloutEntries.forEach(cohort => {
    const sessions = (sessionsByRollout[cohort.rolloutId] || []).slice().sort((a, b) => a.sessionNumber - b.sessionNumber);
    const participants = (participantsByRollout[cohort.rolloutId] || []).slice().sort((a, b) => {
      return (a.fullName || '').localeCompare(b.fullName || '');
    });
    const sessionMap = {};
    sessions.forEach(session => {
      sessionMap[session.sessionNumber] = session.sessionId;
    });
    const sessionCount = sessions.reduce((count, session) => Math.max(count, session.sessionNumber || 0), 0);
    rowsByRollout[cohort.rolloutId] = [];

    // Assumption: if a participant has no attendance record for a session, count as Absent.
    sessions.forEach(session => {
      let presentCount = 0;
      participants.forEach(participant => {
        const status = attendanceMap[session.sessionId + '_' + participant.participantId] || 'not_marked';
        if (status === 'present') presentCount++;
      });
      const total = participants.length;
      const absentCount = total - presentCount;
      const rate = total > 0 ? Math.round((presentCount / total) * 100) : 0;
      sessionSummaryRows.push([
        `${cohort.schoolName || 'Cohort'} (${cohort.period || ''} ${cohort.year || ''})`.replace(/\s+/g, ' ').trim(),
        'Session ' + session.sessionNumber,
        presentCount,
        absentCount,
        rate + '%'
      ]);
    });

    participants.forEach(participant => {
      let presentCount = 0;
      let missedCount = 0;
      const missedLabels = [];
      const sessionValues = [];

      for (let i = 1; i <= sessionCount; i++) {
        if (!sessionMap[i]) {
          sessionValues.push('');
          continue;
        }

        const status = attendanceMap[sessionMap[i] + '_' + participant.participantId] || 'not_marked';
        // Assumption: missing attendance is treated as absent, excused counts as non-present.
        if (status === 'present') {
          presentCount++;
          sessionValues.push('P');
        } else if (status === 'excused') {
          missedCount++;
          sessionValues.push('E');
          missedLabels.push('S' + i);
        } else {
          missedCount++;
          sessionValues.push('A');
          missedLabels.push('S' + i);
        }
      }

      const totalSessions = sessions.length;
      const attendanceRate = totalSessions > 0 ? Math.round((presentCount / totalSessions) * 100) : 0;
      attendanceRates.push(attendanceRate);

      if (attendanceRate >= 90) bandCounts.high++;
      else if (attendanceRate >= 70) bandCounts.mid++;
      else bandCounts.low++;

      const row = [
        participant.fullName || '',
        participant.participantId || '',
        participant.site || cohort.site || '',
        `${cohort.schoolName || 'Cohort'} (${cohort.period || ''} ${cohort.year || ''})`.replace(/\s+/g, ' ').trim(),
        ...sessionValues,
        totalSessions,
        presentCount,
        missedCount,
        attendanceRate + '%',
        missedLabels.length > 0 ? missedLabels.join(', ') : ''
      ];

      rows.push(row);
      rowsByRollout[cohort.rolloutId].push(row);
    });
  });

  rows.sort((a, b) => {
    const rolloutCompare = a[3].localeCompare(b[3]);
    if (rolloutCompare !== 0) return rolloutCompare;
    return a[0].localeCompare(b[0]);
  });

  const exportDate = new Date();
  const siteLabel = (filters && filters.site) ? filters.site : 'All';
  let rolloutLabel = 'All Cohorts';
  if (filters && filters.rolloutId && filters.rolloutId !== 'All') {
    const rolloutMatch = rolloutLookup[filters.rolloutId];
    if (rolloutMatch) {
      rolloutLabel = `${rolloutMatch.schoolName || 'Cohort'} (${rolloutMatch.period || ''} ${rolloutMatch.year || ''})`
        .replace(/\s+/g, ' ')
        .trim();
    } else {
      rolloutLabel = filters.rolloutId;
    }
  }

  const spreadsheet = SpreadsheetApp.create('Attendance Export ' + exportDate.toISOString());
  let sheet = spreadsheet.getSheets()[0];
  sheet.setName('Attendance Export');

  const buildSheetName = cohort => {
    const label = `${cohort.schoolName || 'Cohort'} ${cohort.period || ''} ${cohort.year || ''}`.replace(/\s+/g, ' ').trim();
    return label.substring(0, 90) || 'Cohort';
  };

  const buildHeaderRow = sessionCount => {
    const base = ['Participant Name', 'Participant ID', 'Site', 'Cohort'];
    for (let i = 1; i <= sessionCount; i++) {
      base.push(`Session ${i}`);
    }
    base.push('Total Sessions', 'Sessions Attended', 'Sessions Missed', 'Attendance %', 'Missed Sessions');
    return base;
  };

  const renderAttendanceSheet = (targetSheet, sheetRows, sessionSummary, sessionCount, sheetRolloutLabel) => {
    const sheetHeaderRow = buildHeaderRow(sessionCount);
    const totalColumns = sheetHeaderRow.length;
    let rowCursor = 1;

    targetSheet.getRange(rowCursor, 1).setValue('Attendance Export');
    rowCursor++;
    targetSheet.getRange(rowCursor, 1).setValue('Site: ' + (siteLabel === 'All' ? 'All Sites' : siteLabel));
    targetSheet.getRange(rowCursor, 2).setValue('Cohort: ' + sheetRolloutLabel);
    targetSheet.getRange(rowCursor, 3).setValue('Exported: ' + exportDate.toLocaleString());
    rowCursor += 2;

    const titleRange = targetSheet.getRange(1, 1, 1, totalColumns);
    titleRange.setFontSize(16).setFontWeight('bold').setFontFamily('Arial');
    titleRange.setBackground('#1f2937').setFontColor('#ffffff');

    const metaRange = targetSheet.getRange(2, 1, 1, totalColumns);
    metaRange.setFontSize(10).setFontFamily('Arial').setFontColor('#111827');

    const totalParticipants = sheetRows.length;
    const sheetAttendanceRates = sheetRows.map(row => {
      const rateValue = String(row[sheetHeaderRow.length - 2]).replace('%', '');
      return Number(rateValue) || 0;
    });
    const averageAttendance = sheetAttendanceRates.length > 0
      ? Math.round(sheetAttendanceRates.reduce((sum, rate) => sum + rate, 0) / sheetAttendanceRates.length)
      : 0;

    const bandCounts = {
      high: sheetAttendanceRates.filter(rate => rate >= 90).length,
      mid: sheetAttendanceRates.filter(rate => rate >= 70 && rate < 90).length,
      low: sheetAttendanceRates.filter(rate => rate < 70).length
    };

    targetSheet.getRange(rowCursor, 1).setValue('Summary');
    targetSheet.getRange(rowCursor, 1).setFontWeight('bold').setFontFamily('Arial');
    rowCursor++;
    const summaryRange = targetSheet.getRange(rowCursor, 1, 5, 2);
    summaryRange.setValues([
      ['Total Participants', totalParticipants],
      ['Average Attendance %', averageAttendance + '%'],
      ['90–100%', bandCounts.high],
      ['70–89%', bandCounts.mid],
      ['<70%', bandCounts.low]
    ]);
    summaryRange.setFontFamily('Arial').setFontSize(10);
    targetSheet.getRange(rowCursor, 1, 5, 1).setFontWeight('bold').setBackground('#f3f4f6');
    rowCursor += 6;

    targetSheet.getRange(rowCursor, 1).setValue('Session Summary');
    targetSheet.getRange(rowCursor, 1).setFontWeight('bold').setFontFamily('Arial');
    rowCursor++;
    const sessionHeaderRange = targetSheet.getRange(rowCursor, 1, 1, 5);
    sessionHeaderRange.setValues([['Cohort', 'Session', 'Present', 'Absent', 'Attendance %']]);
    sessionHeaderRange.setFontWeight('bold').setBackground('#f3f4f6').setFontFamily('Arial');
    rowCursor++;
    if (sessionSummary.length > 0) {
      const sessionSummaryRange = targetSheet.getRange(rowCursor, 1, sessionSummary.length, 5);
      sessionSummaryRange.setValues(sessionSummary);
      sessionSummaryRange.setFontFamily('Arial').setFontSize(10);
      rowCursor += sessionSummary.length + 1;
    } else {
      rowCursor++;
    }

    const tableHeaderRow = rowCursor;
    const headerRange = targetSheet.getRange(tableHeaderRow, 1, 1, sheetHeaderRow.length);
    headerRange.setValues([sheetHeaderRow]);
    headerRange.setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    headerRange.setFontFamily('Arial').setFontSize(10);
    rowCursor++;

    if (sheetRows.length > 0) {
      const dataRange = targetSheet.getRange(rowCursor, 1, sheetRows.length, sheetHeaderRow.length);
      dataRange.setValues(sheetRows.map(row => row.slice(0, sheetHeaderRow.length)));
      dataRange.setFontFamily('Arial').setFontSize(10);
    }

    targetSheet.setFrozenRows(0);
    targetSheet.setFrozenColumns(0);

    if (sheetRows.length > 0 && sessionCount > 0) {
      const sessionStartCol = 5;
      const sessionRange = targetSheet.getRange(rowCursor, sessionStartCol, sheetRows.length, sessionCount);
      const rules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('P')
          .setBackground('#dcfce7')
          .setRanges([sessionRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('A')
          .setBackground('#fee2e2')
          .setRanges([sessionRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('E')
          .setBackground('#fef3c7')
          .setRanges([sessionRange])
          .build()
      ];
      targetSheet.setConditionalFormatRules(rules);
    }

    targetSheet.autoResizeColumns(1, Math.min(sheetHeaderRow.length, targetSheet.getMaxColumns()));
    targetSheet.setColumnWidth(1, 180);
    targetSheet.setColumnWidth(2, 150);
    targetSheet.setColumnWidth(3, 90);
    targetSheet.setColumnWidth(4, 220);
  };

  if (rolloutEntries.length > 1) {
    rolloutEntries.forEach((cohort, index) => {
      const sheetName = buildSheetName(cohort);
      const targetSheet = index === 0 ? sheet : spreadsheet.insertSheet(sheetName);
      if (index !== 0) {
        targetSheet.setName(sheetName);
      } else {
        sheet.setName(sheetName);
      }

      const sheetRows = rowsByRollout[cohort.rolloutId] || [];
      const sheetSessionSummary = sessionSummaryRows.filter(row => row[0] === `${cohort.schoolName || 'Cohort'} (${cohort.period || ''} ${cohort.year || ''})`.replace(/\s+/g, ' ').trim());
      const sessionCount = (sessionsByRollout[cohort.rolloutId] || []).reduce((count, session) => {
        return Math.max(count, session.sessionNumber || 0);
      }, 0);

      renderAttendanceSheet(targetSheet, sheetRows, sheetSessionSummary, sessionCount, `${cohort.schoolName || 'Cohort'} (${cohort.period || ''} ${cohort.year || ''})`.replace(/\s+/g, ' ').trim());
    });
  } else {
    const cohort = rolloutEntries[0];
    const sheetRows = cohort ? (rowsByRollout[cohort.rolloutId] || rows) : rows;
    const sessionCount = cohort ? (sessionsByRollout[cohort.rolloutId] || []).reduce((count, session) => {
      return Math.max(count, session.sessionNumber || 0);
    }, 0) : maxSessionCount;
    renderAttendanceSheet(sheet, sheetRows, sessionSummaryRows, sessionCount, rolloutLabel);
  }

  SpreadsheetApp.flush();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    return { success: false, message: 'Failed to export attendance report: ' + response.getContentText() };
  }

  const blob = response.getBlob().setName('attendance_export_' + exportDate.toISOString().split('T')[0] + '.xlsx');
  const base64 = Utilities.base64Encode(blob.getBytes());

  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);

  return {
    success: true,
    file: base64,
    mimeType: MimeType.MICROSOFT_EXCEL,
    filename: 'attendance_export_' + exportDate.toISOString().split('T')[0] + '.xlsx',
    downloadMessage: 'Attendance export downloaded'
  };
}

/**
 * Export summary report
 */

function exportLinksArtifacts(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) return { success: false, message: 'Invalid session' };

  const siteFilter = (filters && filters.site) ? filters.site : 'All';
  const rolloutFilter = (filters && filters.rolloutId) ? filters.rolloutId : 'All';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const sessionsSheet = ss.getSheetByName('StudySessions');
  if (!participantsSheet || !checklistSheet || !sessionsSheet) {
    return { success: false, message: 'Required sheets are missing' };
  }

  function normalizeInstrumentName(name) {
    return String(name || '').toLowerCase().replace(/\s*\([^)]*\)\s*/g, '').replace(/\s+/g, ' ').trim();
  }

  function makeSafeSheetName(name, used) {
    var base = String(name || 'Cohort').replace(/[\\/?*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!base) base = 'Cohort';
    base = base.substring(0, 31);
    var candidate = base;
    var n = 2;
    while (used[candidate]) {
      var suffix = ' (' + n + ')';
      candidate = base.substring(0, Math.max(1, 31 - suffix.length)) + suffix;
      n++;
    }
    used[candidate] = true;
    return candidate;
  }

  function toHyperlinkFormula(link) {
    var value = String(link || '').trim();
    if (!value) return '';
    if (!/^https?:\/\//i.test(value)) return '="' + value.replace(/"/g, '""') + '"';
    var safe = value.replace(/"/g, '""');
    return '=HYPERLINK("' + safe + '","Open")';
  }

  const exclusion = {
    'consent form': true,
    'assent form / pre-test': true,
    'post-test': true,
    'participant feedback survey': true
  };

  const participantsData = participantsSheet.getDataRange().getValues();
  const participantHeaders = participantsData[0] || [];
  const pIdx = {
    participantId: participantHeaders.indexOf('participantId'),
    fullName: participantHeaders.indexOf('fullName'),
    site: participantHeaders.indexOf('site'),
    rolloutId: participantHeaders.indexOf('rolloutId'),
    schoolName: participantHeaders.indexOf('schoolName'),
    period: participantHeaders.indexOf('period'),
    year: participantHeaders.indexOf('year')
  };

  const selectedParticipants = [];
  const selectedParticipantIds = {};
  const rolloutMeta = {};
  for (var i = 1; i < participantsData.length; i++) {
    const row = participantsData[i];
    const site = String(row[pIdx.site] || '');
    const rolloutId = String(row[pIdx.rolloutId] || '');
    if (siteFilter !== 'All' && site !== siteFilter) continue;
    if (rolloutFilter !== 'All' && rolloutId !== String(rolloutFilter)) continue;
    const pid = String(row[pIdx.participantId] || '');
    if (!pid) continue;
    const cohortLabel = String(row[pIdx.schoolName] || '') + (row[pIdx.period] ? ' (' + String(row[pIdx.period]) + ' ' + String(row[pIdx.year] || '') + ')' : '');
    selectedParticipantIds[pid] = true;
    selectedParticipants.push({ participantId: pid, participantName: String(row[pIdx.fullName] || ''), site: site, cohort: cohortLabel, rolloutId: rolloutId });
    if (!rolloutMeta[rolloutId]) rolloutMeta[rolloutId] = { site: site, cohort: cohortLabel };
  }

  if (!selectedParticipants.length) return { success: false, message: 'No participants found for selected filters' };

  const protocolByRollout = {};
  Object.keys(rolloutMeta).forEach(function(rid){ protocolByRollout[rid] = getProtocolItemsForFilter(rid) || []; });

  const checklistData = checklistSheet.getDataRange().getValues();
  const checklistHeaders = checklistData[0] || [];
  const cIdx = { participantId: checklistHeaders.indexOf('participantId'), instrumentName: checklistHeaders.indexOf('instrumentName'), dataLink: checklistHeaders.indexOf('dataLink') };
  const linkMap = {};
  const linkMapNormalized = {};
  for (var j = 1; j < checklistData.length; j++) {
    const crow = checklistData[j];
    const pid2 = String(crow[cIdx.participantId] || '');
    if (!selectedParticipantIds[pid2]) continue;
    const instrument = String(crow[cIdx.instrumentName] || '').trim();
    if (!instrument) continue;
    const link = String(crow[cIdx.dataLink] || '').trim();
    if (!link) continue;
    if (!linkMap[pid2]) linkMap[pid2] = {};
    if (!linkMapNormalized[pid2]) linkMapNormalized[pid2] = {};
    linkMap[pid2][instrument] = link;
    linkMapNormalized[pid2][normalizeInstrumentName(instrument)] = link;
  }

  const groups = {};
  selectedParticipants.forEach(function(p) {
    const key = p.rolloutId || ('cohort:' + p.cohort);
    if (!groups[key]) groups[key] = { rolloutId: p.rolloutId, cohort: p.cohort, participants: [] };
    groups[key].participants.push(p);
  });

  const spreadsheet = SpreadsheetApp.create('Links and Artifacts Export ' + new Date().toISOString());
  const usedNames = {};

  // Sheet 1: Sessions Recording
  const sessionData = sessionsSheet.getDataRange().getValues();
  const sHeaders = sessionData[0] || [];
  const sIdx = {
    rolloutId: sHeaders.indexOf('rolloutId'),
    sessionName: sHeaders.indexOf('sessionName'),
    sessionNumber: sHeaders.indexOf('sessionNumber'),
    goproLink: sHeaders.indexOf('goproLink'),
    tascamLink: sHeaders.indexOf('tascamLink'),
    meetingOwlLink: sHeaders.indexOf('meetingOwlLink'),
    fieldNotesLink: sHeaders.indexOf('fieldNotesLink'),
    recordingLinksJson: sHeaders.indexOf('recordingLinksJson')
  };
  const sessionRows = [];
  for (var sr = 1; sr < sessionData.length; sr++) {
    const row = sessionData[sr];
    const rid = String(row[sIdx.rolloutId] || '');
    if (!rolloutMeta[rid]) continue;
    let linksJson = {};
    if (sIdx.recordingLinksJson >= 0 && row[sIdx.recordingLinksJson]) {
      try { linksJson = JSON.parse(String(row[sIdx.recordingLinksJson])) || {}; } catch (e) {}
    }
    const gp = String((linksJson['GoPro'] || (sIdx.goproLink >= 0 ? row[sIdx.goproLink] : '') || '')).trim();
    const ta = String((linksJson['Tascam'] || (sIdx.tascamLink >= 0 ? row[sIdx.tascamLink] : '') || '')).trim();
    const mo = String((linksJson['Meeting Owl'] || (sIdx.meetingOwlLink >= 0 ? row[sIdx.meetingOwlLink] : '') || '')).trim();
    const sessionLabel = String((sIdx.sessionName >= 0 ? row[sIdx.sessionName] : '') || '').trim() || ('Session ' + String((sIdx.sessionNumber >= 0 ? row[sIdx.sessionNumber] : '') || '').trim());
    const fn = String((sIdx.fieldNotesLink >= 0 ? row[sIdx.fieldNotesLink] : '') || '').trim();
    sessionRows.push([
      rolloutMeta[rid].site,
      rolloutMeta[rid].cohort,
      sessionLabel,
      gp,
      ta,
      mo,
      fn
    ]);
  }

  const firstSheet = spreadsheet.getSheets()[0];
  firstSheet.setName('Sessions Recording');
  usedNames['Sessions Recording'] = true;
  const sessHeaders = ['Site', 'Cohort', 'Sessions', 'GoPro Recording', 'Tascam Recording', 'Meeting Owl Recording', 'Field note'];
  firstSheet.getRange(1,1,1,sessHeaders.length).setValues([sessHeaders]);
  if (sessionRows.length) {
    firstSheet.getRange(2,1,sessionRows.length,3).setValues(sessionRows.map(function(r){ return r.slice(0,3); }));
    firstSheet.getRange(2,4,sessionRows.length,4).setFormulas(sessionRows.map(function(r){
      return [toHyperlinkFormula(r[3]), toHyperlinkFormula(r[4]), toHyperlinkFormula(r[5]), toHyperlinkFormula(r[6])];
    }));
  }
  firstSheet.getRange(1,1,1,sessHeaders.length).setBackground('#2563eb').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  firstSheet.setRowHeight(1, 70);
  firstSheet.setColumnWidths(1, sessHeaders.length, 178);

  Object.keys(groups).forEach(function(groupKey) {
    const group = groups[groupKey];
    const protocolItems = protocolByRollout[group.rolloutId] || [];

    const protocolNames = [];
    const seen = {};
    protocolItems.forEach(function(item) {
      const nm = String(item.instrumentName || item.name || '').trim();
      if (!nm) return;
      const norm = normalizeInstrumentName(nm);
      if (exclusion[norm]) return;
      if (!seen[nm]) { seen[nm] = true; protocolNames.push(nm); }
    });

    const headers = ['Participant Name', 'Site', 'Cohort'].concat(protocolNames);
    const rows = group.participants.map(function(p) {
      const row = [p.participantName, p.site, p.cohort];
      protocolNames.forEach(function(protocolName) {
        var link = '';
        if (linkMap[p.participantId] && linkMap[p.participantId][protocolName]) {
          link = linkMap[p.participantId][protocolName];
        } else {
          var nk = normalizeInstrumentName(protocolName);
          link = (linkMapNormalized[p.participantId] && linkMapNormalized[p.participantId][nk]) ? linkMapNormalized[p.participantId][nk] : '';
        }
        row.push(link);
      });
      return row;
    });

    const sheet = spreadsheet.insertSheet();
    const safeName = makeSafeSheetName(group.cohort, usedNames);
    sheet.setName(safeName);

    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2,1,rows.length,3).setValues(rows.map(function(r){ return r.slice(0,3); }));
      if (headers.length > 3) {
        sheet.getRange(2,4,rows.length,headers.length - 3).setFormulas(rows.map(function(r){
          return r.slice(3).map(function(link){ return toHyperlinkFormula(link); });
        }));
      }
    }
    sheet.getRange(1,1,1,headers.length).setBackground('#2563eb').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
    sheet.setRowHeight(1, 70);
    sheet.setColumnWidths(1, headers.length, 178);
    if (headers.length > 3) sheet.getRange(2,4,Math.max(rows.length,1),headers.length - 3).setHorizontalAlignment('center');
  });

  SpreadsheetApp.flush();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    return { success: false, message: 'Failed to generate export file' };
  }
  const blob = response.getBlob().setName('links_artifacts_export_' + new Date().toISOString().split('T')[0] + '.xlsx');
  const base64 = Utilities.base64Encode(blob.getBytes());
  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
  return { success: true, filename: blob.getName(), mimeType: blob.getContentType(), content: base64, message: 'Links and artifacts export generated' };
}

function exportSummaryReport(token, filters) {
  const currentUser = validateSession(token);
  if (!currentUser) {
    return { success: false, message: 'Unauthorized' };
  }

  let siteFilter = (filters && filters.site) ? filters.site : 'All';
  const rolloutFilter = (filters && filters.rolloutId) ? filters.rolloutId : 'All';

  if (currentUser.role === 'facilitator' && currentUser.site !== 'All') {
    siteFilter = currentUser.site;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const rolloutsSheet = ss.getSheetByName('StudyRollouts');

  if (!participantsSheet || !checklistSheet || !rolloutsSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const rolloutsData = rolloutsSheet.getDataRange().getValues();
  const rHeaders = rolloutsData[0];
  const participantsData = participantsSheet.getDataRange().getValues();
  const pHeaders = participantsData[0];
  const checklistData = checklistSheet.getDataRange().getValues();
  const cHeaders = checklistData[0];

  const rolloutEntries = rolloutsData.slice(1).filter(r => {
    const rolloutSite = r[rHeaders.indexOf('site')];
    const rolloutId = r[rHeaders.indexOf('rolloutId')];

    if (siteFilter && siteFilter !== 'All' && rolloutSite !== siteFilter) return false;
    if (rolloutFilter && rolloutFilter !== 'All' && rolloutId !== rolloutFilter) return false;
    return true;
  }).map(r => ({
    rolloutId: r[rHeaders.indexOf('rolloutId')],
    site: r[rHeaders.indexOf('site')],
    schoolName: r[rHeaders.indexOf('schoolName')],
    period: r[rHeaders.indexOf('period')],
    year: r[rHeaders.indexOf('year')]
  }));

  function getInstrumentStatsForRollout(rolloutId) {
    const participantIds = participantsData.slice(1)
      .filter(p => p[pHeaders.indexOf('rolloutId')] === rolloutId &&
        (!siteFilter || siteFilter === 'All' || p[pHeaders.indexOf('site')] === siteFilter))
      .map(p => p[pHeaders.indexOf('participantId')]);

    const counts = {};
    CONFIG.INSTRUMENTS.forEach(inst => {
      counts[inst.number] = { name: inst.name, completed: 0, total: 0 };
    });

    checklistData.slice(1).forEach(row => {
      const pid = row[cHeaders.indexOf('participantId')];
      if (!participantIds.includes(pid)) return;

      const instNum = row[cHeaders.indexOf('instrumentNumber')];
      const status = row[cHeaders.indexOf('status')];
      if (!counts[instNum]) return;

      counts[instNum].total++;
      if (status === 'completed') {
        counts[instNum].completed++;
      }
    });

    return CONFIG.INSTRUMENTS.map(inst => {
      const entry = counts[inst.number];
      const pct = entry.total > 0 ? Math.round((entry.completed / entry.total) * 100) : 0;
      return {
        name: inst.name,
        percentComplete: pct,
        percentIncomplete: 100 - pct
      };
    });
  }

  const spreadsheet = SpreadsheetApp.create('Summary Report ' + new Date().toISOString());
  const defaultSheet = spreadsheet.getSheets()[0];
  defaultSheet.setName('Summary');
  defaultSheet.getRange(1, 1).setValue('Summary Report');
  defaultSheet.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString());
  defaultSheet.autoResizeColumns(1, 2);

  rolloutEntries.forEach((cohort, index) => {
    const sheetName = (cohort.schoolName || 'Cohort') + ' ' + (cohort.period || '') + ' ' + (cohort.year || '');
    const safeName = sheetName.substring(0, 90) || 'Cohort ' + (index + 1);
    const sheet = spreadsheet.insertSheet(safeName);

    const instrumentStats = getInstrumentStatsForRollout(cohort.rolloutId);

    sheet.getRange('A1').setValue('Instrument');
    sheet.getRange('B1').setValue('Complete %');
    sheet.getRange('C1').setValue('Incomplete %');

    instrumentStats.forEach((inst, idx) => {
      const row = idx + 2;
      sheet.getRange(row, 1).setValue(inst.name);
      sheet.getRange(row, 2).setValue(inst.percentComplete / 100);
      sheet.getRange(row, 3).setValue(inst.percentIncomplete / 100);
    });

    sheet.getRange(1, 2, instrumentStats.length + 1, 2).setNumberFormat('0%');
    sheet.autoResizeColumns(1, 3);

    const ranges = [
      { start: 2, end: Math.min(11, instrumentStats.length + 1), positionRow: 2, positionCol: 5 },
      { start: 12, end: instrumentStats.length + 1, positionRow: 20, positionCol: 5 }
    ];

    ranges.forEach(range => {
      if (range.start > range.end || range.start > instrumentStats.length + 1) return;
      const dataRange = sheet.getRange(`A${range.start}:C${range.end}`);

      const chart = sheet.newChart()
        .addRange(dataRange)
        .setChartType(Charts.ChartType.COLUMN)
        .setOption('title', `${sheetName} – Instrument Completion`)
        .setOption('legend', { position: 'top' })
        .setOption('isStacked', false)
        .setOption('series', {
          0: { color: '#10b981', labelInLegend: 'Complete' },
          1: { color: '#ef4444', labelInLegend: 'Incomplete' }
        })
        .setPosition(range.positionRow, range.positionCol, 0, 0)
        .build();

      sheet.insertChart(chart);
    });
  });

  SpreadsheetApp.flush();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    return { success: false, message: 'Failed to export workbook: ' + response.getContentText() };
  }

  const blob = response.getBlob().setName('summary_report_' + new Date().toISOString().split('T')[0] + '.xlsx');
  const base64 = Utilities.base64Encode(blob.getBytes());

  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);

  return {
    success: true,
    file: base64,
    mimeType: MimeType.MICROSOFT_EXCEL,
    filename: 'summary_report_' + new Date().toISOString().split('T')[0] + '.xlsx',
    downloadMessage: 'Summary downloaded'
  };
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Generate a UUID
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Generate a random password
 */
function generateRandomPassword() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#$%';
  let password = '';
  for (let i = 0; i < 12; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

/**
 * Generate a session token
 */
function generateSessionToken() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

/**
 * Escape a value for CSV
 */
function csvEscape(value) {
  const str = value === null || value === undefined ? '' : String(value);
  return '"' + str.replace(/"/g, '""') + '"';
}

/**
 * Normalize session dates to plain text (YYYY-MM-DD)
 */
function normalizeSessionDateValue(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const dateOnlyMatch = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return dateOnlyMatch ? dateOnlyMatch[1] : text;
}

/**
 * Set a cell value while enforcing plain text format
 */
function setPlainTextCell(sheet, row, column, value) {
  sheet.getRange(row, column).setNumberFormat('@').setValue(value === undefined ? '' : value);
}

/**
 * Hash a password using SHA-256
 */
function hashPassword(password) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return hash.map(byte => {
    const hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/**
 * Verify a password against stored password
 * Using plain text comparison since the Google Sheet is highly protected
 */
function verifyPassword(password, storedPassword) {
  // Plain text comparison - storedPassword is stored as-is in the sheet
  return password === storedPassword;
}

/**
 * Format date for display
 */
function formatDate(dateString) {
  if (!dateString) return '';
  const date = parseDateOnlyString(dateString) || new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
}

/**
 * Format date and time for display
 */
function formatDateTime(dateString) {
  if (!dateString) return '';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}
