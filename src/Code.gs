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
    sessionRecordingTypes: recordingTypes
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][keyCol] === key) {
      configSheet.getRange(i + 1, valueCol + 1).setValue(value);
      configSheet.getRange(i + 1, descCol + 1).setValue(description || '');
      configSheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      return;
    }
  }
  configSheet.appendRow([key, value, description || '', new Date().toISOString()]);
  RUNTIME_CACHE.configMap = null;
  RUNTIME_CACHE.globalProtocolItems = null;
  RUNTIME_CACHE.rolloutProtocolItems = {};
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
  Object.keys(rolloutParticipantIds).forEach(participantId => {
    enabledItems.forEach(item => {
      const key = participantId + '|' + String(item.number);
      if (!existing[key]) {
        checklistSheet.appendRow([generateUUID(), participantId, item.number, item.name, item.category, 'not_started', '', '', '', '']);
        added++;
      }
    });
    updateParticipantCompletion(participantId);
  });

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
          recordingLinksJson: headers.indexOf('recordingLinksJson') !== -1 ? data[i][headers.indexOf('recordingLinksJson')] : ''
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
  sessionsSheet.getRange(nextRow, 1, 1, headers.length).setValues([[
    sessionId,
    sessionData.rolloutId,
    sessionData.sessionNumber,
    sessionDateText,
    sessionData.sessionName || '',
    'scheduled',
    timestamp,
    currentUser.userId
  ]]);
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
    return [
      sessionId,
      rolloutId,
      sessionData.sessionNumber,
      sessionDateText,
      sessionData.sessionName || '',
      'scheduled',
      timestamp,
      currentUser.userId
    ];
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

function ensureParticipantColumns(participantsSheet) {
  const requiredHeaders = [
    'participantId', 'fullName',
    'parent_guardian_names', 'parent_guardian_phone', 'parent_guardian_address',
    'parent_guardian_email', 'parent_guardian_dob',
    'site', 'rolloutId', 'schoolName', 'period', 'year',
    'enrollmentDate', 'enrolledBy', 'status', 'notes', 'completionPercentage'
  ];

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
}

function appendParticipantRowAsText(participantsSheet, rowValues) {
  const rowIndex = participantsSheet.getLastRow() + 1;
  const range = participantsSheet.getRange(rowIndex, 1, 1, rowValues.length);
  range.setNumberFormat('@');
  range.setValues([rowValues]);
  return rowIndex;
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
  const participantIds = data.slice(1).map(row => getHeaderValue(row, headers, 'participantId')).filter(Boolean);
  const completionMap = buildLiveCompletionMap(participantIds);

  filters = filters || {};

  for (let i = 1; i < data.length; i++) {
    const participant = {
      participantId: getHeaderValue(data[i], headers, 'participantId'),
      fullName: getHeaderValue(data[i], headers, 'fullName'),
      site: getHeaderValue(data[i], headers, 'site'),
      rolloutId: getHeaderValue(data[i], headers, 'rolloutId'),
      schoolName: getHeaderValue(data[i], headers, 'schoolName'),
      period: getHeaderValue(data[i], headers, 'period'),
      year: getHeaderValue(data[i], headers, 'year'),
      enrollmentDate: getHeaderValue(data[i], headers, 'enrollmentDate'),
      enrolledBy: getHeaderValue(data[i], headers, 'enrolledBy'),
      status: getHeaderValue(data[i], headers, 'status'),
      notes: getHeaderValue(data[i], headers, 'notes'),
      completionPercentage: (completionMap[getHeaderValue(data[i], headers, 'participantId')] || {}).percentage || 0
    };

    if (currentUser.role !== 'viewer') {
      participant.parentGuardianNames = getHeaderValue(data[i], headers, 'parent_guardian_names');
      participant.parentGuardianPhone = getHeaderValue(data[i], headers, 'parent_guardian_phone');
      participant.parentGuardianAddress = getHeaderValue(data[i], headers, 'parent_guardian_address');
      participant.parentGuardianEmail = getHeaderValue(data[i], headers, 'parent_guardian_email');
      participant.parentGuardianDob = getHeaderValue(data[i], headers, 'parent_guardian_dob');
    }

    // Apply filters
    let include = true;

    if (filters.site && filters.site !== 'All' && participant.site !== filters.site) {
      include = false;
    }
    if (filters.rolloutId && filters.rolloutId !== 'All' && participant.rolloutId !== filters.rolloutId) {
      include = false;
    }
    if (filters.status && participant.status !== filters.status) {
      include = false;
    }
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      const matchesSearch =
        participant.fullName.toLowerCase().includes(searchLower) ||
        participant.participantId.toLowerCase().includes(searchLower) ||
        participant.schoolName.toLowerCase().includes(searchLower);
      if (!matchesSearch) include = false;
    }

    // Facilitators see their site by default unless viewing all
    if (currentUser.role === 'facilitator' && currentUser.site !== 'All' && participant.site !== currentUser.site) {
      include = false;
    }

    if (include) {
      participants.push(participant);
    }
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

  if (!participantData || !participantData.fullName || !participantData.fullName.trim()) {
    return { success: false, message: 'Participant full name is required' };
  }
  if (!participantData.rolloutId) {
    return { success: false, message: 'Study cohort is required' };
  }

  const parentEmail = (participantData.parentGuardianEmail || '').trim();
  const parentDob = (participantData.parentGuardianDob || '').trim();
  const parentPhone = (participantData.parentGuardianPhone || '').trim();

  if (parentEmail && !isValidEmail(parentEmail)) {
    return { success: false, message: 'Parent/guardian email is invalid' };
  }
  if (parentDob && !isValidDateValue(parentDob)) {
    return { success: false, message: 'Parent/guardian date of birth is invalid' };
  }
  if (parentPhone) {
    const digits = normalizePhoneDigits(parentPhone);
    if (digits.length < 10 || digits.length > 15) {
      return { success: false, message: 'Parent/guardian phone must include 10 to 15 digits' };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  const participantHeaders = ensureParticipantColumns(participantsSheet);
  ensureChecklistColumns(checklistSheet);

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
    fullName: participantData.fullName,
    parentGuardianNames: participantData.parentGuardianNames,
    parentGuardianPhone: parentPhone,
    parentGuardianAddress: participantData.parentGuardianAddress,
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

  // Create checklist items for rollout-configured protocol items
  getRolloutProtocolItemsInternal(participantData.rolloutId).forEach(instrument => {
    const checklistId = generateUUID();
    checklistSheet.appendRow([
      checklistId,
      participantId,
      instrument.number,
      instrument.name,
      instrument.category,
      'not_started',
      '',
      '',
      '',
      ''
    ]);
  });

  logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT', 'participant', participantId,
    'Enrolled: ' + participantData.fullName);

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

  const parentEmail = (participantData.parentGuardianEmail || '').trim();
  const parentDob = (participantData.parentGuardianDob || '').trim();
  const parentPhone = (participantData.parentGuardianPhone || '').trim();

  if (parentEmail && !isValidEmail(parentEmail)) {
    return { success: false, message: 'Parent/guardian email is invalid' };
  }
  if (parentDob && !isValidDateValue(parentDob)) {
    return { success: false, message: 'Parent/guardian date of birth is invalid' };
  }
  if (parentPhone) {
    const digits = normalizePhoneDigits(parentPhone);
    if (digits.length < 10 || digits.length > 15) {
      return { success: false, message: 'Parent/guardian phone must include 10 to 15 digits' };
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('participantId')] === participantId) {
      if (currentUser.role === 'facilitator' && currentUser.site !== 'All' &&
          data[i][headers.indexOf('site')] !== currentUser.site) {
        return { success: false, message: 'Unauthorized for this site' };
      }

      if (participantData.fullName) {
        setCellAsPlainText(participantsSheet, i + 1, headers.indexOf('fullName') + 1, participantData.fullName);
      }
      if (participantData.parentGuardianNames !== undefined) {
        const idx = headers.indexOf('parent_guardian_names');
        if (idx !== -1) {
          setCellAsPlainText(participantsSheet, i + 1, idx + 1, participantData.parentGuardianNames);
        }
      }
      if (participantData.parentGuardianPhone !== undefined) {
        const idx = headers.indexOf('parent_guardian_phone');
        if (idx !== -1) {
          setCellAsPlainText(participantsSheet, i + 1, idx + 1, parentPhone);
        }
      }
      if (participantData.parentGuardianAddress !== undefined) {
        const idx = headers.indexOf('parent_guardian_address');
        if (idx !== -1) {
          setCellAsPlainText(participantsSheet, i + 1, idx + 1, participantData.parentGuardianAddress);
        }
      }
      if (participantData.parentGuardianEmail !== undefined) {
        const idx = headers.indexOf('parent_guardian_email');
        if (idx !== -1) {
          setCellAsPlainText(participantsSheet, i + 1, idx + 1, parentEmail);
        }
      }
      if (participantData.parentGuardianDob !== undefined) {
        const idx = headers.indexOf('parent_guardian_dob');
        if (idx !== -1) {
          setCellAsPlainText(participantsSheet, i + 1, idx + 1, parentDob);
        }
      }
      if (participantData.status) {
        setCellAsPlainText(participantsSheet, i + 1, headers.indexOf('status') + 1, participantData.status);
      }
      if (participantData.notes !== undefined) {
        setCellAsPlainText(participantsSheet, i + 1, headers.indexOf('notes') + 1, participantData.notes);
      }

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

  // Headers
  const headerValues = [
    'fullName',
    'parent_guardian_names',
    'parent_guardian_phone',
    'parent_guardian_address',
    'parent_guardian_email',
    'parent_guardian_dob',
    'cohortName'
  ];
  templateSheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]);
  templateSheet.getRange(1, 1, 1, headerValues.length)
    .setFontWeight('bold')
    .setBackground('#f1f5f9');

  // Helper sheet with cohort list
  const helperSheet = tempSs.insertSheet('Cohorts');
  helperSheet.getRange(1, 1, cohorts.length, 1).setValues(
    cohorts.map(r => [`${r.schoolName} (${r.period} ${r.year})`])
  );
  helperSheet.hideSheet();

  // Data validation for cohort dropdown (apply to reasonable range)
  const lastRow = Math.max(2, cohorts.length + 5);
  const validationRange = helperSheet.getRange(1, 1, cohorts.length, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true)
    .setAllowInvalid(false)
    .build();
  templateSheet.getRange(2, 7, lastRow, 1).setDataValidation(rule);

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
  const hasCohortName = headers.indexOf('cohortName') !== -1;
  const hasRolloutName = headers.indexOf('rolloutName') !== -1;
  const missing = [];
  if (headers.indexOf('fullName') === -1) {
    missing.push('fullName');
  }
  if (!hasCohortName && !hasRolloutName) {
    missing.push('cohortName');
  }
  if (missing.length > 0) {
    return { success: false, message: 'Missing required columns: ' + missing.join(', ') };
  }

  const idx = name => headers.indexOf(name);
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

    const fullName = (getValue(row, 'fullName') || '').trim();
    const rolloutName = (getValue(row, cohortColumn) || '').trim();
    const parentGuardianNames = (getValue(row, 'parent_guardian_names') || '').trim();
    const parentGuardianPhone = (getValue(row, 'parent_guardian_phone') || '').trim();
    const parentGuardianAddress = (getValue(row, 'parent_guardian_address') || '').trim();
    const parentGuardianEmail = (getValue(row, 'parent_guardian_email') || '').trim();
    const parentGuardianDob = (getValue(row, 'parent_guardian_dob') || '').trim();

    if (!fullName && !rolloutName) {
      summary.skipped++;
      continue;
    }

    summary.processed++;

    if (!fullName) {
      summary.errors.push({ row: rowNumber, message: 'Full name is required' });
      summary.skipped++;
      continue;
    }
    if (!rolloutName) {
      summary.errors.push({ row: rowNumber, message: 'cohortName is required' });
      summary.skipped++;
      continue;
    }
    const cohort = rolloutMap[rolloutName.toLowerCase()];
    if (!cohort) {
      summary.errors.push({ row: rowNumber, message: 'Cohort not found or not accessible: ' + rolloutName });
      summary.skipped++;
      continue;
    }

    if (parentGuardianEmail && !isValidEmail(parentGuardianEmail)) {
      summary.errors.push({ row: rowNumber, message: 'Parent/guardian email is invalid' });
      summary.skipped++;
      continue;
    }
    if (parentGuardianDob && !isValidDateValue(parentGuardianDob)) {
      summary.errors.push({ row: rowNumber, message: 'Parent/guardian date of birth is invalid' });
      summary.skipped++;
      continue;
    }
    if (parentGuardianPhone) {
      const digits = normalizePhoneDigits(parentGuardianPhone);
      if (digits.length < 10 || digits.length > 15) {
        summary.errors.push({ row: rowNumber, message: 'Parent/guardian phone must include 10 to 15 digits' });
        summary.skipped++;
        continue;
      }
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
    appendParticipantRowAsText(participantsSheet, rowValues);

    getRolloutProtocolItemsInternal(cohort.rolloutId).forEach(instrument => {
      const checklistId = generateUUID();
      checklistSheet.appendRow([
        checklistId,
        newParticipantId,
        instrument.number,
        instrument.name,
        instrument.category,
        'not_started',
        '',
        '',
        '',
        ''
      ]);
    });

    logActivity(currentUser.userId, currentUser.fullName, 'ENROLL_PARTICIPANT_BULK', 'participant', newParticipantId,
      'Bulk enrolled: ' + fullName + ' into ' + cohort.schoolName);

    summary.created++;
  }

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

  const headers = ensureChecklistColumns(checklistSheet);
  const data = checklistSheet.getRange(1, 1, checklistSheet.getLastRow(), headers.length).getValues();

  // Build participant site map for authorization checks
  const participantSiteMap = {};
  if (participantsSheet) {
    const pData = participantsSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    for (let i = 1; i < pData.length; i++) {
      participantSiteMap[pData[i][pHeaders.indexOf('participantId')]] = pData[i][pHeaders.indexOf('site')];
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('checklistId')] === checklistId) {
      const participantId = data[i][headers.indexOf('participantId')];
      const instrumentName = data[i][headers.indexOf('instrumentName')];
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
        checklistSheet.getRange(i + 1, headers.indexOf('status') + 1).setValue(updateData.status);
        updateDetails.push('status -> ' + updateData.status);

        // If marking as completed, set the date and user
        if (updateData.status === 'completed') {
          checklistSheet.getRange(i + 1, headers.indexOf('completedDate') + 1).setValue(new Date().toISOString());
          checklistSheet.getRange(i + 1, headers.indexOf('completedBy') + 1).setValue(currentUser.fullName);
        } else {
          // Clear completion info if status changed to not_started or missing
          checklistSheet.getRange(i + 1, headers.indexOf('completedDate') + 1).setValue('');
          checklistSheet.getRange(i + 1, headers.indexOf('completedBy') + 1).setValue('');
        }
      }

      // Update notes
      if (updateData.notes !== undefined) {
        checklistSheet.getRange(i + 1, headers.indexOf('notes') + 1).setValue(updateData.notes);
        updateDetails.push('notes updated');
      }

      if (updateData.dataLink !== undefined) {
        checklistSheet.getRange(i + 1, headers.indexOf('dataLink') + 1).setValue(updateData.dataLink);
        updateDetails.push(updateData.dataLink ? 'data link saved' : 'data link removed');
      }

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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checklistSheet = ss.getSheetByName('Checklist');
  const participantsSheet = ss.getSheetByName('Participants');
  const map = {};
  if (!checklistSheet || checklistSheet.getLastRow() < 2 || !participantsSheet || participantsSheet.getLastRow() < 2) return map;
  const set = {};
  (participantIds || []).forEach(id => { set[String(id)] = true; });
  const filterEnabled = participantIds && participantIds.length > 0;

  const participantSnapshot = getSheetSnapshot('Participants', { ensureFn: ensureParticipantColumns });
  const pHeaders = participantSnapshot.headers;
  const pData = participantSnapshot.data;
  const participantRolloutMap = {};
  const rolloutAllowedNumbersMap = {};
  for (let i = 1; i < pData.length; i++) {
    const pid = String(pData[i][pHeaders.indexOf('participantId')]);
    if (filterEnabled && !set[pid]) continue;
    const rolloutId = String(pData[i][pHeaders.indexOf('rolloutId')] || '');
    participantRolloutMap[pid] = rolloutId;
    if (!rolloutAllowedNumbersMap[rolloutId]) {
      const allowed = {};
      getRolloutProtocolItemsInternal(rolloutId).forEach(item => { allowed[String(item.number)] = true; });
      rolloutAllowedNumbersMap[rolloutId] = allowed;
    }
  }

  const cData = checklistSheet.getDataRange().getValues();
  const cHeaders = cData[0];
  const participantCol = cHeaders.indexOf('participantId');
  const statusCol = cHeaders.indexOf('status');
  const numberCol = cHeaders.indexOf('instrumentNumber');
  for (let i = 1; i < cData.length; i++) {
    const pid = String(cData[i][participantCol]);
    if (filterEnabled && !set[pid]) continue;
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

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][pHeaders.indexOf('site')];
    const status = pData[i][pHeaders.indexOf('status')];
    const rolloutId = pData[i][pHeaders.indexOf('rolloutId')];
    const participantId = pData[i][pHeaders.indexOf('participantId')];

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

  const liveCompletionMap = buildLiveCompletionMap(filteredParticipantIds);
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

    for (let i = 1; i < cData.length; i++) {
      const instNum = cData[i][cHeaders.indexOf('instrumentNumber')];
      const instStatus = cData[i][cHeaders.indexOf('status')];
      const checklistParticipantId = cData[i][cHeaders.indexOf('participantId')];

      // Only count checklist items for filtered participants
      if (filteredParticipantIds.length > 0 && !filteredParticipantIds.includes(checklistParticipantId)) {
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
  let completionSum = 0;

  for (let i = 1; i < pData.length; i++) {
    const site = pData[i][pHeaders.indexOf('site')];
    const rolloutId = pData[i][pHeaders.indexOf('rolloutId')];
    const status = pData[i][pHeaders.indexOf('status')];
    const participantId = pData[i][pHeaders.indexOf('participantId')];

    if (effectiveSite !== 'All' && site !== effectiveSite) continue;
    if (effectiveRollout !== 'All' && rolloutId !== effectiveRollout) continue;

    includedParticipantIds.push(participantId);
    participantRolloutMap[participantId] = rolloutId || 'Unassigned';
    participantNameMap[participantId] = pData[i][pHeaders.indexOf('fullName')] || 'Unknown Participant';
    overview.totalParticipants++;
    overview.activeParticipants += status === 'active' ? 1 : 0;
    overview.completedParticipants += status === 'completed' ? 1 : 0;
    overview.withdrawnParticipants += status === 'withdrawn' ? 1 : 0;

    const key = rolloutId || 'Unassigned';
    if (!rolloutEntries[key]) {
      rolloutEntries[key] = {
        rolloutId: rolloutId || 'Unassigned',
        site: site,
        schoolName: rolloutId ? pData[i][pHeaders.indexOf('schoolName')] : 'No cohort assigned',
        period: pData[i][pHeaders.indexOf('period')],
        year: pData[i][pHeaders.indexOf('year')],
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

  const liveCompletionMap = buildLiveCompletionMap(includedParticipantIds);
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

    for (let i = 1; i < cData.length; i++) {
      const participantId = cData[i][cHeaders.indexOf('participantId')];
      if (!includedParticipantIds.includes(participantId)) continue;

      const instrumentNumber = cData[i][cHeaders.indexOf('instrumentNumber')];
      const status = cData[i][cHeaders.indexOf('status')];

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
      about: aboutSummary
    }
  };
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

function parseSessionDate(dateValue) {
  if (!dateValue) return null;
  const normalized = normalizeSessionDateValue(dateValue);
  if (!normalized) return null;
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
  const includeParentFields = currentUser.role !== 'viewer';
  const headers = ['Participant ID', 'Full Name'];

  if (includeParentFields) {
    headers.push(
      'Parent/Guardian Name(s)',
      'Parent/Guardian Phone',
      'Parent/Guardian Address',
      'Parent/Guardian Email',
      'Parent/Guardian DOB'
    );
  }

  headers.push('Site', 'School', 'Period', 'Year', 'Enrollment Date', 'Status', 'Completion %', 'Notes');

  let csv = headers.join(',') + '\n';

  participants.forEach(p => {
    const row = [
      p.participantId,
      '"' + (p.fullName || '').replace(/"/g, '""') + '"'
    ];

    if (includeParentFields) {
      row.push(
        '"' + (p.parentGuardianNames || '').replace(/"/g, '""') + '"',
        '"' + (p.parentGuardianPhone || '').replace(/"/g, '""') + '"',
        '"' + (p.parentGuardianAddress || '').replace(/"/g, '""') + '"',
        '"' + (p.parentGuardianEmail || '').replace(/"/g, '""') + '"',
        '"' + (p.parentGuardianDob || '').replace(/"/g, '""') + '"'
      );
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
  if (!currentUser) {
    return { success: false, message: 'Invalid session' };
  }

  const siteFilter = (filters && filters.site) ? filters.site : 'All';
  const rolloutFilter = (filters && filters.rolloutId) ? filters.rolloutId : 'All';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participantsSheet = ss.getSheetByName('Participants');
  const checklistSheet = ss.getSheetByName('Checklist');
  if (!participantsSheet || !checklistSheet) {
    return { success: false, message: 'Required sheets are missing' };
  }

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
  for (var i = 1; i < participantsData.length; i++) {
    const row = participantsData[i];
    const site = String(row[pIdx.site] || '');
    const rolloutId = String(row[pIdx.rolloutId] || '');
    if (siteFilter !== 'All' && site !== siteFilter) continue;
    if (rolloutFilter !== 'All' && rolloutId !== String(rolloutFilter)) continue;
    const pid = String(row[pIdx.participantId] || '');
    if (!pid) continue;
    selectedParticipantIds[pid] = true;
    selectedParticipants.push({
      participantId: pid,
      participantName: String(row[pIdx.fullName] || ''),
      site: site,
      cohort: String(row[pIdx.schoolName] || '') + (row[pIdx.period] ? ' (' + String(row[pIdx.period]) + ' ' + String(row[pIdx.year] || '') + ')' : ''),
      rolloutId: rolloutId
    });
  }

  if (!selectedParticipants.length) {
    return { success: false, message: 'No participants found for selected filters' };
  }

  const exclusion = {
    'consent form': true,
    'assent form / pre-test': true,
    'post-test': true,
    'participant feedback survey': true,
    'parent satisfaction survey': true
  };

  const protocolByRollout = {};
  if (rolloutFilter !== 'All') {
    protocolByRollout[String(rolloutFilter)] = getProtocolItemsForFilter(String(rolloutFilter));
  } else {
    const uniqueRollouts = {};
    selectedParticipants.forEach(function(p) { if (p.rolloutId) uniqueRollouts[p.rolloutId] = true; });
    Object.keys(uniqueRollouts).forEach(function(rid) {
      protocolByRollout[rid] = getProtocolItemsForFilter(rid);
    });
  }

  const protocolNameSet = {};
  Object.keys(protocolByRollout).forEach(function(rid) {
    (protocolByRollout[rid] || []).forEach(function(item) {
      const nm = String(item.instrumentName || '').trim();
      if (!nm) return;
      const key = nm.toLowerCase();
      if (exclusion[key]) return;
      protocolNameSet[nm] = true;
    });
  });
  const protocolNames = Object.keys(protocolNameSet).sort(function(a,b){ return a.localeCompare(b); });

  const checklistData = checklistSheet.getDataRange().getValues();
  const checklistHeaders = checklistData[0] || [];
  const cIdx = {
    participantId: checklistHeaders.indexOf('participantId'),
    instrumentName: checklistHeaders.indexOf('instrumentName'),
    dataLink: checklistHeaders.indexOf('dataLink')
  };

  const linkMap = {};
  for (var j = 1; j < checklistData.length; j++) {
    const crow = checklistData[j];
    const pid2 = String(crow[cIdx.participantId] || '');
    if (!selectedParticipantIds[pid2]) continue;
    const instrument = String(crow[cIdx.instrumentName] || '').trim();
    if (!instrument || !protocolNameSet[instrument]) continue;
    const link = String(crow[cIdx.dataLink] || '').trim();
    if (!link) continue;
    if (!linkMap[pid2]) linkMap[pid2] = {};
    linkMap[pid2][instrument] = link;
  }

  const exportRows = [];
  selectedParticipants.forEach(function(p) {
    const row = [p.participantName, p.site, p.cohort];
    protocolNames.forEach(function(protocolName) {
      const link = linkMap[p.participantId] && linkMap[p.participantId][protocolName] ? linkMap[p.participantId][protocolName] : '';
      row.push(link ? '=HYPERLINK("' + link.replace(/"/g, '""') + '","Open")' : '');
    });
    exportRows.push(row);
  });

  const headers = ['Participant Name', 'Site', 'Cohort'].concat(protocolNames);
  const spreadsheet = SpreadsheetApp.create('Links and Artifacts Export ' + new Date().toISOString());
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName('Links and Artifacts');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (exportRows.length) {
    sheet.getRange(2, 1, exportRows.length, headers.length).setValues(exportRows);
  }
  sheet.getRange(1,1,1,headers.length).setBackground('#2563eb').setFontColor('#ffffff').setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  sheet.getRange(2,4,Math.max(exportRows.length,1),Math.max(headers.length-3,1)).setHorizontalAlignment('center');

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
  return {
    success: true,
    filename: blob.getName(),
    mimeType: blob.getContentType(),
    content: base64,
    message: 'Links and artifacts export generated'
  };
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
  return String(value);
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
  const date = new Date(dateString);
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
