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
  for (let i = 1; i < data.length; i++) {
    result[data[i][keyCol]] = data[i][valueCol];
  }
  RUNTIME_CACHE.configMap = result;
  return result;
}

function getProtocolItemsForFilter(rolloutFilter) {
  if (rolloutFilter && rolloutFilter !== 'All') {
    return getRolloutProtocolItemsInternal(rolloutFilter);
  }
  return getGlobalProtocolItems();
}
