// ============================================================
// Shared user-preference storage for taskpane.html.
// Prefers Office.context.roamingSettings (syncs per-user across
// devices); falls back to localStorage when roaming settings
// aren't available (e.g. opened directly in a browser).
// ============================================================
const APP_SETTINGS_KEYS = { theme: "sp_theme", history: "sp_history" };
const HISTORY_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days
const HISTORY_MAX_ENTRIES = 50;

function loadAppSetting(key, fallback) {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      const v = Office.context.roamingSettings.get(key);
      if (v != null) return v;
    }
  } catch (e) {}
  try {
    const v = localStorage.getItem(key);
    if (v != null) return v;
  } catch (e) {}
  return fallback;
}

function saveAppSetting(key, value) {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      Office.context.roamingSettings.set(key, value);
      Office.context.roamingSettings.saveAsync();
    }
  } catch (e) {}
  try {
    localStorage.setItem(key, value);
  } catch (e) {}
}

// ============================================================
// Filing history: { name, path, timestamp } entries, newest first,
// pruned to the last 30 days / HISTORY_MAX_ENTRIES on every write.
// ============================================================
function loadHistoryEntries() {
  let list;
  try {
    list = JSON.parse(loadAppSetting(APP_SETTINGS_KEYS.history, "[]"));
  } catch (e) {
    list = [];
  }
  if (!Array.isArray(list)) list = [];
  const cutoff = Date.now() - HISTORY_MAX_AGE_MS;
  return list
    .filter((e) => e && e.timestamp && new Date(e.timestamp).getTime() >= cutoff)
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

function addHistoryEntry(entry) {
  const list = loadHistoryEntries();
  list.unshift(entry);
  saveAppSetting(APP_SETTINGS_KEYS.history, JSON.stringify(list.slice(0, HISTORY_MAX_ENTRIES)));
}
