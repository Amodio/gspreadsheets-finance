/**
 * ECB EUR/USD reference rate fetcher (Sheets-safe)
 *
 * - Caches one year at a time using CacheService (24h TTL)
 * - Uses non-blocking LockService to prevent race conditions
 * - Safe for concurrent custom function execution
 */

const ECB_CACHE_TTL_S = 24 * 60 * 60; // 24 hours

/* ----------------- Normalize Sheets date ----------------- */
function _normalizeSheetsDate_(dateObj) {
  if (!(dateObj instanceof Date))
    throw new Error("Invalid date");

  return new Date(Date.UTC(
    dateObj.getFullYear(),
    dateObj.getMonth(),
    dateObj.getDate()
  ));
}

/* ----------------- Year cache (CacheService) ----------------- */
function _ecb_getYearCache_(year) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(`ecb_year_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _ecb_setYearCache_(year, data) {
  const cache = CacheService.getScriptCache();
  const key = `ecb_year_${year}`;
  cache.put(key, JSON.stringify(data), ECB_CACHE_TTL_S);

  // Track keys for cleanup
  let keys = cache.get("ecb_year_keys");
  keys = keys ? JSON.parse(keys) : [];
  if (!keys.includes(key)) keys.push(key);
  cache.put("ecb_year_keys", JSON.stringify(keys), ECB_CACHE_TTL_S);
}

/* ----------------- Fetch + parse ECB XML ----------------- */
function _ecb_fetchAll_() {
  const url =
    "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/" +
    "euro_reference_exchange_rates/html/usd.xml";

  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200)
    throw new Error("ECB HTTP error " + resp.getResponseCode());

  const doc = XmlService.parse(resp.getContentText());
  const root = doc.getRootElement();

  const daily = {};

  function walk(el) {
    if (el.getName() === "Obs") {
      const d = el.getAttribute("TIME_PERIOD")?.getValue();
      const v = el.getAttribute("OBS_VALUE")?.getValue();
      if (d && v) daily[d] = parseFloat(v);
    }
    el.getChildren().forEach(walk);
  }

  walk(root);

  return daily; // { "YYYY-MM-DD": rate }
}

/* ----------------- Fetch one year (non-blocking lock) ----------------- */
function _ecb_fetchYear_(year) {
  const lock = LockService.getScriptLock();

  // Non-blocking: if another execution is fetching, skip
  if (!lock.tryLock(0)) {
    return _ecb_getYearCache_(year) ?? {};
  }

  try {
    // Re-check cache after acquiring lock
    const cached = _ecb_getYearCache_(year);
    if (cached) return cached;

    const all = _ecb_fetchAll_();
    const yearly = {};

    for (const d in all) {
      if (d.startsWith(String(year))) {
        yearly[d] = all[d];
      }
    }

    _ecb_setYearCache_(year, yearly);
    return yearly;
  } finally {
    lock.releaseLock();
  }
}

/* ----------------- Clear ECB cache ----------------- */
function clearECBHistory() {
  const cache = CacheService.getScriptCache();
  const keysRaw = cache.get("ecb_year_keys");

  if (keysRaw) {
    const keys = JSON.parse(keysRaw);
    cache.removeAll(keys);
    cache.remove("ecb_year_keys");
  }

  Logger.log("ECB cache cleared.");
}

/* ----------------- Public Sheets function ----------------- */
/**
 * Returns the ECB EUR/USD reference rate for a given date.
 *
 * @param {Date} dateObj Date-formatted Sheets cell
 * @return {number|string} Exchange rate or "No data"
 * @customfunction
 */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date))
    throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");

  let yearly = _ecb_getYearCache_(year);
  if (!yearly) {
    yearly = _ecb_fetchYear_(year); // non-blocking, safe
  }

  return yearly[key] ?? "No data";
}
