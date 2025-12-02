/**
 * ECB Euro/USD reference rate fetcher
 * 
 * This script fetches historical Euro/USD rates from ECB in XML (SDMX Compact)
 * and caches yearly data in Google Apps Script cache for 24 hours.
 * 
 * @customfunction
 */

const ECB_CACHE_TTL = 24 * 60 * 60; // 24 hours in seconds

/* ---------- Normalize Sheets Date ---------- */
function _normalizeSheetsDate_(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Invalid date");
  return new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
}

/* ---------- Parse ECB XML ---------- */
function _ecb_parseXML_(xml) {
  const document = XmlService.parse(xml);
  const root = document.getRootElement();

  // Recursive helper to find all <Obs> elements
  function findObsElements(element) {
    let obsList = [];
    if (element.getName() === "Obs") {
      obsList.push(element);
    }
    element.getChildren().forEach(child => {
      obsList = obsList.concat(findObsElements(child));
    });
    return obsList;
  }

  const observations = findObsElements(root);
  if (observations.length === 0) throw new Error("No <Obs> elements found in ECB XML");

  const daily = {};
  observations.forEach(obs => {
    const dateAttr = obs.getAttribute("TIME_PERIOD");
    const valueAttr = obs.getAttribute("OBS_VALUE");
    if (dateAttr && valueAttr) {
      const date = dateAttr.getValue();
      const value = parseFloat(valueAttr.getValue());
      daily[date] = value;
    }
  });

  return daily;
}

/* ---------- 24-HOUR CACHE USING CacheService ---------- */
function _ecb_getYearCache_(year) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(`ecb_usd_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _ecb_setYearCache_(year, data) {
  const cache = CacheService.getScriptCache();
  const key = `ecb_usd_${year}`;
  cache.put(key, JSON.stringify(data), ECB_CACHE_TTL);

  // Track cache keys
  let keys = cache.get("ecb_cache_keys");
  keys = keys ? JSON.parse(keys) : [];
  if (!keys.includes(key)) keys.push(key);
  cache.put("ecb_cache_keys", JSON.stringify(keys), ECB_CACHE_TTL);
}

/* ---------- Fetch ECB Year Data ---------- */
function _ecb_fetchYearData_(year) {
  const url = "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/usd.xml";
  const resp = UrlFetchApp.fetch(url);

  if (resp.getResponseCode() !== 200)
    throw new Error("ECB fetch error: " + resp.getResponseCode());

  const daily = _ecb_parseXML_(resp.getContentText());

  // Filter only this year
  const yearData = {};
  Object.keys(daily).forEach(date => {
    if (date.startsWith(year.toString())) yearData[date] = daily[date];
  });

  return yearData;
}

/* ---------- Clear ECB Cache ---------- */
function clearECBHistory() {
  const cache = CacheService.getScriptCache();
  const keysRaw = cache.get("ecb_cache_keys");

  if (keysRaw) {
    const keys = JSON.parse(keysRaw);
    cache.removeAll(keys);
    cache.remove("ecb_cache_keys");
    Logger.log(`Removed ECB cache keys: ${keys}`);
  }

  Logger.log("All ECB cache cleared.");
}

/* ---------- PUBLIC FUNCTION ---------- */
/**
 * Returns the ECB Euro/USD exchange rate for a given date.
 *
 * This caches yearly data in CacheService for 24 hours.
 * 
 * @param {Date} dateObj A date-formatted cell
 * @return {number|string} Exchange rate for the given date, or "No data"
 * @customfunction
 */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Date must be a valid date object");

  const utcDate = _normalizeSheetsDate_(dateObj);
  const year = utcDate.getUTCFullYear();
  const key = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd");

  let yearly = _ecb_getYearCache_(year);
  if (!yearly) {
    yearly = _ecb_fetchYearData_(year);
    _ecb_setYearCache_(year, yearly);
  }

  return yearly[key] !== undefined ? yearly[key] : "No data";
}
