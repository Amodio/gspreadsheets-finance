/**
 * Euronext historical price fetcher
 */

const EURONEXT_RATE_LIMIT = 5;           // max calls per RATE_PERIOD_MS
const EURONEXT_RATE_PERIOD_MS = 60 * 1000; // 1 minute
const EURONEXT_CACHE_TTL = 24 * 60 * 60; // 24 hours in seconds

/* ---------- Shared function with polygon.gs ---------- */
function _normalizeSheetsDate_(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Invalid date");

  // Extract year, month, day
  const y = dateObj.getFullYear();
  const m = dateObj.getMonth(); // 0-based
  const d = dateObj.getDate();

  // Construct a date in UTC midnight
  return new Date(Date.UTC(y, m, d));
}

/* ---------- Parse Euronext HTML Table ---------- */
function _euronext_parseEuronextHTML_(html) {
  const daily = {};
  const regex = /<tr class="[^"]*">[\s\S]*?<td class="historical-time">[\s\S]*?<span>(\d{2}\/\d{2}\/\d{4})<\/span>[\s\S]*?<td class="historical-close[^"]*">([\d.,]+)<\/td>/g;
  let match;
  while ((match = regex.exec(html)) !== null) {
    const dateParts = match[1].split("/");
    const dateStr = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`; // yyyy-MM-dd
    const close = parseFloat(match[2].replace(",", ""));
    daily[dateStr] = close;
  }

  return daily;
}

/* ---------- 24-HOUR CACHE USING CacheService ---------- */
function _euronext_getYearCache_(ticker, year) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(`euronext_${ticker}_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _euronext_setYearCache_(ticker, year, data) {
  const cache = CacheService.getScriptCache();
  const key = `euronext_${ticker}_${year}`;
  cache.put(key, JSON.stringify(data), EURONEXT_CACHE_TTL);

  // Track this key
  let keys = cache.get("euronext_cache_keys");
  keys = keys ? JSON.parse(keys) : [];
  if (!keys.includes(key)) keys.push(key);
  cache.put("euronext_cache_keys", JSON.stringify(keys), EURONEXT_CACHE_TTL);
}

/* ---------- RATE LIMITER ---------- */
function _euronext_waitForRateLimit_() {
  const cache = CacheService.getScriptCache();
  const now = Date.now();

  let ts = cache.get("euronext_api_timestamps");
  ts = ts ? JSON.parse(ts) : [];

  ts = ts.filter(t => now - t < EURONEXT_RATE_PERIOD_MS);

  if (ts.length >= EURONEXT_RATE_LIMIT) {
    const waitTime = EURONEXT_RATE_PERIOD_MS - (now - ts[0]) + 50;
    Utilities.sleep(waitTime);
    return _euronext_waitForRateLimit_();
  }

  ts.push(now);
  cache.put("euronext_api_timestamps", JSON.stringify(ts), Math.ceil(EURONEXT_RATE_PERIOD_MS/1000));
}

/* ---------- Fetch Year Data from Euronext ---------- */
function _euronext_fetchYearData_(ticker, year) {
  _euronext_waitForRateLimit_();

  const start = `${year}-01-01`;
  const end = `${year}-12-31`;

  const payload = {
    adjusted: "Y",
    startdate: start,
    enddate: end,
    nbSession: 250  // max sessions per year, adjust if needed
  };

  const options = {
    method: "post",
    payload: payload,
    muteHttpExceptions: true
  };

  const url = `https://live.euronext.com/en/ajax/getHistoricalPricePopup/${ticker}`;
  Logger.log("Fetching URL: " + url);
  const resp = UrlFetchApp.fetch(url, options);
  Logger.log("Response code: " + resp.getResponseCode());
  Logger.log("Response body: " + resp.getContentText().substring(0, 500));

  if (resp.getResponseCode() !== 200)
    throw new Error("Euronext Error " + resp.getResponseCode());

  return _euronext_parseEuronextHTML_(resp.getContentText());
}

/* ---------- Clear Cache ---------- */
function clearEuronextHistory() {
  const cache = CacheService.getScriptCache();

  // Remove tracked year caches
  const keysRaw = cache.get("euronext_cache_keys");
  if (keysRaw) {
    const keys = JSON.parse(keysRaw);
    cache.removeAll(keys);
    Logger.log(`Removing keys from the cache: ${keys}.`);
    cache.remove("euronext_cache_keys");
  }

  // Remove rate-limit key
  cache.remove("euronext_api_timestamps");

  Logger.log("All Euronext cache cleared.");
}

/* ---------- PUBLIC FUNCTION ---------- */
/**
 * Returns the daily closing price for a ticker on a given date.
 *
 * This function caches yearly data and supports EU tickers.
 * Use a date-formatted cell for the date argument.
 *
 * @param {string} ticker The ticker symbol, e.g., "SPY"
 * @param {Date} date The date to retrieve the closing price for (date-formatted cell)
 * @return {number|string} Closing price for the given date, or "No data" if not available
 * @customfunction
 */
function EURONEXT_HIST(ticker, dateObj) {
  if (!(dateObj instanceof Date))
    throw new Error("Date must be a valid date object");

  const year = dateObj.getUTCFullYear();
  // Normalize date from Sheets cell
  const utcDate = _normalizeSheetsDate_(dateObj);
  const key = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd");
  //Logger.log(`year=${year} key=${key}.`);

  let yearly = _euronext_getYearCache_(ticker, year);
  if (!yearly) {
    yearly = _euronext_fetchYearData_(ticker, year);
    _euronext_setYearCache_(ticker, year, yearly);
  }

  return yearly[key] !== undefined ? yearly[key] : "No data";
}
