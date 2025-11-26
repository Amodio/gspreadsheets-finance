/**
 * Use the Polygon.io / massive.com stocks API free plan.
 */

const POLYGON_API_KEY = "_YOUR_API_KEY_";
const RATE_LIMIT = 5;           // max calls per RATE_PERIOD_MS
const RATE_PERIOD_MS = 60 * 1000; // 1 minute
const CACHE_TTL_S = 24 * 60 * 60; // TTL 24h

/* ---------- 24-HOUR CACHE ---------- */
function _poly_getYearCache_(ticker, year) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(`yearcache_${ticker}_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _poly_setYearCache_(ticker, year, data) {
  const cache = CacheService.getScriptCache();
  const key = `yearcache_${ticker}_${year}`;
  cache.put(key, JSON.stringify(data), CACHE_TTL_S);

  // Track the cache keys
  let keys = cache.get("yearcache_keys");
  keys = keys ? JSON.parse(keys) : [];
  if (!keys.includes(key)) keys.push(key);
  cache.put("yearcache_keys", JSON.stringify(keys), CACHE_TTL_S);
}

/* ---------- RATE LIMITER ---------- */
function _poly_waitForRateLimit_() {
  const cache = CacheService.getScriptCache();
  const now = Date.now();

  // Get timestamps from cache
  let ts = cache.get("api_timestamps");
  ts = ts ? JSON.parse(ts) : [];

  // Remove old timestamps
  ts = ts.filter(t => now - t < RATE_PERIOD_MS);

  if (ts.length >= RATE_LIMIT) {
    const waitTime = RATE_PERIOD_MS - (now - ts[0]) + 50; // small buffer
    Utilities.sleep(waitTime);
    return _poly_waitForRateLimit_(); // retry
  }

  // Record this call
  ts.push(now);
  cache.put("api_timestamps", JSON.stringify(ts), Math.ceil(RATE_PERIOD_MS / 1000));
}

/* ---------- Fetch Entire Year ---------- */
function _poly_fetchYearData_(ticker, year) {
  _poly_waitForRateLimit_();

  const start = `${year}-01-01`;
  const end = `${year}-12-31`;

  const url = `https://api.polygon.io/v2/aggs/ticker/${ticker}/range/1/day/${start}/${end}?adjusted=true&sort=asc&limit=5000&apiKey=${POLYGON_API_KEY}`;

  Logger.log("Fetching URL: " + url);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log("Response code: " + resp.getResponseCode());
  Logger.log("Response body: " + resp.getContentText().substring(0, 500));
  if (resp.getResponseCode() !== 200) {
    throw new Error("Polygon Error " + resp.getResponseCode());
  }

  const data = JSON.parse(resp.getContentText());
  Logger.log("Number of results: " + (data.results ? data.results.length : 0));
  if (!data || !data.results) return {};

  // Map: "2023-01-09" → close price
  const daily = {};
  for (const p of data.results) {
    const date = new Date(p.t).toISOString().slice(0, 10);
    daily[date] = p.c;
  }
  return daily;
}

/* ---------- Clear Cache ---------- */
function clearPolygonHistory() {
  const cache = CacheService.getScriptCache();

  // Remove tracked year caches
  const keysRaw = cache.get("yearcache_keys");
  if (keysRaw) {
    const keys = JSON.parse(keysRaw);
    Logger.log(`Removing keys from the cache: ${keys}.`);
    cache.removeAll(keys);
    cache.remove("yearcache_keys");
  }

  // Remove rate-limit keys
  cache.remove("api_timestamps");

  Logger.log("All POLY_HIST cache cleared.");
}

/* ---------- PUBLIC FUNCTION ---------- */
/**
 * Returns the daily closing price for a ticker on a given date.
 *
 * This function caches yearly data and supports US tickers.
 * Use a date-formatted cell for the date argument.
 *
 * @param {string} ticker The ticker symbol, e.g., "SPY"
 * @param {Date} date The date to retrieve the closing price for (date-formatted cell)
 * @return {number|string} Closing price for the given date, or "No data" if not available
 * @customfunction
 */
function POLY_HIST(ticker, dateObj) {
  if (!(dateObj instanceof Date))
    throw new Error("Date must be a valid date object");

  const year = dateObj.getUTCFullYear();
  // Normalize date from Sheets cell
  const utcDate = _normalizeSheetsDate_(dateObj);
  const key = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd");

  let yearly = _poly_getYearCache_(ticker, year);
  if (!yearly) {
    yearly = _poly_fetchYearData_(ticker, year); // queued automatically
    _poly_setYearCache_(ticker, year, yearly);
  }

  return yearly[key] !== undefined ? yearly[key] : "No data";
}
