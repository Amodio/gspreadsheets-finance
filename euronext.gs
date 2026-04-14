/**
 * Euronext historical price fetcher
 */

const EURONEXT_RATE_LIMIT = 5;           // max calls per RATE_PERIOD_MS
const EURONEXT_RATE_PERIOD_MS = 60 * 1000; // 1 minute
const EURONEXT_CACHE_TTL = 24 * 60 * 60; // 24 hours in seconds

const EURONEXT_KEY = "24ayqVo7yJma";

/**
 * Decrypt a CryptoJS-format AES-CBC payload.
 * Replaces the broken Utilities.CipherAlgorithm approach.
 */
function _euronext_decrypt_(json, password) {
  if (typeof json === "string") json = JSON.parse(json);

  const cache = CacheService.getScriptCache();
  let cryptoJsSrc = cache.get("cryptojs_src");
  if (!cryptoJsSrc) {
    const resp = UrlFetchApp.fetch(
      "https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.2.0/crypto-js.min.js"
    );
    cryptoJsSrc = resp.getContentText();
    cache.put("cryptojs_src", cryptoJsSrc, 21600);
  }
  eval(cryptoJsSrc);

  // CryptoJSAesJson format:
  //   ct  → Base64
  //   iv  → HEX
  //   s   → HEX
  var CryptoJSAesJson = {
    parse: function(jsonStr) {
      var j = JSON.parse(jsonStr);
      var cipherParams = CryptoJS.lib.CipherParams.create({
        ciphertext: CryptoJS.enc.Base64.parse(j.ct)
      });
      if (j.iv) cipherParams.iv   = CryptoJS.enc.Hex.parse(j.iv);
      if (j.s)  cipherParams.salt = CryptoJS.enc.Hex.parse(j.s);
      return cipherParams;
    }
  };

  const cipherParams = CryptoJSAesJson.parse(JSON.stringify(json));

  const decrypted = CryptoJS.AES.decrypt(cipherParams, password, {
    iv:      cipherParams.iv,
    salt:    cipherParams.salt,
    mode:    CryptoJS.mode.CBC,
    padding: CryptoJS.pad.Pkcs7
  });

  const result = decrypted.toString(CryptoJS.enc.Utf8);
  if (!result) throw new Error("Decryption produced empty result — wrong key or format?");
  return result;
}

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

  const dateOnlyRegex = /<td class="historical-time">\s*<span>(\d{2}\/\d{2}\/\d{4})<\/span>/g;
  const closeOnlyRegex = /<td class="historical-close[^"]*">([\d.,]+)<\/td>/g;

  const dates = [];
  const closes = [];
  let m;

  while ((m = dateOnlyRegex.exec(html)) !== null) dates.push(m[1]);
  while ((m = closeOnlyRegex.exec(html)) !== null) closes.push(m[1]);

  if (dates.length !== closes.length) {
    throw new Error(`Euronext parse mismatch: ${dates.length} dates vs ${closes.length} closes`);
  }

  for (let i = 0; i < dates.length; i++) {
    const p = dates[i].split("/");
    daily[`${p[2]}-${p[1]}-${p[0]}`] = parseFloat(closes[i].replace(/,/g, ""));
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

  const body = resp.getContentText();

  let decrypted;
    try {
      const parsed = JSON.parse(body);
      if (parsed.ct && parsed.iv && parsed.s) {
        decrypted = _euronext_decrypt_(parsed, EURONEXT_KEY);
      } else {
        decrypted = body;
      }
    } catch (e) {
      throw new Error("Decrypt/parse failed: " + e);
    }

    // Decrypted value is a JSON-encoded string — unwrap it
    let html;
    try {
      html = JSON.parse(decrypted);
    } catch(e) {
      html = decrypted;
    }

    return _euronext_parseEuronextHTML_(html);
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
  const utcDate = _normalizeSheetsDate_(dateObj);
  const key = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd");

  let yearly = _euronext_getYearCache_(ticker, year);

  // Fetch if cache is null OR if the object is empty
  if (!yearly || Object.keys(yearly).length === 0) {
    Logger.log(`Cache miss or empty for ${ticker} ${year}. Fetching...`);
    yearly = _euronext_fetchYearData_(ticker, year);
    
    // Safety: Only cache if we actually got data back
    if (yearly && Object.keys(yearly).length > 0) {
      _euronext_setYearCache_(ticker, year, yearly);
    }
  }

  return yearly[key] !== undefined ? yearly[key] : "No data";
}

function test_EURONEXT_HIST() {
  Logger.log(EURONEXT_HIST("LU1681048804-XPAR", new Date(2026, 3, 13)));
}
