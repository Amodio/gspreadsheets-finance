/**
 * ECB EUR/USD reference rate fetcher (fetch-on-demand, concurrency-safe, with logging)
 *
 * - Fetches data only if missing
 * - Past years: fetch once
 * - Current year: fetch missing days
 * - Throws for today/future dates
 * - Logs each step to help diagnose execution time
 * - Flush all cache with logging
 */

const LOCK_TIMEOUT_MS = 5000; // 5s lock wait

/* ----------------- Helpers ----------------- */
function _normalizeSheetsDate_(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Invalid date");
  return new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
}

function _getYearStore_(year) {
  const raw = PropertiesService.getScriptProperties().getProperty(`ecb_usd_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _setYearStore_(year, data) {
  PropertiesService.getScriptProperties().setProperty(
    `ecb_usd_${year}`,
    JSON.stringify(data)
  );
}

/* ----------------- Fetch ECB XML ----------------- */
function _fetchECBXML_() {
  console.time("ECB fetch time");
  const url =
    "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/" +
    "euro_reference_exchange_rates/html/usd.xml";

  const resp = UrlFetchApp.fetch(url);
  if (resp.getResponseCode() !== 200) throw new Error("ECB HTTP error " + resp.getResponseCode());

  const document = XmlService.parse(resp.getContentText());
  const root = document.getRootElement();

  function findObs(el) {
    let out = [];
    if (el.getName() === "Obs") out.push(el);
    el.getChildren().forEach(c => out = out.concat(findObs(c)));
    return out;
  }

  const observations = findObs(root);
  const daily = {};
  observations.forEach(obs => {
    const d = obs.getAttribute("TIME_PERIOD")?.getValue();
    const v = obs.getAttribute("OBS_VALUE")?.getValue();
    if (d && v) daily[d] = parseFloat(v);
  });

  console.timeEnd("ECB fetch time");
  return daily;
}

/* ----------------- Load year data (with lock and logging) ----------------- */
function _loadYearData_(year, requestedKey) {
  const start = Date.now();
  console.log(`_loadYearData_ started for year=${year}, requestedKey=${requestedKey}`);

  const todayKey = Utilities.formatDate(_normalizeSheetsDate_(new Date()), "UTC", "yyyy-MM-dd");
  if (requestedKey >= todayKey) throw new Error("ECB reference rate not available for today or future dates");

  // Check cache
  let stored = _getYearStore_(year);
  if (stored && requestedKey in stored) {
    console.log(`Cache hit for ${requestedKey}, returning immediately`);
    console.log(`_loadYearData_ finished in ${Date.now() - start} ms`);
    return stored;
  }

  console.log("Cache miss, acquiring lock...");
  const lock = LockService.getScriptLock();
  const maxRetries = 5;
  const retryDelay = 500; // ms

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const lockStart = Date.now();
      lock.waitLock(LOCK_TIMEOUT_MS);
      console.log(`Lock acquired in ${Date.now() - lockStart} ms (attempt ${attempt + 1})`);

      // Double-check cache after lock
      stored = _getYearStore_(year);
      if (stored && requestedKey in stored) {
        console.log(`Cache populated during wait, returning immediately`);
        console.log(`_loadYearData_ finished in ${Date.now() - start} ms`);
        return stored;
      }

      // Fetch ECB XML
      console.log("Fetching ECB XML...");
      const all = _fetchECBXML_();
      console.log("ECB XML fetched, storing year data...");

      const yearData = stored || {};
      for (const d in all) {
        if (d.startsWith(String(year))) yearData[d] = all[d];
      }
      _setYearStore_(year, yearData);

      console.log(`Year ${year} stored, total duration ${Date.now() - start} ms`);
      return yearData;

    } catch (e) {
      console.warn(`Lock attempt ${attempt + 1} failed: ${e.message}`);
      if (attempt === maxRetries - 1) throw e;
      Utilities.sleep(retryDelay);
    } finally {
      try { lock.releaseLock(); } catch(_) {}
    }
  }
}

/* ----------------- Manual flush all cache ----------------- */
function flushAllECBCache() {
  const props = PropertiesService.getScriptProperties();
  const keys = Object.keys(props.getProperties()).filter(k => k.startsWith('ecb_usd_'));

  if (keys.length === 0) {
    console.log("No ECB cache to delete.");
    return;
  }

  keys.forEach(k => props.deleteProperty(k));
  console.log("Deleted ECB cache keys:", keys.join(", "));
}

/* ----------------- Public Sheets function ----------------- */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");

  const yearly = _loadYearData_(year, key);
  if (!(key in yearly)) throw new Error("No ECB data for requested date");

  return yearly[key];
}
