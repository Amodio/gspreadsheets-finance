/**
 * Optimized ECB EUR/USD fetcher
 * - Lock only when fetching XML
 * - Daily prefetch trigger for missing years
 * - Refresh current year if past date missing
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

/* ----------------- Load year data ----------------- */
function _loadYearData_(year) {
  // Check cache first
  let stored = _getYearStore_(year);
  if (stored) {
    console.log(`Year ${year} already cached, returning immediately`);
    return stored;
  }

  console.log(`Year ${year} missing, acquiring lock to fetch...`);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);

    // Re-check cache after lock
    stored = _getYearStore_(year);
    if (stored) {
      console.log(`Year ${year} cached while waiting, returning`);
      return stored;
    }

    // Fetch ECB XML
    console.log("Fetching ECB XML...");
    const all = _fetchECBXML_();
    const yearData = {};
    for (const d in all) if (d.startsWith(String(year))) yearData[d] = all[d];

    _setYearStore_(year, yearData);
    console.log(`Year ${year} cached successfully`);
    return yearData;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

/* ----------------- Refresh current year if requested date missing ----------------- */
function _ensureCurrentYearData_(requestedKey) {
  const todayKey = Utilities.formatDate(_normalizeSheetsDate_(new Date()), "UTC", "yyyy-MM-dd");
  if (requestedKey >= todayKey) throw new Error("ECB reference rate not available for today or future dates");

  const currentYear = _normalizeSheetsDate_(new Date()).getUTCFullYear();
  let currentYearData = _getYearStore_(currentYear);

  if (!currentYearData) {
    // Year not cached at all → fetch
    console.log(`Current year ${currentYear} not cached. Fetching...`);
    currentYearData = _loadYearData_(currentYear);
    return currentYearData;
  }

  // Check if any later date exists
  const laterDatesExist = Object.keys(currentYearData).some(d => d > requestedKey);
  if (laterDatesExist) {
    console.log(`Requested date ${requestedKey} has no data, but later dates exist. Returning "No data" without fetching.`);
    return currentYearData;
  }

  // Otherwise, requested past date is missing and no later data exists → fetch
  console.log(`Refreshing current year ${currentYear} for missing requested date ${requestedKey}`);
  currentYearData = _loadYearData_(currentYear);
  return currentYearData;
}

/* ----------------- Public Sheets function ----------------- */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");
  const todayKey = Utilities.formatDate(_normalizeSheetsDate_(new Date()), "UTC", "yyyy-MM-dd");

  // Past years: fetch only if missing
  if (year < new Date().getUTCFullYear()) {
    const yearly = _loadYearData_(year);
    return yearly[key] ?? "No data";
  }

  // Current year: ensure requested past date is in cache
  if (year === new Date().getUTCFullYear()) {
    const yearly = _ensureCurrentYearData_(key);
    return yearly[key] ?? "No data";
  }

  throw new Error("ECB reference rate not available for future years");
}

/* ----------------- Daily prefetch trigger ----------------- */
function prefetchMissingYears() {
  const props = PropertiesService.getScriptProperties();
  const cachedYears = Object.keys(props.getProperties())
    .filter(k => k.startsWith('ecb_usd_'))
    .map(k => parseInt(k.replace('ecb_usd_', '')));

  const currentYear = new Date().getUTCFullYear();
  const lastCachedYear = cachedYears.length ? Math.max(...cachedYears) : 0;

  for (let year = lastCachedYear + 1; year <= currentYear; year++) {
    console.log(`Prefetching missing year ${year}`);
    _loadYearData_(year);
  }
}

/* ----------------- Install daily trigger ----------------- */
function createDailyPrefetchTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "prefetchMissingYears") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("prefetchMissingYears")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  console.log("Daily prefetch trigger installed");
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

/* ----------------- Test ----------------- */
function testECB() {
  console.log(ECB_USD_RATE(new Date("2026-01-02")));
}
