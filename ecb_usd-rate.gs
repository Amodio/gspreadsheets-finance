/**
 * ECB EUR/USD fetcher (optimized)
 * - Past years: cache once if missing
 * - Current year: flush and refresh daily via trigger
 * - Lock only around caching, no retry
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

function _deleteYearStore_(year) {
  PropertiesService.getScriptProperties().deleteProperty(`ecb_usd_${year}`);
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

/* ----------------- Daily prefetch trigger ----------------- */
function prefetchECBData() {
  console.log("Fetching ECB data...");

  const currentYear = new Date().getUTCFullYear();
  const allData = _fetchECBXML_();
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    console.log(`Lock acquired for caching`);

    // Group data by year
    const dataByYear = {};
    for (const d in allData) {
      const year = parseInt(d.slice(0, 4), 10);
      if (!dataByYear[year]) dataByYear[year] = {};
      dataByYear[year][d] = allData[d];
    }

    // Cache each year
    for (const yearStr in dataByYear) {
      const year = parseInt(yearStr, 10);

      // Past years: cache only if missing
      if (year < currentYear && !_getYearStore_(year)) {
        console.log(`Caching past year ${year}`);
        _setYearStore_(year, dataByYear[year]);
      }

      // Current year: always cache (already flushed above)
      if (year === currentYear) {
        _deleteYearStore_(currentYear);
        console.log(`Caching current year ${currentYear}`);
        _setYearStore_(currentYear, dataByYear[year]);
      }
    }

    console.log("Prefetch trigger finished");
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Automatically create the daily prefetch trigger when the script is installed
 * Don’t run at midnight otherwise it times out as:
 * - Google executes a lot of App Scripts at midnight
 * - ECB updates at about 4pm GMT+1.
 */
/* ----------------- Install daily trigger + first prefetch ----------------- */
function onInstall(e) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "prefetchECBData") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("prefetchECBData")
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(0)
    .create();
  console.log("Daily prefetch trigger installed");

  prefetchECBData();
}

/* ----------------- Public Sheets function ----------------- */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");

  const yearly = _getYearStore_(year);
  return yearly ? (yearly[key] ?? "No data") : "No data for this year";
}

/* ----------------- Manual flush of the cache ----------------- */
function flushAllECBCache() {
  const props = PropertiesService.getScriptProperties();
  const keys = Object.keys(props.getProperties()).filter(k => k.startsWith('ecb_usd_'));
  keys.forEach(k => props.deleteProperty(k));
  console.log("Deleted ECB cache keys:", keys.join(", "));
}
