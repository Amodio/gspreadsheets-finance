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
function _prefetchECBData_() {
  console.log("Prefetch trigger started");

  const currentYear = new Date().getUTCFullYear();

  // 1️⃣ Flush current year's cache BEFORE downloading XML
  console.log(`Flushing current year ${currentYear} cache`);
  _deleteYearStore_(currentYear);

  // 2️⃣ Fetch full ECB XML
  const allData = _fetchECBXML_();

  // 3️⃣ Lock around caching (no retry)
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
        console.log(`Caching current year ${currentYear}`);
        _setYearStore_(currentYear, dataByYear[year]);
      }
    }

    console.log("Prefetch trigger finished");
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ----------------- Public Sheets function ----------------- */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");

  const yearly = _getYearStore_(year);
  return yearly ? (yearly[key] ?? "No data") : "No data";
}

/**
 * Automatically create the daily prefetch trigger when the script is installed
 */
/* ----------------- Install daily trigger ----------------- */
function onInstall(e) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "_prefetchECBData_") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("_prefetchECBData_")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  console.log("Daily prefetch trigger installed");
}

/* ----------------- Manual flush & repopulate all the cache ----------------- */
function flushAllECBCache() {
  const props = PropertiesService.getScriptProperties();
  const keys = Object.keys(props.getProperties()).filter(k => k.startsWith('ecb_usd_'));
  keys.forEach(k => props.deleteProperty(k));
  console.log("Deleted ECB cache keys:", keys.join(", "));
  _prefetchECBData_();
}

/* ----------------- Test ----------------- */
function testECB() {
  console.log(ECB_USD_RATE(new Date("2026-01-02")));
}

function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    console.log("No triggers found.");
    return;
  }
  
  triggers.forEach((t, i) => {
    console.log(`Trigger ${i + 1}:`);
    console.log(`  ID: ${t.getUniqueId()}`);
    console.log(`  Handler Function: ${t.getHandlerFunction()}`);
    console.log(`  Trigger Type: ${t.getEventType()}`); // e.g., CLOCK, ON_OPEN
    console.log(`  Source ID: ${t.getTriggerSourceId()}`);
    console.log(`  Source: ${t.getTriggerSource()}`);
    console.log(`  Creation Time: ${t.getTriggerSourceId()}`);
  });
}
