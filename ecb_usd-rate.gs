/**
 * ECB Euro/USD reference rate fetcher (concurrency safe)
 *
 * - Only fetches once per day per year
 * - Uses PropertiesService for global shared storage
 * - Prevents race conditions using a lock
 * - Custom function: =ECB_USD_RATE(A1)
 */

const ECB_TTL_SECONDS = 24 * 60 * 60;  // 24 hours
const ECB_TTL_MS = ECB_TTL_SECONDS * 1000;
const LOCK_TIMEOUT_MS = 5000;          // Lock expires after 5 seconds

/* ----------------- Helper: normalize Sheets date ----------------- */
function _normalizeSheetsDate_(dateObj) {
  if (!(dateObj instanceof Date)) throw new Error("Invalid date");
  return new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
}

/* ----------------- ECB XML parsing ----------------- */
function _ecb_parseXML_(xml) {
  const document = XmlService.parse(xml);
  const root = document.getRootElement();

  function findObs(element) {
    let list = [];
    if (element.getName() === "Obs") list.push(element);
    element.getChildren().forEach(c => list = list.concat(findObs(c)));
    return list;
  }

  const observations = findObs(root);
  if (!observations.length) throw new Error("No <Obs> entries found");

  const daily = {};
  observations.forEach(obs => {
    const d = obs.getAttribute("TIME_PERIOD")?.getValue();
    const v = obs.getAttribute("OBS_VALUE")?.getValue();
    if (d && v) daily[d] = parseFloat(v);
  });

  return daily;
}

/* ----------------- Global Storage (PropertiesService) ----------------- */
function _getYearStore_(year) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(`ecb_usd_${year}`);
  return raw ? JSON.parse(raw) : null;
}

function _setYearStore_(year, data) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(
    `ecb_usd_${year}`,
    JSON.stringify({ ts: Date.now(), data })
  );
}

/* ----------------- Distributed Lock ----------------- */

function _tryLock_(year) {
  const key = `lock_ecb_${year}`;
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  const existing = props.getProperty(key);
  if (existing && now - parseInt(existing, 10) < LOCK_TIMEOUT_MS) {
    return false;  // lock held
  }

  props.setProperty(key, now.toString());
  return true;
}

function _unlock_(year) {
  PropertiesService.getScriptProperties().deleteProperty(`lock_ecb_${year}`);
}

/* ----------------- Fetch ECB XML ----------------- */
function _fetchECBXML_() {
  const url = "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/" +
              "euro_reference_exchange_rates/html/usd.xml";
  const resp = UrlFetchApp.fetch(url);

  if (resp.getResponseCode() !== 200)
    throw new Error("ECB HTTP error " + resp.getResponseCode());

  return _ecb_parseXML_(resp.getContentText());
}

/* ----------------- Load year data, using shared cache + lock ----------------- */
function _loadYearData_(year) {
  let entry = _getYearStore_(year);

  // Fresh? Return immediately.
  if (entry && Date.now() - entry.ts < ECB_TTL_MS) {
    return entry.data;
  }

  // Otherwise, try to acquire lock (1 other fetch happening allowed)
  if (_tryLock_(year)) {
    try {
      // We own the lock — fetch fresh data
      const all = _fetchECBXML_();

      // Filter by year
      const yearData = {};
      Object.keys(all).forEach(d => {
        if (d.startsWith(String(year))) yearData[d] = all[d];
      });

      _setYearStore_(year, yearData);
      return yearData;
    } finally {
      _unlock_(year);
    }
  }

  // Another execution is fetching → wait briefly & reload
  Utilities.sleep(250);
  entry = _getYearStore_(year);

  if (entry) return entry.data;

  // Fallback (should rarely fire)
  return {};
}

/* ----------------- Public Sheets function ----------------- */
/**
 * Returns the ECB EUR/USD reference rate for a given date.
 * @param {Date} dateObj Sheets date
 * @return {number|string} Exchange rate or "No data"
 * @customfunction
 */
function ECB_USD_RATE(dateObj) {
  if (!(dateObj instanceof Date))
    throw new Error("Argument must be a date");

  const utc = _normalizeSheetsDate_(dateObj);
  const year = utc.getUTCFullYear();
  const key = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");

  const yearly = _loadYearData_(year);
  return yearly[key] ?? "No data";
}
