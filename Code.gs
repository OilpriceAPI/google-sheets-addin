/**
 * OilPriceAPI Google Sheets reference implementation.
 *
 * This repository is not published in the Google Workspace Marketplace.
 * Product facts: https://api.oilpriceapi.com/product-facts.json
 */

const API_BASE_URL = 'https://api.oilpriceapi.com/v1';
const KEY_PROPERTY = 'OILPRICEAPI_KEY';
const MAX_BATCH_CODES = 25;
const PRICING_URL = 'https://www.oilpriceapi.com/pricing';
const SIGNUP_URL = 'https://www.oilpriceapi.com/auth/signup';
const CACHE_TTL_SECONDS = {
  latest: 300,
  history: 3600,
  futures: 300,
  rigCount: 3600
};

// Conversion support is intentionally narrower than API catalog access.
const COMMODITY_MAP = {
  'BRENT_CRUDE_USD': { type: 'BRENT_CRUDE_OIL', unit: 'barrel' },
  'WTI_USD': { type: 'WTI_CRUDE_OIL', unit: 'barrel' },
  'NATURAL_GAS_USD': { type: 'NATURAL_GAS', unit: 'MBtu' },
  'NATURAL_GAS_GBP': { type: 'NATURAL_GAS', unit: 'therm' },
  'DUTCH_TTF_EUR': { type: 'NATURAL_GAS', unit: 'MWh' },
  'COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  'CAPP_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'PRB_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'ILLINOIS_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'NEWCASTLE_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  'COKING_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  'CME_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'NYMEX_APPALACHIAN_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'NYMEX_WESTERN_RAIL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' }
};

// Reference conversion factors. Verify suitability for the source dataset.
const HEAT_CONTENT = {
  'BRENT_CRUDE_OIL': 5.8,
  'WTI_CRUDE_OIL': 5.8,
  'NATURAL_GAS': 1.037,
  'COAL_BITUMINOUS': 24.0
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OilPriceAPI')
    .addItem('Configure API Key', 'showSidebar')
    .addItem('Fetch Latest Available Prices', 'showFetchDialog')
    .addItem('Convert to $/MMBtu', 'convertToMBtu')
    .addSeparator()
    .addItem('Fetch Bunker Prices (Data Connector)', 'fetchDataConnectorPrices')
    .addItem('Futures Formula Help', 'showFuturesInfo')
    .addItem('Rig Count Formula Help', 'showRigCountInfo')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('OilPriceAPI Reference')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'OilPriceAPI Google Sheets Reference',
    'Version 1.1.0\n\n' +
      'Source-timestamped energy price data. Dataset access and freshness vary.\n\n' +
      'This reference implementation is not published in the Google Workspace Marketplace.\n\n' +
      'Website: https://www.oilpriceapi.com\n' +
      'Docs: https://docs.oilpriceapi.com',
    ui.ButtonSet.OK
  );
}

function getApiKey_() {
  return PropertiesService.getUserProperties().getProperty(KEY_PROPERTY);
}

function requireApiKey_() {
  const apiKey = getApiKey_();
  if (!apiKey) {
    throw new Error(`Configure an API key from OilPriceAPI > Configure API Key, or create one at ${SIGNUP_URL}.`);
  }
  return apiKey;
}

function saveApiKey(apiKey) {
  if (typeof apiKey !== 'string' || !apiKey.trim()) {
    throw new Error('API key is required.');
  }
  PropertiesService.getUserProperties().setProperty(KEY_PROPERTY, apiKey.trim());
  return { success: true, message: 'API key saved in your Apps Script user properties.' };
}

function deleteApiKey() {
  PropertiesService.getUserProperties().deleteProperty(KEY_PROPERTY);
  return { success: true, message: 'Stored API key deleted.' };
}

function getApiKeyStatus() {
  return { configured: Boolean(getApiKey_()) };
}

function normalizeCode_(value, label) {
  const code = String(value || '').trim().toUpperCase();
  if (!code) {
    throw new Error(`${label || 'Code'} is required.`);
  }
  if (!/^[A-Z0-9_:-]+$/.test(code)) {
    throw new Error(`${label || 'Code'} contains unsupported characters.`);
  }
  return code;
}

function requestJson_(path, apiKey) {
  let response;
  try {
    response = UrlFetchApp.fetch(`${API_BASE_URL}${path}`, {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
  } catch (error) {
    throw new Error('OilPriceAPI request failed or timed out. Retry later.');
  }

  const statusCode = response.getResponseCode();
  if (statusCode === 401) {
    throw new Error(`Invalid or revoked API key. Replace it from ${SIGNUP_URL}.`);
  }
  if (statusCode === 402 || statusCode === 403) {
    throw new Error(`This account cannot access the requested dataset. Review ${PRICING_URL}.`);
  }
  if (statusCode === 429) {
    throw new Error('OilPriceAPI rate or quota limit reached. Wait for the current limit window before retrying.');
  }
  if (statusCode === 404) {
    throw new Error('The requested OilPriceAPI resource was not found. Check the code and endpoint.');
  }
  if (statusCode >= 500) {
    throw new Error(`OilPriceAPI is temporarily unavailable (HTTP ${statusCode}). Retry later.`);
  }
  if (statusCode !== 200) {
    throw new Error(`OilPriceAPI request failed with HTTP ${statusCode}.`);
  }

  let body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error('OilPriceAPI returned malformed JSON for a successful request.');
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    throw new Error('OilPriceAPI returned a non-object successful response.');
  }
  return body;
}

function extractPriceRecords_(body, subject) {
  const payload = body.data;
  let records = null;
  if (Array.isArray(payload)) {
    records = payload;
  } else if (payload && typeof payload === 'object') {
    if (Array.isArray(payload.prices)) {
      records = payload.prices;
    } else if (payload.price && typeof payload.price === 'object') {
      records = [payload.price];
    } else if ('code' in payload || 'price' in payload) {
      records = [payload];
    }
  }
  if (!Array.isArray(records) || records.length === 0 || records.some((item) => !item || typeof item !== 'object')) {
    throw new Error(`OilPriceAPI returned no usable ${subject || 'price'} in a successful response.`);
  }
  return records;
}

function extractDataArray_(body, key, subject) {
  const payload = body.data;
  const records = payload && typeof payload === 'object' ? payload[key] : null;
  if (!Array.isArray(records) || records.length === 0 || records.some((item) => !item || typeof item !== 'object')) {
    throw new Error(`OilPriceAPI returned no usable ${subject} in a successful response.`);
  }
  return records;
}

function sourceTimestamp_(record, subject) {
  const timestamp = record.created_at || record.updated_at || record.timestamp || record.date;
  if (typeof timestamp !== 'string' || !timestamp.trim() || !Number.isFinite(new Date(timestamp).getTime())) {
    throw new Error(`${subject} is missing a valid source timestamp.`);
  }
  return timestamp;
}

function validatePriceRecord_(record, subject) {
  const code = normalizeCode_(record.code || record.symbol, `${subject} code`);
  const price = Number(record.price);
  if (!Number.isFinite(price)) {
    throw new Error(`${subject} ${code} is missing a finite price.`);
  }
  if (typeof record.currency !== 'string' || !record.currency.trim()) {
    throw new Error(`${subject} ${code} is missing currency.`);
  }
  if (typeof record.unit !== 'string' || !record.unit.trim()) {
    throw new Error(`${subject} ${code} is missing unit.`);
  }
  return {
    code,
    price,
    currency: record.currency.trim(),
    unit: record.unit.trim(),
    source: typeof record.source === 'string' ? record.source : '',
    timestamp: sourceTimestamp_(record, `${subject} ${code}`)
  };
}

function getCachedValue_(cacheKey, maxAgeSeconds) {
  const cache = CacheService.getUserCache();
  const raw = cache.get(cacheKey);
  if (!raw) return null;

  let envelope;
  try {
    envelope = JSON.parse(raw);
  } catch (error) {
    cache.remove(cacheKey);
    return null;
  }
  if (
    !envelope ||
    typeof envelope.cachedAt !== 'number' ||
    !Object.prototype.hasOwnProperty.call(envelope, 'value') ||
    Date.now() - envelope.cachedAt > maxAgeSeconds * 1000
  ) {
    cache.remove(cacheKey);
    return null;
  }
  return envelope.value;
}

function putCachedValue_(cacheKey, value, ttlSeconds) {
  CacheService.getUserCache().put(
    cacheKey,
    JSON.stringify({ cachedAt: Date.now(), value }),
    ttlSeconds
  );
}

function getLatestRecord_(commodityCode) {
  const code = normalizeCode_(commodityCode, 'Commodity code');
  const cacheKey = `latest_${code}`;
  const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.latest);
  if (cached) return cached;

  const body = requestJson_(`/prices/latest?by_code=${encodeURIComponent(code)}`, requireApiKey_());
  const record = validatePriceRecord_(extractPriceRecords_(body, 'price')[0], 'Price record');
  putCachedValue_(cacheKey, record, CACHE_TTL_SECONDS.latest);
  return record;
}

function testConnection() {
  if (!getApiKey_()) {
    return { success: false, message: `Configure an API key first at ${SIGNUP_URL}.` };
  }
  try {
    const body = requestJson_('/prices/latest?by_code=BRENT_CRUDE_USD', getApiKey_());
    validatePriceRecord_(extractPriceRecords_(body, 'price')[0], 'Price record');
    return { success: true, message: 'Connection and response schema verified.' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getUserInfo() {
  if (!getApiKey_()) {
    return { tier: 'none', limit: null, used: null };
  }
  try {
    const body = requestJson_('/users/me', getApiKey_());
    const data = body.data && typeof body.data === 'object' ? body.data : body;
    const tier = typeof data.tier === 'string' ? data.tier : (typeof data.plan === 'string' ? data.plan : null);
    const limit = Number(data.request_limit);
    const used = Number(data.requests_this_month);
    if (!tier || !Number.isFinite(limit) || !Number.isFinite(used)) {
      return { tier: 'unknown', limit: null, used: null };
    }
    return { tier, limit, used };
  } catch (error) {
    return { tier: 'unknown', limit: null, used: null, message: error.message };
  }
}

function fetchLatestPrices(commodityCodes) {
  if (!Array.isArray(commodityCodes) || commodityCodes.length === 0) {
    throw new Error('Select at least one commodity code.');
  }
  if (commodityCodes.length > MAX_BATCH_CODES) {
    throw new Error(`Select at most ${MAX_BATCH_CODES} commodity codes per refresh.`);
  }
  const codes = [...new Set(commodityCodes.map((code) => normalizeCode_(code, 'Commodity code')))];
  const body = requestJson_(
    `/prices/latest?by_code=${encodeURIComponent(codes.join(','))}`,
    requireApiKey_()
  );
  const prices = extractPriceRecords_(body, 'prices').map((record) => validatePriceRecord_(record, 'Price record'));
  writeToDataSheet(prices);
  return { success: true, count: prices.length };
}

function writeToDataSheet(prices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data');
  if (!sheet) sheet = ss.insertSheet('Data');
  sheet.clear();

  const headers = [['Commodity Code', 'Price', 'Currency', 'Unit', 'Source', 'Source Timestamp', 'Retrieved At']];
  sheet.getRange(1, 1, 1, 7).setValues(headers);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
  const retrievedAt = new Date().toISOString();
  const rows = prices.map((record) => [
    record.code,
    record.price,
    record.currency,
    record.unit,
    record.source,
    record.timestamp,
    retrievedAt
  ]);
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0.00');
  sheet.autoResizeColumns(1, 7);
  sheet.activate();
}

/**
 * Latest available numeric value for a commodity code.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {number} API-provided numeric price.
 * @customfunction
 */
function OILPRICE(commodityCode) {
  return getLatestRecord_(commodityCode).price;
}

/**
 * Historical source timestamp and price pairs.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @param {number} days Requested lookback selector from 1 through 365.
 * @return {Array} Rows of [source timestamp, price].
 * @customfunction
 */
function OILPRICE_HISTORY(commodityCode, days) {
  const code = normalizeCode_(commodityCode, 'Commodity code');
  const requestedDays = days === undefined || days === null || days === '' ? 30 : Number(days);
  if (!Number.isInteger(requestedDays) || requestedDays < 1 || requestedDays > 365) {
    throw new Error('History days must be an integer from 1 through 365.');
  }
  const cacheKey = `history_${code}_${requestedDays}`;
  const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.history);
  if (cached) return cached;

  let endpoint = 'past_year';
  if (requestedDays <= 1) endpoint = 'past_day';
  else if (requestedDays <= 7) endpoint = 'past_week';
  else if (requestedDays <= 30) endpoint = 'past_month';

  const body = requestJson_(
    `/prices/${endpoint}?by_code=${encodeURIComponent(code)}`,
    requireApiKey_()
  );
  const records = extractPriceRecords_(body, 'historical price');
  const result = records.map((record) => {
    const normalized = validatePriceRecord_(record, 'Historical price record');
    return [normalized.timestamp, normalized.price];
  });
  putCachedValue_(cacheKey, result, CACHE_TTL_SECONDS.history);
  return result;
}

function fetchExchangeRates() {
  const body = requestJson_('/prices/latest?by_code=GBP_USD,EUR_USD', requireApiKey_());
  const records = extractPriceRecords_(body, 'exchange rates').map((record) => validatePriceRecord_(record, 'Exchange-rate record'));
  const gbp = records.find((record) => record.code === 'GBP_USD');
  const eur = records.find((record) => record.code === 'EUR_USD');
  if (!gbp || !eur) {
    throw new Error('OilPriceAPI response did not contain both GBP_USD and EUR_USD exchange rates.');
  }
  return { gbpUsd: gbp.price, eurUsd: eur.price };
}

function toUsd_(price, currency, rates) {
  if (currency === 'USD') return price;
  if (currency === 'GBP') return price * rates.gbpUsd;
  if (currency === 'GBp') return (price / 100) * rates.gbpUsd;
  if (currency === 'EUR') return price * rates.eurUsd;
  throw new Error(`No reference conversion is implemented for currency ${currency}.`);
}

function heatContent_(commodityInfo) {
  if (commodityInfo.unit === 'therm') return 0.1;
  if (commodityInfo.unit === 'MWh') return 3.412;
  if (commodityInfo.unit === 'MBtu') return 1;
  const factor = HEAT_CONTENT[commodityInfo.type];
  if (!Number.isFinite(factor)) {
    throw new Error(`No heat-content factor is configured for ${commodityInfo.type}.`);
  }
  return factor;
}

/**
 * Reference conversion to USD/MMBtu for codes in COMMODITY_MAP.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {number} Reference converted value.
 * @customfunction
 */
function OILPRICE_CONVERT(commodityCode) {
  const code = normalizeCode_(commodityCode, 'Commodity code');
  const commodityInfo = COMMODITY_MAP[code];
  if (!commodityInfo) {
    throw new Error('This code has no reference heat-content conversion mapping.');
  }
  const record = getLatestRecord_(code);
  const rates = record.currency === 'USD' ? null : fetchExchangeRates();
  return toUsd_(record.price, record.currency, rates) / heatContent_(commodityInfo);
}

function convertToMBtu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('No Data sheet found. Fetch prices first.');
    return;
  }
  const data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No price rows found. Fetch prices first.');
    return;
  }

  const needsRates = data.slice(1).some((row) => row[2] && row[2] !== 'USD');
  const rates = needsRates ? fetchExchangeRates() : null;
  const processRows = [];
  for (let index = 1; index < data.length; index += 1) {
    const [code, rawPrice, currency, unit] = data[index];
    if (!code || !COMMODITY_MAP[code]) continue;
    const price = Number(rawPrice);
    if (!Number.isFinite(price) || typeof currency !== 'string' || !currency) {
      throw new Error(`Data row ${index + 1} is missing a valid price or currency.`);
    }
    const commodityInfo = COMMODITY_MAP[code];
    const usdPrice = toUsd_(price, currency, rates);
    const heatContent = heatContent_(commodityInfo);
    processRows.push([code, price, currency, unit, usdPrice, heatContent, usdPrice / heatContent]);
  }
  if (processRows.length === 0) {
    throw new Error('No Data rows have a configured reference conversion mapping.');
  }

  let processSheet = ss.getSheetByName('Process');
  if (!processSheet) processSheet = ss.insertSheet('Process');
  processSheet.clear();
  const headers = [['Commodity', 'Original Price', 'Currency', 'API Unit', 'USD Price', 'Reference MMBtu/Unit', 'Reference USD/MMBtu']];
  processSheet.getRange(1, 1, 1, 7).setValues(headers);
  processSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e3f2fd');
  processSheet.getRange(2, 1, processRows.length, 7).setValues(processRows);
  processSheet.getRange(2, 2, processRows.length, 1).setNumberFormat('#,##0.00');
  processSheet.getRange(2, 5, processRows.length, 1).setNumberFormat('$#,##0.00');
  processSheet.getRange(2, 6, processRows.length, 1).setNumberFormat('0.000');
  processSheet.getRange(2, 7, processRows.length, 1).setNumberFormat('$#,##0.00');
  processSheet.autoResizeColumns(1, 7);
  processSheet.activate();
  SpreadsheetApp.getUi().alert(`Converted ${processRows.length} rows using the documented reference factors.`);
}

function showFetchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('FetchDialog')
    .setWidth(420)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fetch Latest Available Prices');
}

function validateBunkerRecord_(record) {
  const price = Number(record.price);
  if (!Number.isFinite(price)) throw new Error('Bunker record is missing a finite price.');
  if (typeof record.port !== 'string' || !record.port) throw new Error('Bunker record is missing port.');
  if (typeof record.fuel_type !== 'string' || !record.fuel_type) throw new Error('Bunker record is missing fuel type.');
  if (typeof record.currency !== 'string' || !record.currency) throw new Error('Bunker record is missing currency.');
  if (typeof record.unit !== 'string' || !record.unit) throw new Error('Bunker record is missing unit.');
  return {
    port: record.port,
    fuelType: record.fuel_type,
    price,
    currency: record.currency,
    unit: record.unit,
    region: typeof record.region === 'string' ? record.region : '',
    source: typeof record.source === 'string' ? record.source : '',
    timestamp: sourceTimestamp_(record, 'Bunker record')
  };
}

function fetchBunkerRecords_(query) {
  const body = requestJson_(`/prices/data-connector${query || ''}`, requireApiKey_());
  return extractPriceRecords_(body, 'bunker price').map(validateBunkerRecord_);
}

function fetchDataConnectorPrices() {
  try {
    const prices = fetchBunkerRecords_('');
    writeToDataConnectorSheet(prices);
    SpreadsheetApp.getUi().alert(`Fetched ${prices.length} source-timestamped bunker price records.`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Bunker price request failed: ${error.message}`);
  }
}

function writeToDataConnectorSheet(prices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bunker Prices');
  if (!sheet) sheet = ss.insertSheet('Bunker Prices');
  sheet.clear();
  const headers = [['Port', 'Fuel Type', 'Price', 'Currency', 'Unit', 'Region', 'Source', 'Source Timestamp', 'Retrieved At']];
  sheet.getRange(1, 1, 1, 9).setValues(headers);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#e8f5e9');
  const retrievedAt = new Date().toISOString();
  const rows = prices.map((record) => [
    record.port,
    record.fuelType,
    record.price,
    record.currency,
    record.unit,
    record.region,
    record.source,
    record.timestamp,
    retrievedAt
  ]);
  sheet.getRange(2, 1, rows.length, 9).setValues(rows);
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0.00');
  sheet.autoResizeColumns(1, 9);
  sheet.activate();
}

/** @customfunction */
function BUNKER_PRICE(port, fuelType) {
  const normalizedPort = normalizeCode_(port, 'Port');
  const normalizedFuel = normalizeCode_(fuelType, 'Fuel type');
  const records = fetchBunkerRecords_(
    `?port=${encodeURIComponent(normalizedPort)}&fuel_type=${encodeURIComponent(normalizedFuel)}`
  );
  return records[0].price;
}

/** @customfunction */
function BUNKER_PORT_PRICES(port) {
  const normalizedPort = normalizeCode_(port, 'Port');
  const records = fetchBunkerRecords_(`?port=${encodeURIComponent(normalizedPort)}`);
  return [['Fuel Type', 'Price', 'Currency', 'Unit', 'Source Timestamp']].concat(
    records.map((record) => [record.fuelType, record.price, record.currency, record.unit, record.timestamp])
  );
}

function validateFutureContract_(record, subject) {
  const price = Number(record.price);
  if (!Number.isFinite(price)) throw new Error(`${subject} is missing a finite price.`);
  return {
    month: typeof record.month === 'string' ? record.month : '',
    price,
    change: Number.isFinite(Number(record.change)) ? Number(record.change) : ''
  };
}

/** @customfunction */
function FUTURES_PRICE(contract) {
  const code = normalizeCode_(contract, 'Contract code');
  const cacheKey = `futures_price_${code}`;
  const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.futures);
  if (cached !== null) return cached;
  const body = requestJson_(`/futures/latest?contract=${encodeURIComponent(code)}`, requireApiKey_());
  const value = validateFutureContract_(extractDataArray_(body, 'contracts', 'futures contracts')[0], 'Futures contract').price;
  putCachedValue_(cacheKey, value, CACHE_TTL_SECONDS.futures);
  return value;
}

/** @customfunction */
function FUTURES_CURVE(contract) {
  const code = normalizeCode_(contract, 'Contract code');
  const cacheKey = `futures_curve_${code}`;
  const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.futures);
  if (cached) return cached;
  const body = requestJson_(`/futures/curve?contract=${encodeURIComponent(code)}`, requireApiKey_());
  const contracts = extractDataArray_(body, 'contracts', 'futures contracts').map((record) => validateFutureContract_(record, 'Futures contract'));
  const result = [['Month', 'Price', 'Change']].concat(
    contracts.map((record) => [record.month, record.price, record.change])
  );
  putCachedValue_(cacheKey, result, CACHE_TTL_SECONDS.futures);
  return result;
}

function showFuturesInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Futures Data Functions',
    '=FUTURES_PRICE("BZ")\n=FUTURES_PRICE("CL")\n=FUTURES_CURVE("BZ")\n\nDataset access depends on your account entitlements. Review ' + PRICING_URL,
    ui.ButtonSet.OK
  );
}

function validateRigData_(data) {
  if (!data || typeof data !== 'object') throw new Error('Rig count response is missing data.');
  const oil = Number(data.oil);
  const gas = Number(data.gas);
  const total = Number(data.total);
  if (![oil, gas, total].every(Number.isFinite)) {
    throw new Error('Rig count response is missing numeric oil, gas, or total values.');
  }
  if (typeof data.date !== 'string' || !data.date || !Number.isFinite(new Date(data.date).getTime())) {
    throw new Error('Rig count response is missing a valid source date.');
  }
  return { oil, gas, total, date: data.date };
}

/** @customfunction */
function RIG_COUNT(type) {
  const selectedType = String(type || 'total').toLowerCase();
  const cacheKey = 'rig_count_data';
  let rigData = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.rigCount);
  if (!rigData) {
    const body = requestJson_('/rig-counts/latest', requireApiKey_());
    rigData = validateRigData_(body.data);
    putCachedValue_(cacheKey, rigData, CACHE_TTL_SECONDS.rigCount);
  }
  if (selectedType === 'oil') return rigData.oil;
  if (selectedType === 'gas') return rigData.gas;
  if (selectedType === 'total') return rigData.total;
  if (selectedType === 'all') {
    return [
      ['Type', 'Count'],
      ['Oil', rigData.oil],
      ['Gas', rigData.gas],
      ['Total', rigData.total],
      ['Source Date', rigData.date]
    ];
  }
  throw new Error('Rig count type must be oil, gas, total, or all.');
}

function showRigCountInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Rig Count Functions',
    '=RIG_COUNT("oil")\n=RIG_COUNT("gas")\n=RIG_COUNT("total")\n=RIG_COUNT("all")\n\nDataset access depends on your account entitlements.',
    ui.ButtonSet.OK
  );
}
