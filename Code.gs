/**
 * OilPriceAPI Google Sheets add-on.
 *
 * Public listing: https://workspace.google.com/marketplace/app/oilpriceapi_for_google_sheets/991152473434
 * Product facts: https://api.oilpriceapi.com/product-facts.json
 */

const API_BASE_URL = 'https://api.oilpriceapi.com/v1';
const ADDON_VERSION = '1.3.0';
const KEY_PROPERTY = 'OILPRICEAPI_KEY';
const LAST_DIAGNOSTIC_PROPERTY = 'OILPRICEAPI_LAST_DIAGNOSTIC';
const CACHE_GENERATION_PROPERTY = 'OILPRICEAPI_CACHE_GENERATION';
const MAX_BATCH_CODES = 25;
const PRICING_URL = 'https://www.oilpriceapi.com/pricing';
const SIGNUP_URL = 'https://www.oilpriceapi.com/auth/signup';
const MARKETPLACE_URL = 'https://workspace.google.com/marketplace/app/oilpriceapi_for_google_sheets/991152473434';
const CACHE_TTL_SECONDS = {
  latest: 300,
  latestFree: 3600,
  latestEnterprise: 60,
  generic: 300,
  catalog: 21600,
  bunker: 300,
  exchangeRates: 3600,
  history: 3600,
  futures: 300,
  rigCount: 3600
};
const MAX_CACHE_TTL_SECONDS = 21600;
let cachedGeneration_ = null;

// Keep generic worksheet requests aligned with the Excel add-in's reviewed
// endpoint catalog. Add endpoints deliberately after API-shape tests exist.
const ENDPOINT_CATALOG = [
  /^\/v1\/status$/,
  /^\/v1\/prices$/,
  /^\/v1\/prices\/latest$/,
  /^\/v1\/prices\/past_day$/,
  /^\/v1\/prices\/past_week$/,
  /^\/v1\/prices\/past_month$/,
  /^\/v1\/prices\/past_year$/,
  /^\/v1\/prices\/historical$/,
  /^\/v1\/prices\/all$/,
  /^\/v1\/prices\/all\/health$/,
  /^\/v1\/diesel-prices$/,
  /^\/v1\/futures\/(ice-brent|ice-wti|ice-gasoil|natural-gas|eua-carbon)(\/(historical|ohlc|intraday|spreads|curve|spread-history))?$/,
  /^\/v1\/commodities$/,
  /^\/v1\/commodities\/categories$/,
  /^\/v1\/commodities\/[A-Za-z0-9_.-]+$/
];

const SENSITIVE_QUERY_KEYS = new Set([
  'accesstoken',
  'apikey',
  'authorization',
  'auth',
  'bearer',
  'bearertoken',
  'clientsecret',
  'credential',
  'credentials',
  'key',
  'password',
  'secret',
  'token',
  'xapikey'
]);

const NUMERIC_FIELDS = new Set([
  'open',
  'high',
  'low',
  'close',
  'settlement',
  'settlement_price',
  'last_price',
  'price',
  'spread_value',
  'spread_percentage',
  'front_price',
  'back_price',
  'change_percent',
  'volume',
  'open_interest'
]);

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
  const ui = SpreadsheetApp.getUi();
  const menu = typeof ui.createAddonMenu === 'function'
    ? ui.createAddonMenu()
    : ui.createMenu('OilPriceAPI');
  menu
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

function onInstall(event) {
  onOpen(event);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('OilPriceAPI')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'OilPriceAPI for Google Sheets™',
    `Runtime version: ${ADDON_VERSION}\n\n` +
      'Source-timestamped energy price data. Dataset access and freshness vary.\n\n' +
      'Available in Google Workspace Marketplace. The listing runtime is managed separately during staged releases.\n\n' +
      `Install: ${MARKETPLACE_URL}\n\n` +
      'Website: https://www.oilpriceapi.com\n' +
      'Docs: https://docs.oilpriceapi.com',
    ui.ButtonSet.OK
  );
}

function getDocumentProperties_() {
  try {
    return PropertiesService.getDocumentProperties();
  } catch (error) {
    return null;
  }
}

function getActiveSpreadsheetId_() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet && typeof spreadsheet.getId === 'function'
      ? spreadsheet.getId()
      : null;
  } catch (error) {
    return null;
  }
}

function getSpreadsheetKeyProperty_() {
  const spreadsheetId = getActiveSpreadsheetId_();
  return spreadsheetId ? `${KEY_PROPERTY}:${spreadsheetId}` : null;
}

function getApiKey_() {
  const documentProperties = getDocumentProperties_();
  const documentKey = documentProperties
    ? documentProperties.getProperty(KEY_PROPERTY)
    : null;
  if (documentKey) return documentKey;

  // Custom functions run in a distinct authorization context where Google can
  // make document properties unavailable. User properties resolve to the
  // spreadsheet owner in that context, so key the compatibility copy by the
  // active spreadsheet ID to prevent credentials crossing between sheets.
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetKeyProperty = getSpreadsheetKeyProperty_();
  const spreadsheetKey = spreadsheetKeyProperty
    ? userProperties.getProperty(spreadsheetKeyProperty)
    : null;
  if (spreadsheetKey) return spreadsheetKey;

  // Migration fallback for keys saved by releases before Apps Script version 6.
  return userProperties.getProperty(KEY_PROPERTY);
}

function requireApiKey_() {
  const apiKey = getApiKey_();
  if (!apiKey) {
    throw makeError_('AUTH_REQUIRED', `Configure an API key from OilPriceAPI > Configure API Key, or create one at ${SIGNUP_URL}.`);
  }
  return apiKey;
}

function saveApiKey(apiKey) {
  if (typeof apiKey !== 'string' || !apiKey.trim()) {
    throw new Error('API key is required.');
  }
  const documentProperties = getDocumentProperties_();
  if (!documentProperties) {
    throw makeError_(
      'ADDON_CONTEXT_REQUIRED',
      'Open the OilPriceAPI sidebar from a Google Sheet before saving the API key.'
    );
  }
  const spreadsheetKeyProperty = getSpreadsheetKeyProperty_();
  if (!spreadsheetKeyProperty) {
    throw makeError_(
      'ADDON_CONTEXT_REQUIRED',
      'Open the OilPriceAPI sidebar from a Google Sheet before saving the API key.'
    );
  }
  documentProperties.setProperty(KEY_PROPERTY, apiKey.trim());
  const previousGeneration = Number(documentProperties.getProperty(CACHE_GENERATION_PROPERTY));
  const cacheGeneration = String(
    Math.max(Date.now(), Number.isFinite(previousGeneration) ? previousGeneration + 1 : 0)
  );
  documentProperties.setProperty(CACHE_GENERATION_PROPERTY, cacheGeneration);
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(spreadsheetKeyProperty, apiKey.trim());
  userProperties.setProperty(`${CACHE_GENERATION_PROPERTY}:${getActiveSpreadsheetId_()}`, cacheGeneration);
  userProperties.deleteProperty(KEY_PROPERTY);
  cachedGeneration_ = null;
  return {
    success: true,
    message: 'API key saved for this spreadsheet in Apps Script properties.'
  };
}

function deleteApiKey() {
  const documentProperties = getDocumentProperties_();
  if (documentProperties) {
    documentProperties.deleteProperty(KEY_PROPERTY);
    documentProperties.deleteProperty(LAST_DIAGNOSTIC_PROPERTY);
    documentProperties.deleteProperty(CACHE_GENERATION_PROPERTY);
  }
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetKeyProperty = getSpreadsheetKeyProperty_();
  if (spreadsheetKeyProperty) userProperties.deleteProperty(spreadsheetKeyProperty);
  const spreadsheetId = getActiveSpreadsheetId_();
  if (spreadsheetId) userProperties.deleteProperty(`${CACHE_GENERATION_PROPERTY}:${spreadsheetId}`);
  userProperties.deleteProperty(KEY_PROPERTY);
  userProperties.deleteProperty(LAST_DIAGNOSTIC_PROPERTY);
  cachedGeneration_ = null;
  return {
    success: true,
    message: 'Stored spreadsheet API key and request diagnostic deleted.'
  };
}

function getApiKeyStatus() {
  return { configured: Boolean(getApiKey_()) };
}

function normalizeCode_(value, label) {
  const code = String(value || '').trim().toUpperCase();
  if (!code) {
    throw makeError_('INVALID_CODE', `${label || 'Code'} is required.`);
  }
  if (!/^[A-Z0-9_:-]+$/.test(code)) {
    throw makeError_('INVALID_CODE', `${label || 'Code'} contains unsupported characters.`);
  }
  return code;
}

function makeError_(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function errorCode_(error) {
  if (error && typeof error.code === 'string' && error.code) return error.code;
  return 'ERROR';
}

function formulaError_(error) {
  const message = error && typeof error.message === 'string' ? error.message : 'Unexpected error.';
  return `#${errorCode_(error)}: ${message}`;
}

function formulaTableError_(error) {
  const message = error && typeof error.message === 'string' ? error.message : 'Unexpected error.';
  return [[`#${errorCode_(error)}`, message]];
}

function requestEndpoint_(path) {
  const withoutQuery = String(path || '').split('?', 1)[0];
  return withoutQuery.startsWith('/v1/') ? withoutQuery : `/v1${withoutQuery}`;
}

function responseHeader_(response, name) {
  if (!response || typeof response.getHeaders !== 'function') return '';
  const headers = response.getHeaders() || {};
  const target = String(name).toLowerCase();
  const key = Object.keys(headers).find((candidate) => candidate.toLowerCase() === target);
  return key && typeof headers[key] === 'string' ? headers[key].slice(0, 128) : '';
}

function cacheGeneration_() {
  if (cachedGeneration_ !== null) return cachedGeneration_;
  const documentProperties = getDocumentProperties_();
  const documentGeneration = documentProperties
    ? documentProperties.getProperty(CACHE_GENERATION_PROPERTY)
    : null;
  if (documentGeneration) {
    cachedGeneration_ = documentGeneration;
    return cachedGeneration_;
  }

  const spreadsheetId = getActiveSpreadsheetId_();
  if (!spreadsheetId) {
    cachedGeneration_ = 'legacy';
    return cachedGeneration_;
  }
  cachedGeneration_ = PropertiesService.getUserProperties().getProperty(
    `${CACHE_GENERATION_PROPERTY}:${spreadsheetId}`
  ) || 'legacy';
  return cachedGeneration_;
}

function stableCacheHash_(value) {
  let hash = 2166136261;
  const text = String(value || '');
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  const digest = (hash >>> 0).toString(36);
  const slug = text.replace(/[^A-Za-z0-9_-]+/g, '_').slice(0, 96);
  return `${slug}_${digest}`;
}

function requestBlockKeys_(path) {
  const endpoint = requestEndpoint_(path);
  return {
    global: 'request_block_global',
    endpoint: `request_block_endpoint_${stableCacheHash_(endpoint)}`,
    request: `request_block_request_${stableCacheHash_(String(path))}`
  };
}

function cachedRequestBlock_(path) {
  const keys = requestBlockKeys_(path);
  for (const key of [keys.global, keys.endpoint, keys.request]) {
    const blocked = getCachedValue_(key, MAX_CACHE_TTL_SECONDS, 'document');
    if (blocked && typeof blocked.code === 'string' && typeof blocked.message === 'string') {
      return blocked;
    }
  }
  return null;
}

function clearRequestBlocks_(path) {
  const keys = requestBlockKeys_(path);
  for (const key of [keys.global, keys.endpoint, keys.request]) {
    removeCachedValue_(key, 'document');
  }
}

function retryWindowSeconds_(response, fallbackSeconds) {
  const nowSeconds = Math.floor(Date.now() / 1000);
  const reset = Number(responseHeader_(response, 'X-RateLimit-Reset'));
  if (Number.isFinite(reset) && reset > nowSeconds) {
    return Math.min(MAX_CACHE_TTL_SECONDS, Math.max(1, Math.ceil(reset - nowSeconds)));
  }

  const retryAfter = responseHeader_(response, 'Retry-After');
  const numericRetry = Number(retryAfter);
  if (Number.isFinite(numericRetry) && numericRetry > 0) {
    const seconds = numericRetry > nowSeconds ? numericRetry - nowSeconds : numericRetry;
    return Math.min(MAX_CACHE_TTL_SECONDS, Math.max(1, Math.ceil(seconds)));
  }
  const retryDate = new Date(retryAfter).getTime();
  if (Number.isFinite(retryDate) && retryDate > Date.now()) {
    return Math.min(
      MAX_CACHE_TTL_SECONDS,
      Math.max(1, Math.ceil((retryDate - Date.now()) / 1000))
    );
  }
  return fallbackSeconds;
}

function responseMessage_(response) {
  try {
    const body = JSON.parse(response.getContentText());
    const data = body && (body.data || body);
    if (data && typeof data.message === 'string' && data.message.trim()) {
      return data.message.trim().slice(0, 320);
    }
  } catch (error) {
    // Use status-derived recovery text when the error body is unreadable.
  }
  return '';
}

function blockRequest_(path, statusCode, response, error) {
  const keys = requestBlockKeys_(path);
  let key = null;
  let ttlSeconds = 0;
  if (statusCode === 401) {
    key = keys.global;
    ttlSeconds = MAX_CACHE_TTL_SECONDS;
  } else if (statusCode === 402) {
    key = keys.global;
    ttlSeconds = retryWindowSeconds_(response, MAX_CACHE_TTL_SECONDS);
  } else if (statusCode === 429) {
    key = keys.global;
    ttlSeconds = retryWindowSeconds_(response, 60);
  } else if (statusCode === 403) {
    key = keys.endpoint;
    ttlSeconds = MAX_CACHE_TTL_SECONDS;
  } else if (statusCode === 404) {
    key = keys.request;
    ttlSeconds = 3600;
  }
  if (key && ttlSeconds > 0) {
    putCachedValue_(
      key,
      { code: error.code, message: error.message },
      ttlSeconds,
      'document'
    );
  }
}

function rememberResponseTier_(response) {
  const tier = responseHeader_(response, 'X-RateLimit-Tier').trim().toLowerCase();
  if (/^[a-z0-9_-]{1,32}$/.test(tier)) {
    putCachedValue_('account_tier', tier, MAX_CACHE_TTL_SECONDS, 'document');
  }
}

function latestCacheTtl_() {
  const tier = getCachedValue_('account_tier', MAX_CACHE_TTL_SECONDS, 'document');
  if (tier === 'free') return CACHE_TTL_SECONDS.latestFree;
  if (tier === 'enterprise') return CACHE_TTL_SECONDS.latestEnterprise;
  return CACHE_TTL_SECONDS.latest;
}

function persistDiagnostic_(input) {
  try {
    const diagnostic = {
      schemaVersion: 1,
      source: input.source || 'custom-function',
      result: input.result || 'error',
      code: input.code || 'ERROR',
      endpoint: requestEndpoint_(input.endpoint || ''),
      at: new Date().toISOString(),
      durationMs: Math.max(0, Math.round(Number(input.durationMs) || 0))
    };
    if (Number.isInteger(input.httpStatus)) diagnostic.httpStatus = input.httpStatus;
    if (typeof input.requestId === 'string' && input.requestId) {
      diagnostic.requestId = input.requestId.slice(0, 128);
    }
    const properties = getDocumentProperties_() || PropertiesService.getUserProperties();
    properties.setProperty(
      LAST_DIAGNOSTIC_PROPERTY,
      JSON.stringify(diagnostic)
    );
  } catch (error) {
    // Diagnostics must never break a worksheet function.
  }
}

function getLastDiagnostic() {
  const properties = getDocumentProperties_() || PropertiesService.getUserProperties();
  const raw = properties.getProperty(LAST_DIAGNOSTIC_PROPERTY);
  if (!raw) return null;
  try {
    const value = JSON.parse(raw);
    if (
      !value ||
      value.schemaVersion !== 1 ||
      typeof value.code !== 'string' ||
      typeof value.endpoint !== 'string' ||
      typeof value.at !== 'string'
    ) {
      return null;
    }
    return value;
  } catch (error) {
    return null;
  }
}

function requestJson_(path, apiKey, options) {
  const startedAt = Date.now();
  const endpoint = requestEndpoint_(path);
  const bypassBlock = options && options.bypassBlock === true;
  if (!bypassBlock) {
    const blocked = cachedRequestBlock_(path);
    if (blocked) throw makeError_(blocked.code, blocked.message);
  }
  let response;
  try {
    const relativePath = String(path).startsWith('/v1/') ? String(path).slice(3) : String(path);
    response = UrlFetchApp.fetch(`${API_BASE_URL}${relativePath}`, {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Accept': 'application/json',
        // Apps Script locks User-Agent, so the server's client classifier is fed
        // via X-API-Client instead (MinimalAnalyticsService.explicit_client_marker,
        // which already maps oilpriceapi-google-sheets -> client_type
        // 'sdk-google-sheets'). Without this, every call from this add-on lands as
        // client_type 'unknown' and the add-on is invisible in adoption reporting.
        // 285 users were sitting in 'unknown' when this was found. (#6167)
        'X-API-Client': 'oilpriceapi-google-sheets',
        'X-Client-Version': ADDON_VERSION
      },
      muteHttpExceptions: true
    });
  } catch (error) {
    persistDiagnostic_({
      result: 'timeout',
      code: 'TIMEOUT',
      endpoint,
      durationMs: Date.now() - startedAt
    });
    throw makeError_('TIMEOUT', 'OilPriceAPI request failed or timed out. Retry later.');
  }

  const statusCode = response.getResponseCode();
  const requestId = responseHeader_(response, 'x-request-id');
  rememberResponseTier_(response);
  if (statusCode === 401) {
    const apiError = makeError_('AUTH_INVALID', `Invalid or revoked API key. Replace it from ${SIGNUP_URL}.`);
    persistDiagnostic_({ result: 'http-error', code: 'AUTH_INVALID', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    blockRequest_(path, statusCode, response, apiError);
    throw apiError;
  }
  if (statusCode === 402) {
    const detail = responseMessage_(response);
    const apiError = makeError_(
      'UPGRADE_REQUIRED',
      `${detail ? `${detail} ` : ''}Review or upgrade access at ${PRICING_URL}.`
    );
    persistDiagnostic_({ result: 'http-error', code: 'UPGRADE_REQUIRED', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    blockRequest_(path, statusCode, response, apiError);
    throw apiError;
  }
  if (statusCode === 403) {
    const apiError = makeError_('UPGRADE_REQUIRED', `This account cannot access the requested dataset. Review ${PRICING_URL}.`);
    persistDiagnostic_({ result: 'http-error', code: 'UPGRADE_REQUIRED', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    blockRequest_(path, statusCode, response, apiError);
    throw apiError;
  }
  if (statusCode === 429) {
    const apiError = makeError_('RATE_LIMITED', 'OilPriceAPI rate or quota limit reached. Wait for the current limit window before retrying.');
    persistDiagnostic_({ result: 'http-error', code: 'RATE_LIMITED', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    blockRequest_(path, statusCode, response, apiError);
    throw apiError;
  }
  if (statusCode === 404) {
    const apiError = makeError_('NO_DATA', 'The requested OilPriceAPI resource was not found. Check the code and endpoint.');
    persistDiagnostic_({ result: 'http-error', code: 'NO_DATA', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    blockRequest_(path, statusCode, response, apiError);
    throw apiError;
  }
  if (statusCode === 400 || statusCode === 422) {
    let message = `OilPriceAPI rejected the request (HTTP ${statusCode}).`;
    try {
      const errorBody = JSON.parse(response.getContentText());
      const errorData = errorBody && (errorBody.data || errorBody);
      if (errorData && typeof errorData.message === 'string' && errorData.message.trim()) {
        message = errorData.message.trim();
      }
    } catch (error) {
      // Keep the status-derived message for an unreadable error body.
    }
    persistDiagnostic_({ result: 'http-error', code: 'INVALID_CODE', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    throw makeError_('INVALID_CODE', message);
  }
  if (statusCode >= 500) {
    persistDiagnostic_({ result: 'http-error', code: 'SERVER_ERROR', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    throw makeError_('SERVER_ERROR', `OilPriceAPI is temporarily unavailable (HTTP ${statusCode}). Retry later.`);
  }
  if (statusCode !== 200) {
    persistDiagnostic_({ result: 'http-error', code: 'ERROR', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    throw makeError_('ERROR', `OilPriceAPI request failed with HTTP ${statusCode}.`);
  }

  clearRequestBlocks_(path);

  let body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (error) {
    persistDiagnostic_({ result: 'invalid-response', code: 'INVALID_RESPONSE', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    throw makeError_('INVALID_RESPONSE', 'OilPriceAPI returned malformed JSON for a successful request.');
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    persistDiagnostic_({ result: 'invalid-response', code: 'INVALID_RESPONSE', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
    throw makeError_('INVALID_RESPONSE', 'OilPriceAPI returned a non-object successful response.');
  }
  persistDiagnostic_({ result: 'success', code: 'OK', endpoint, durationMs: Date.now() - startedAt, httpStatus: statusCode, requestId });
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
    throw makeError_('INVALID_RESPONSE', `OilPriceAPI returned no usable ${subject || 'price'} in a successful response.`);
  }
  return records;
}

function extractDataArray_(body, key, subject) {
  const payload = body.data;
  const records = payload && typeof payload === 'object' ? payload[key] : null;
  if (!Array.isArray(records) || records.length === 0 || records.some((item) => !item || typeof item !== 'object')) {
    throw makeError_('INVALID_RESPONSE', `OilPriceAPI returned no usable ${subject} in a successful response.`);
  }
  return records;
}

function sourceTimestamp_(record, subject) {
  const timestamp = record.as_of || record.created_at || record.updated_at || record.timestamp || record.date || record.collected_at;
  if (typeof timestamp !== 'string' || !timestamp.trim() || !Number.isFinite(new Date(timestamp).getTime())) {
    throw makeError_('INVALID_RESPONSE', `${subject} is missing a valid source timestamp.`);
  }
  return timestamp;
}

function validatePriceRecord_(record, subject) {
  const code = normalizeCode_(record.code || record.symbol, `${subject} code`);
  const price = Number(record.price);
  if (!Number.isFinite(price)) {
    throw makeError_('INVALID_RESPONSE', `${subject} ${code} is missing a finite price.`);
  }
  if (typeof record.currency !== 'string' || !record.currency.trim()) {
    throw makeError_('INVALID_RESPONSE', `${subject} ${code} is missing currency.`);
  }
  if (typeof record.unit !== 'string' || !record.unit.trim()) {
    throw makeError_('INVALID_RESPONSE', `${subject} ${code} is missing unit.`);
  }
  return {
    code,
    price,
    currency: record.currency.trim(),
    unit: record.unit.trim(),
    source: typeof record.source === 'string' ? record.source : '',
    sourceDescription:
      record.metadata && typeof record.metadata.source_description === 'string'
        ? record.metadata.source_description
        : '',
    timestamp: sourceTimestamp_(record, `${subject} ${code}`),
    collectedAt: typeof record.collected_at === 'string' ? record.collected_at : '',
    formatted: typeof record.formatted === 'string' ? record.formatted : '',
    dataStatus:
      typeof record.data_status === 'string'
        ? record.data_status
        : record.freshness && typeof record.freshness.status === 'string'
          ? record.freshness.status
          : '',
    stale: typeof record.stale === 'boolean' ? record.stale : '',
    ageDays: Number.isFinite(Number(record.age_days)) ? Number(record.age_days) : ''
  };
}

function cacheStore_(scope) {
  if (typeof CacheService === 'undefined') return null;
  if (scope === 'document' && typeof CacheService.getDocumentCache === 'function') {
    try {
      const documentCache = CacheService.getDocumentCache();
      if (documentCache) return documentCache;
    } catch (error) {
      // Fall through to the per-user cache when the document cache is unavailable.
    }
  }
  try {
    return typeof CacheService.getUserCache === 'function'
      ? CacheService.getUserCache()
      : null;
  } catch (error) {
    return null;
  }
}

function namespacedCacheKey_(cacheKey) {
  return `opa_${cacheGeneration_()}_${cacheKey}`;
}

function getCachedValue_(cacheKey, maxAgeSeconds, scope) {
  try {
    const cache = cacheStore_(scope || 'document');
    if (!cache || typeof cache.get !== 'function') return null;
    const namespacedKey = namespacedCacheKey_(cacheKey);
    const raw = cache.get(namespacedKey);
    if (!raw) return null;

    let envelope;
    try {
      envelope = JSON.parse(raw);
    } catch (error) {
      try {
        cache.remove(namespacedKey);
      } catch (removeError) {
        // An invalid entry can expire naturally if removal is unavailable.
      }
      return null;
    }
    if (
      !envelope ||
      typeof envelope.cachedAt !== 'number' ||
      !Object.prototype.hasOwnProperty.call(envelope, 'value') ||
      Date.now() - envelope.cachedAt > maxAgeSeconds * 1000
    ) {
      try {
        cache.remove(namespacedKey);
      } catch (removeError) {
        // An expired entry can expire naturally if removal is unavailable.
      }
      return null;
    }
    return envelope.value;
  } catch (error) {
    // Cache failures must degrade to a live request.
    return null;
  }
}

function putCachedValue_(cacheKey, value, ttlSeconds, scope) {
  try {
    cacheStore_(scope || 'document').put(
      namespacedCacheKey_(cacheKey),
      JSON.stringify({ cachedAt: Date.now(), value }),
      Math.min(MAX_CACHE_TTL_SECONDS, Math.max(1, Math.round(ttlSeconds)))
    );
  } catch (error) {
    // Cache limits or transient cache failures must not replace live API data.
  }
}

function removeCachedValue_(cacheKey, scope) {
  try {
    const cache = cacheStore_(scope || 'document');
    if (cache && typeof cache.remove === 'function') {
      cache.remove(namespacedCacheKey_(cacheKey));
    }
  } catch (error) {
    // Request-block cleanup must not replace a valid live response.
  }
}

function cacheMissLock_() {
  if (
    typeof LockService === 'undefined' ||
    typeof LockService.getDocumentLock !== 'function'
  ) return null;
  try {
    return LockService.getDocumentLock();
  } catch (error) {
    return null;
  }
}

function withCacheMissLock_(cacheKey, maxAgeSeconds, loader) {
  const lock = cacheMissLock_();
  if (!lock) return loader();
  let acquired;
  try {
    acquired = lock.tryLock(5000);
  } catch (error) {
    return loader();
  }
  if (!acquired) {
    const afterWait = getCachedValue_(cacheKey, maxAgeSeconds, 'document');
    if (afterWait !== null) return afterWait;
    throw makeError_(
      'RETRY_LATER',
      'Another sheet calculation is refreshing this value. Recalculate shortly.'
    );
  }
  try {
    const afterLock = getCachedValue_(cacheKey, maxAgeSeconds, 'document');
    return afterLock !== null ? afterLock : loader();
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      // Releasing a transient service handle must not replace live data.
    }
  }
}

function cachedRequestJson_(cacheKey, path, ttlSeconds) {
  const cached = getCachedValue_(cacheKey, ttlSeconds, 'document');
  if (cached !== null) return cached;
  return withCacheMissLock_(cacheKey, ttlSeconds, () => {
    const body = requestJson_(path, requireApiKey_());
    putCachedValue_(cacheKey, body, ttlSeconds, 'document');
    return body;
  });
}

function getLatestRecord_(commodityCode) {
  const code = normalizeCode_(commodityCode, 'Commodity code');
  const cacheKey = `latest_${code}`;
  const ttlSeconds = latestCacheTtl_();
  const cached = getCachedValue_(cacheKey, ttlSeconds, 'document');
  if (cached) return cached;

  return withCacheMissLock_(cacheKey, ttlSeconds, () => {
    const body = requestJson_(`/prices/latest?by_code=${encodeURIComponent(code)}`, requireApiKey_());
    let record;
    try {
      record = validatePriceRecord_(extractPriceRecords_(body, 'price')[0], 'Price record');
    } catch (error) {
      persistDiagnostic_({
        result: 'invalid-response',
        code: 'INVALID_RESPONSE',
        endpoint: '/v1/prices/latest',
        durationMs: 0,
        httpStatus: 200
      });
      throw error;
    }
    putCachedValue_(cacheKey, record, latestCacheTtl_(), 'document');
    return record;
  });
}

function normalizeCodeRange_(values) {
  const rows = Array.isArray(values) ? values : [values];
  const flattened = [];
  for (const row of rows) {
    if (Array.isArray(row)) flattened.push(...row);
    else flattened.push(row);
  }
  const codes = [];
  for (const value of flattened) {
    if (value === null || value === undefined || String(value).trim() === '') continue;
    const code = normalizeCode_(value, 'Commodity code');
    if (!codes.includes(code)) codes.push(code);
  }
  if (codes.length === 0) {
    throw makeError_('INVALID_CODE', 'Select at least one commodity code.');
  }
  if (codes.length > MAX_BATCH_CODES) {
    throw makeError_('INVALID_CODE', `Select at most ${MAX_BATCH_CODES} commodity codes.`);
  }
  return codes;
}

function readLatestRecordsFromCache_(codes, ttlSeconds) {
  const records = new Map();
  for (const code of codes) {
    const record = getCachedValue_(`latest_${code}`, ttlSeconds, 'document');
    if (record) records.set(code, record);
  }
  return records;
}

function getLatestRecords_(codes) {
  let ttlSeconds = latestCacheTtl_();
  let records = readLatestRecordsFromCache_(codes, ttlSeconds);
  if (records.size === codes.length) return codes.map((code) => records.get(code));

  let lock = cacheMissLock_();
  let acquired = true;
  if (lock) {
    try {
      acquired = lock.tryLock(5000);
    } catch (error) {
      lock = null;
    }
  }
  if (!acquired) {
    records = readLatestRecordsFromCache_(codes, latestCacheTtl_());
    if (records.size === codes.length) return codes.map((code) => records.get(code));
    throw makeError_(
      'RETRY_LATER',
      'Another sheet calculation is refreshing these values. Recalculate shortly.'
    );
  }

  try {
    ttlSeconds = latestCacheTtl_();
    records = readLatestRecordsFromCache_(codes, ttlSeconds);
    const missingCodes = codes.filter((code) => !records.has(code));
    if (missingCodes.length > 0) {
      const body = requestJson_(
        `/prices/latest?by_code=${encodeURIComponent(missingCodes.join(','))}`,
        requireApiKey_()
      );
      const missingSet = new Set(missingCodes);
      const recordErrors = new Map();
      for (const rawRecord of extractPriceRecords_(body, 'prices')) {
        let rawCode;
        try {
          rawCode = normalizeCode_(rawRecord && (rawRecord.code || rawRecord.symbol), 'Price record code');
        } catch (error) {
          continue;
        }
        if (!missingSet.has(rawCode)) continue;
        let record;
        try {
          record = validatePriceRecord_(rawRecord, 'Price record');
        } catch (error) {
          recordErrors.set(rawCode, error);
          continue;
        }
        if (!missingSet.has(record.code)) continue;
        records.set(record.code, record);
        putCachedValue_(`latest_${record.code}`, record, latestCacheTtl_(), 'document');
      }
      for (const code of missingCodes) {
        if (!records.has(code) && !recordErrors.has(code)) {
          recordErrors.set(code, makeError_('NO_DATA', `No latest value was returned for ${code}.`));
        }
      }
      return codes.map((code) => records.get(code) || { code, error: recordErrors.get(code) });
    }
    return codes.map((code) => records.get(code));
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (error) {
        // Releasing a transient service handle must not replace live data.
      }
    }
  }
}

function testConnection() {
  if (!getApiKey_()) {
    return { success: false, message: `Configure an API key first at ${SIGNUP_URL}.` };
  }
  try {
    const body = requestJson_(
      '/prices/latest?by_code=BRENT_CRUDE_USD',
      getApiKey_(),
      { bypassBlock: true }
    );
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

function normalizeApiPath_(path) {
  const value = String(path || '').trim();
  if (
    !value.startsWith('/v1/') ||
    value.includes('?') ||
    value.includes('#') ||
    value.length > 200 ||
    !ENDPOINT_CATALOG.some((pattern) => pattern.test(value))
  ) {
    throw makeError_('UNSUPPORTED_ENDPOINT', 'Use a supported OilPriceAPI GET endpoint.');
  }
  return value;
}

function normalizeQueryKey_(key) {
  return String(key || '').trim().toLowerCase().replace(/[-_.]/g, '');
}

function queryKeyParts_(key) {
  const parts = [String(key).split('[', 1)[0]];
  const pattern = /\[([^\]]*)\]/g;
  let match = pattern.exec(String(key));
  while (match) {
    if (match[1]) parts.push(match[1]);
    match = pattern.exec(String(key));
  }
  return parts.filter((part) => part.trim());
}

function normalizeApiQuery_(query) {
  const cleaned = String(query || '').trim().replace(/^\?/, '');
  if (!cleaned) return '';
  if (cleaned.length > 2000 || /[\r\n#]/.test(cleaned)) {
    throw makeError_('UNSUPPORTED_QUERY', 'Use a valid OilPriceAPI query string.');
  }

  for (const pair of cleaned.split('&')) {
    if (!pair) continue;
    const rawKey = pair.split('=', 1)[0];
    let key;
    try {
      key = decodeURIComponent(rawKey.replace(/\+/g, ' '));
    } catch (error) {
      throw makeError_('UNSUPPORTED_QUERY', 'Use a valid OilPriceAPI query string.');
    }
    const sensitive = queryKeyParts_(key).some((part) =>
      SENSITIVE_QUERY_KEYS.has(normalizeQueryKey_(part))
    );
    if (sensitive) {
      throw makeError_(
        'UNSUPPORTED_QUERY',
        'Do not pass API keys or credentials in query strings.'
      );
    }
  }
  return cleaned;
}

function valueToCell_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') {
    return value;
  }
  return JSON.stringify(value);
}

function cellFor_(field, value) {
  if (typeof value === 'string' && NUMERIC_FIELDS.has(field)) {
    const trimmed = value.trim();
    const numeric = Number(trimmed);
    if (trimmed && Number.isFinite(numeric)) return numeric;
  }
  return valueToCell_(value);
}

function objectToTable_(value) {
  return [['Field', 'Value']].concat(
    Object.keys(value).map((key) => [key, cellFor_(key, value[key])])
  );
}

function arrayToTable_(data) {
  if (data.length === 0) return [['#NO_DATA', 'No data returned.']];
  if (!data[0] || typeof data[0] !== 'object' || Array.isArray(data[0])) {
    return [['Value']].concat(data.map((entry) => [valueToCell_(entry)]));
  }
  const headers = Object.keys(data[0]);
  return [headers].concat(
    data.map((entry) => headers.map((header) => cellFor_(header, entry[header])))
  );
}

function pricesHashToTable_(prices) {
  const codes = Object.keys(prices);
  if (codes.length === 0) return [['#NO_DATA', 'No data returned.']];
  const allObjects = codes.every((code) =>
    prices[code] && typeof prices[code] === 'object' && !Array.isArray(prices[code])
  );
  if (!allObjects) {
    return [['Code', 'Value']].concat(
      codes.map((code) => [code, valueToCell_(prices[code])])
    );
  }
  const fields = [];
  codes.forEach((code) => {
    Object.keys(prices[code]).forEach((field) => {
      if (field !== 'code' && !fields.includes(field)) fields.push(field);
    });
  });
  return [['Code'].concat(fields)].concat(
    codes.map((code) => [code].concat(fields.map((field) => cellFor_(field, prices[code][field]))))
  );
}

function flattenNestedContracts_(contracts, childKey, parentFields) {
  const rows = [];
  for (const contract of contracts) {
    const children = contract && contract[childKey];
    if (!Array.isArray(children)) return null;
    const parent = {};
    parentFields.forEach((field) => {
      parent[field] = contract[field];
    });
    children.forEach((child) => rows.push(Object.assign({}, parent, child)));
  }
  return rows.length ? arrayToTable_(rows) : [['#NO_DATA', 'No data returned.']];
}

function futuresToTable_(payload) {
  if (Array.isArray(payload.spread_data)) {
    return arrayToTable_(
      payload.spread_data.map((entry) => ({
        trading_date: entry && entry.trading_date,
        front_month: entry && entry.front_contract && entry.front_contract.contract_month,
        front_price: entry && entry.front_contract && entry.front_contract.price,
        back_month: entry && entry.back_contract && entry.back_contract.contract_month,
        back_price: entry && entry.back_contract && entry.back_contract.price,
        spread_value: entry && entry.spread_value,
        spread_percentage: entry && entry.spread_percentage
      }))
    );
  }
  if (Array.isArray(payload.spreads)) {
    return flattenNestedContracts_(payload.spreads, 'daily_data', [
      'front_contract',
      'back_contract',
      'spread_type'
    ]) || [['#NO_DATA', 'No data returned.']];
  }
  if (!Array.isArray(payload.contracts) || payload.contracts.length === 0) {
    return [['#NO_DATA', 'No data returned.']];
  }
  if (Array.isArray(payload.contracts[0].daily_data)) {
    return flattenNestedContracts_(payload.contracts, 'daily_data', ['contract_month']);
  }
  if (Array.isArray(payload.contracts[0].price_data)) {
    return flattenNestedContracts_(payload.contracts, 'price_data', [
      'contract_month',
      'contract_code'
    ]);
  }
  return arrayToTable_(payload.contracts);
}

function responseToTable_(payload) {
  if (
    payload &&
    typeof payload === 'object' &&
    (Array.isArray(payload.contracts) ||
      Array.isArray(payload.spreads) ||
      Array.isArray(payload.spread_data))
  ) {
    return futuresToTable_(payload);
  }

  let data = payload && payload.data;
  if (
    data &&
    typeof data === 'object' &&
    !Array.isArray(data) &&
    data.data &&
    typeof data.data === 'object'
  ) {
    data = data.data;
  }
  if (Array.isArray(data)) return arrayToTable_(data);
  if (data && typeof data === 'object') {
    if (
      data.regional_average &&
      typeof data.regional_average === 'object' &&
      !Array.isArray(data.regional_average)
    ) {
      return objectToTable_(data.regional_average);
    }
    if (
      data.summary &&
      typeof data.summary === 'object' &&
      !Array.isArray(data.summary) &&
      !data.prices
    ) {
      return objectToTable_(data.summary);
    }
    if (Array.isArray(data.prices)) return arrayToTable_(data.prices);
    if (data.prices && typeof data.prices === 'object') {
      return pricesHashToTable_(data.prices);
    }
    if (Array.isArray(data.commodities)) {
      if (!data.commodities.length) return [['#NO_DATA', 'No data returned.']];
      return [['Code', 'Name', 'Category']].concat(
        data.commodities.map((commodity) => [
          String(commodity.code || ''),
          String(commodity.name || commodity.code || ''),
          String(commodity.category || '')
        ])
      );
    }
    return objectToTable_(data);
  }
  return [['#NO_DATA', 'No data returned.']];
}

function appendTruncationNote_(path, payload, table) {
  if (!/^\/v1\/prices\/(past_week|past_month)$/.test(path)) return table;
  const prices = payload && payload.data && payload.data.prices;
  if (!Array.isArray(prices) || prices.length < 100) return table;
  const width = table[0] ? table[0].length : 1;
  return table.concat([
    [
      'TRUNCATED: this endpoint returned the 100 most recent ticks. Use /v1/prices/historical with start_date and end_date for the full range.'
    ].concat(Array(Math.max(0, width - 1)).fill(''))
  ]);
}

/**
 * Latest available numeric value for a commodity code.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {number} API-provided numeric price.
 * @customfunction
 */
function OILPRICE(commodityCode) {
  try {
    return getLatestRecord_(commodityCode).price;
  } catch (error) {
    return formulaError_(error);
  }
}

/**
 * Fetches up to 25 commodity codes from a range in one request and spills a table.
 * @param {Array} commodityCodes One-column range of OilPriceAPI commodity codes.
 * @return {Array} Source-aware latest-price rows or a worksheet-readable error.
 * @customfunction
 */
function OILPRICE_TABLE(commodityCodes) {
  try {
    const codes = normalizeCodeRange_(commodityCodes);
    const records = getLatestRecords_(codes);
    return [['Code', 'Price', 'Currency', 'Unit', 'Source', 'Source Timestamp']].concat(
      records.map((record) => record.error
        ? [record.code, formulaError_(record.error), '', '', '', '']
        : [
          record.code,
          record.price,
          record.currency,
          record.unit,
          record.source,
          record.timestamp
        ])
    );
  } catch (error) {
    return formulaTableError_(error);
  }
}

/**
 * Excel-equivalent latest numeric price.
 * Google Sheets custom function names cannot contain a dot, so
 * OILPRICE_PRICE is the Sheets equivalent of Excel's OILPRICE.PRICE.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {number|string} Numeric price or worksheet-readable error.
 * @customfunction
 */
function OILPRICE_PRICE(commodityCode) {
  try {
    return getLatestRecord_(commodityCode).price;
  } catch (error) {
    return formulaError_(error);
  }
}

/**
 * Calls an allowlisted OilPriceAPI GET endpoint and spills a table.
 * @param {string} path Versioned API path such as /v1/prices/latest.
 * @param {string} query Query string such as by_code=WTI_USD.
 * @return {Array} API data rendered as a worksheet table.
 * @customfunction
 */
function OILPRICE_GET(path, query) {
  try {
    const normalizedPath = normalizeApiPath_(path);
    const normalizedQuery = normalizeApiQuery_(query);
    const requestPath = `${normalizedPath}${normalizedQuery ? `?${normalizedQuery}` : ''}`;
    const ttlSeconds = normalizedPath === '/v1/commodities'
      ? CACHE_TTL_SECONDS.catalog
      : CACHE_TTL_SECONDS.generic;
    const payload = cachedRequestJson_(
      `get_${stableCacheHash_(requestPath)}`,
      requestPath,
      ttlSeconds
    );
    return appendTruncationNote_(normalizedPath, payload, responseToTable_(payload));
  } catch (error) {
    return formulaTableError_(error);
  }
}

/**
 * Returns commodity codes available to the configured API key.
 * @return {Array} Code, name, and category rows.
 * @customfunction
 */
function OILPRICE_CODES() {
  return OILPRICE_GET('/v1/commodities', '');
}

/**
 * Reports API-provided freshness status for a latest record.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {string} API freshness state or worksheet-readable error.
 * @customfunction
 */
function OILPRICE_STATUS(commodityCode) {
  try {
    const record = getLatestRecord_(commodityCode);
    if (record.dataStatus) return record.dataStatus;
    if (typeof record.stale === 'boolean') return record.stale ? 'stale' : 'current';
    throw makeError_('NO_DATA', 'No freshness status returned.');
  } catch (error) {
    return formulaError_(error);
  }
}

/**
 * Returns currency/unit for a latest record.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {string} Currency and unit or worksheet-readable error.
 * @customfunction
 */
function OILPRICE_UNIT(commodityCode) {
  try {
    const record = getLatestRecord_(commodityCode);
    return `${record.currency}/${record.unit}`;
  } catch (error) {
    return formulaError_(error);
  }
}

/**
 * Spills source, timestamp, unit, and freshness context for a latest record.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @return {Array} Field/value rows.
 * @customfunction
 */
function OILPRICE_INFO(commodityCode) {
  try {
    const record = getLatestRecord_(commodityCode);
    return [
      ['Field', 'Value'],
      ['code', record.code],
      ['price', record.price],
      ['currency', record.currency],
      ['unit', record.unit],
      ['formatted', record.formatted],
      ['source', record.source],
      ['source_description', record.sourceDescription],
      ['source_timestamp', record.timestamp],
      ['collected_at', record.collectedAt],
      ['data_status', record.dataStatus],
      ['stale', record.stale],
      ['age_days', record.ageDays]
    ];
  } catch (error) {
    return formulaTableError_(error);
  }
}

/**
 * Historical source timestamp and price pairs.
 * @param {string} commodityCode OilPriceAPI commodity code.
 * @param {number} days Requested lookback selector from 1 through 365.
 * @return {Array} Rows of [source timestamp, price].
 * @customfunction
 */
function OILPRICE_HISTORY(commodityCode, days) {
  try {
    const code = normalizeCode_(commodityCode, 'Commodity code');
    const requestedDays = days === undefined || days === null || days === '' ? 30 : Number(days);
    if (!Number.isInteger(requestedDays) || requestedDays < 1 || requestedDays > 365) {
      throw makeError_('INVALID_INPUT', 'History days must be an integer from 1 through 365.');
    }
    let endpoint = 'past_year';
    if (requestedDays <= 1) endpoint = 'past_day';
    else if (requestedDays <= 7) endpoint = 'past_week';
    else if (requestedDays <= 30) endpoint = 'past_month';

    const cacheKey = `history_${code}_${endpoint}`;
    const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.history, 'document');
    if (cached) return cached;

    return withCacheMissLock_(cacheKey, CACHE_TTL_SECONDS.history, () => {
      const body = requestJson_(
        `/prices/${endpoint}?by_code=${encodeURIComponent(code)}`,
        requireApiKey_()
      );
      const records = extractPriceRecords_(body, 'historical price');
      const result = records.map((record) => {
        const normalized = validatePriceRecord_(record, 'Historical price record');
        return [normalized.timestamp, normalized.price];
      });
      putCachedValue_(cacheKey, result, CACHE_TTL_SECONDS.history, 'document');
      return result;
    });
  } catch (error) {
    return formulaTableError_(error);
  }
}

function fetchExchangeRates() {
  const body = cachedRequestJson_(
    'exchange_rates_gbp_eur_usd',
    '/prices/latest?by_code=GBP_USD,EUR_USD',
    CACHE_TTL_SECONDS.exchangeRates
  );
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
  try {
    const code = normalizeCode_(commodityCode, 'Commodity code');
    const commodityInfo = COMMODITY_MAP[code];
    if (!commodityInfo) {
      throw makeError_('INVALID_CODE', 'This code has no reference heat-content conversion mapping.');
    }
    const record = getLatestRecord_(code);
    const rates = record.currency === 'USD' ? null : fetchExchangeRates();
    return toUsd_(record.price, record.currency, rates) / heatContent_(commodityInfo);
  } catch (error) {
    return formulaError_(error);
  }
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
  const path = `/prices/data-connector${query || ''}`;
  const body = cachedRequestJson_(
    `bunker_${stableCacheHash_(path)}`,
    path,
    CACHE_TTL_SECONDS.bunker
  );
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
  try {
    const normalizedPort = normalizeCode_(port, 'Port');
    const normalizedFuel = normalizeCode_(fuelType, 'Fuel type');
    const records = fetchBunkerRecords_(
      `?port=${encodeURIComponent(normalizedPort)}&fuel_type=${encodeURIComponent(normalizedFuel)}`
    );
    return records[0].price;
  } catch (error) {
    return formulaError_(error);
  }
}

/** @customfunction */
function BUNKER_PORT_PRICES(port) {
  try {
    const normalizedPort = normalizeCode_(port, 'Port');
    const records = fetchBunkerRecords_(`?port=${encodeURIComponent(normalizedPort)}`);
    return [['Fuel Type', 'Price', 'Currency', 'Unit', 'Source Timestamp']].concat(
      records.map((record) => [record.fuelType, record.price, record.currency, record.unit, record.timestamp])
    );
  } catch (error) {
    return formulaTableError_(error);
  }
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
  try {
    const code = normalizeCode_(contract, 'Contract code');
    const cacheKey = `futures_price_${code}`;
    const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.futures, 'document');
    if (cached !== null) return cached;
    return withCacheMissLock_(cacheKey, CACHE_TTL_SECONDS.futures, () => {
      const body = requestJson_(`/futures/latest?contract=${encodeURIComponent(code)}`, requireApiKey_());
      const value = validateFutureContract_(extractDataArray_(body, 'contracts', 'futures contracts')[0], 'Futures contract').price;
      putCachedValue_(cacheKey, value, CACHE_TTL_SECONDS.futures, 'document');
      return value;
    });
  } catch (error) {
    return formulaError_(error);
  }
}

/** @customfunction */
function FUTURES_CURVE(contract) {
  try {
    const code = normalizeCode_(contract, 'Contract code');
    const cacheKey = `futures_curve_${code}`;
    const cached = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.futures, 'document');
    if (cached) return cached;
    return withCacheMissLock_(cacheKey, CACHE_TTL_SECONDS.futures, () => {
      const body = requestJson_(`/futures/curve?contract=${encodeURIComponent(code)}`, requireApiKey_());
      const contracts = extractDataArray_(body, 'contracts', 'futures contracts').map((record) => validateFutureContract_(record, 'Futures contract'));
      const result = [['Month', 'Price', 'Change']].concat(
        contracts.map((record) => [record.month, record.price, record.change])
      );
      putCachedValue_(cacheKey, result, CACHE_TTL_SECONDS.futures, 'document');
      return result;
    });
  } catch (error) {
    return formulaTableError_(error);
  }
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
  try {
    const selectedType = String(type || 'total').toLowerCase();
    const cacheKey = 'rig_count_data';
    let rigData = getCachedValue_(cacheKey, CACHE_TTL_SECONDS.rigCount, 'document');
    if (!rigData) {
      rigData = withCacheMissLock_(cacheKey, CACHE_TTL_SECONDS.rigCount, () => {
        const body = requestJson_('/rig-counts/latest', requireApiKey_());
        const validated = validateRigData_(body.data);
        putCachedValue_(cacheKey, validated, CACHE_TTL_SECONDS.rigCount, 'document');
        return validated;
      });
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
    throw makeError_('INVALID_INPUT', 'Rig count type must be oil, gas, total, or all.');
  } catch (error) {
    return String(type || 'total').toLowerCase() === 'all'
      ? formulaTableError_(error)
      : formulaError_(error);
  }
}

function showRigCountInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Rig Count Functions',
    '=RIG_COUNT("oil")\n=RIG_COUNT("gas")\n=RIG_COUNT("total")\n=RIG_COUNT("all")\n\nDataset access depends on your account entitlements.',
    ui.ButtonSet.OK
  );
}
