const OPA_PRODUCT = Object.freeze({
  "id": "gas-spread-monitor",
  "name": "Gas Spread Monitor by OilPriceAPI",
  "menu": "Gas Spread Monitor",
  "builder": "buildGasSpreadWorkbook",
  "tagline": "Normalize Henry Hub, TTF, and JKM to compare global gas spreads.",
  "landingPath": "/integrations/gas-spread-monitor",
  "signupCampaign": "gas_spread_monitor",
  "activationHeader": "gas-spread-monitor",
  "allowedCodes": [
    "NATURAL_GAS_USD",
    "DUTCH_TTF_EUR",
    "JKM_LNG_USD",
    "EUR_USD"
  ],
  "sheets": [
    "Gas Spread Monitor",
    "Market Data",
    "Conversion Audit"
  ],
  "workflow": "Creates a cross-market gas monitor that normalizes currencies and energy units before calculating TTF-Henry Hub and JKM-Henry Hub spreads.",
  "category": "Accounting and Finance"
});

/**
 * Shared security and workbook runtime for the OilPriceAPI Workspace portfolio.
 * Product behavior is supplied by exactly one Product.gs file at build time.
 */

const OPA_API_BASE_URL = 'https://api.oilpriceapi.com/v1';
const OPA_KEY_PROPERTY = 'OILPRICEAPI_KEY';
const OPA_ACTIVATED_PROPERTY = 'OILPRICEAPI_ACTIVATED_AT';
const OPA_VERSION = '0.1.0';
const OPA_SIGNUP_URL = 'https://www.oilpriceapi.com/auth/signup';
const OPA_ALLOWED_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.currentonly',
  'https://www.googleapis.com/auth/script.external_request',
  'https://www.googleapis.com/auth/script.container.ui'
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = typeof ui.createAddonMenu === 'function'
    ? ui.createAddonMenu()
    : ui.createMenu(OPA_PRODUCT.menu);
  menu
    .addItem(`Build ${OPA_PRODUCT.menu}`, OPA_PRODUCT.builder)
    .addItem('Configure OilPriceAPI key', 'showSidebar')
    .addSeparator()
    .addItem('About and data use', 'showAbout')
    .addToUi();
}

function onInstall(event) {
  onOpen(event);
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar');
  html.productName = OPA_PRODUCT.name;
  html.tagline = OPA_PRODUCT.tagline;
  html.builder = OPA_PRODUCT.builder;
  html.signupUrl = signupUrl_();
  html.landingUrl = landingUrl_();
  SpreadsheetApp.getUi().showSidebar(
    html.evaluate().setTitle(OPA_PRODUCT.menu).setWidth(340)
  );
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    OPA_PRODUCT.name,
    `${OPA_PRODUCT.workflow}\n\n` +
      `Version ${OPA_VERSION}. Dataset access and freshness depend on the configured OilPriceAPI account.\n\n` +
      'The add-on sends the requested market identifiers to OilPriceAPI and identifies this product in an HTTP header. ' +
      'It does not send spreadsheet contents, formulas, or cell values for analytics.\n\n' +
      'Google Sheets™ is a trademark of Google LLC.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function documentProperties_() {
  try {
    return PropertiesService.getDocumentProperties();
  } catch (error) {
    return null;
  }
}

function getApiKey_() {
  const properties = documentProperties_();
  return properties ? properties.getProperty(OPA_KEY_PROPERTY) : null;
}

function requireApiKey_() {
  const key = getApiKey_();
  if (!key) {
    throw new Error(`Configure an OilPriceAPI key first, or create one at ${signupUrl_()}.`);
  }
  return key;
}

function saveApiKey(apiKey) {
  const value = typeof apiKey === 'string' ? apiKey.trim() : '';
  if (!value || value.length > 512) throw new Error('Enter a valid OilPriceAPI key.');
  const properties = documentProperties_();
  if (!properties) throw new Error('Open the add-on from a spreadsheet before saving a key.');
  properties.setProperty(OPA_KEY_PROPERTY, value);
  return { success: true, configured: true };
}

function deleteApiKey() {
  const properties = documentProperties_();
  if (properties) {
    properties.deleteProperty(OPA_KEY_PROPERTY);
    properties.deleteProperty(OPA_ACTIVATED_PROPERTY);
  }
  return { success: true, configured: false };
}

function getSidebarState() {
  const properties = documentProperties_();
  return {
    product: OPA_PRODUCT.name,
    configured: Boolean(getApiKey_()),
    activated: Boolean(properties && properties.getProperty(OPA_ACTIVATED_PROPERTY)),
    landingUrl: landingUrl_(),
    signupUrl: signupUrl_()
  };
}

function signupUrl_() {
  return `${OPA_SIGNUP_URL}?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=${encodeURIComponent(OPA_PRODUCT.signupCampaign)}`;
}

function landingUrl_() {
  return `https://www.oilpriceapi.com${OPA_PRODUCT.landingPath}?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=${encodeURIComponent(OPA_PRODUCT.signupCampaign)}`;
}

function normalizeCode_(value) {
  const code = String(value || '').trim().toUpperCase();
  if (!/^[A-Z0-9_:-]+$/.test(code)) throw new Error('Unsupported market identifier.');
  if (OPA_PRODUCT.allowedCodes.indexOf(code) === -1) {
    throw new Error(`${code} is outside this product's reviewed market catalog.`);
  }
  return code;
}

function requestJson_(path, apiKey) {
  const normalizedPath = String(path || '');
  if (!normalizedPath.startsWith('/') || normalizedPath.includes('://') || normalizedPath.includes('..')) {
    throw new Error('Unsupported OilPriceAPI path.');
  }
  const response = UrlFetchApp.fetch(`${OPA_API_BASE_URL}${normalizedPath}`, {
    method: 'get',
    headers: {
      Authorization: `Token ${apiKey}`,
      Accept: 'application/json',
      'X-OilPriceAPI-Client': `${OPA_PRODUCT.activationHeader}/${OPA_VERSION}`
    },
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  if (status === 401) throw new Error('The OilPriceAPI key is invalid or revoked. Replace it in the sidebar.');
  if (status === 402 || status === 403) throw new Error('This dataset is not enabled for the account. Review OilPriceAPI pricing or use an entitled key.');
  if (status === 429) throw new Error('The OilPriceAPI rate or quota limit was reached. Retry later or review the account limit.');
  if (status < 200 || status >= 300) throw new Error(`OilPriceAPI returned HTTP ${status}. Retry later.`);
  let payload;
  try {
    payload = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error('OilPriceAPI returned an unreadable response.');
  }
  if (!payload || typeof payload !== 'object') throw new Error('OilPriceAPI returned an empty response.');
  return payload;
}

function priceRecords_(payload) {
  const data = payload && payload.data;
  let records = [];
  if (Array.isArray(data)) records = data;
  else if (data && Array.isArray(data.prices)) records = data.prices;
  else if (data && data.prices && typeof data.prices === 'object') records = Object.values(data.prices);
  else if (data && typeof data === 'object' && ('price' in data || 'code' in data)) records = [data];
  const normalized = records.map((record) => ({
    code: String(record.code || '').toUpperCase(),
    price: Number(record.price),
    currency: String(record.currency || ''),
    unit: String(record.unit || ''),
    source: String(record.source || ''),
    timestamp: String(record.created_at || record.as_of || record.timestamp || '')
  }));
  if (!normalized.length || normalized.some((record) => !record.code || !Number.isFinite(record.price))) {
    throw new Error('OilPriceAPI response is missing a finite market price.');
  }
  return normalized;
}

function latestPrices_(codes) {
  const reviewed = codes.map(normalizeCode_);
  const payload = requestJson_(
    `/prices/latest?by_code=${encodeURIComponent(reviewed.join(','))}`,
    requireApiKey_()
  );
  const records = priceRecords_(payload);
  const missing = reviewed.filter((code) => !records.some((record) => record.code === code));
  if (missing.length) throw new Error(`OilPriceAPI response did not include: ${missing.join(', ')}.`);
  return reviewed.map((code) => records.find((record) => record.code === code));
}

function history_(code, days) {
  const reviewed = normalizeCode_(code);
  const requestedDays = Number(days);
  if (![7, 30, 365].includes(requestedDays)) throw new Error('History window must be 7, 30, or 365 days.');
  const endpoint = requestedDays === 7 ? 'past_week' : requestedDays === 30 ? 'past_month' : 'past_year';
  return priceRecords_(requestJson_(
    `/prices/${endpoint}?by_code=${encodeURIComponent(reviewed)}`,
    requireApiKey_()
  ));
}

function testConnection() {
  const probe = OPA_PRODUCT.allowedCodes.length
    ? latestPrices_([OPA_PRODUCT.allowedCodes[0]])[0]
    : productConnectionProbe_();
  return { success: true, code: probe.code || probe.contract || 'curve', timestamp: probe.timestamp || '' };
}

function activateProduct_() {
  const properties = documentProperties_();
  if (properties) properties.setProperty(OPA_ACTIVATED_PROPERTY, new Date().toISOString());
}

function sheet_(name) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const existing = workbook.getSheetByName(name);
  const sheet = existing || workbook.insertSheet(name);
  sheet.clear();
  return sheet;
}

function writeTable_(sheet, startRow, startColumn, rows) {
  if (!rows.length || !rows[0].length) throw new Error('Cannot write an empty table.');
  sheet.getRange(startRow, startColumn, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(startRow, startColumn, 1, rows[0].length)
    .setBackground('#0f3557')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.autoResizeColumns(startColumn, rows[0].length);
}

function writeTitle_(sheet, title, subtitle) {
  sheet.getRange('A1').setValue(title).setFontSize(18).setFontWeight('bold').setFontColor('#0f3557');
  sheet.getRange('A2').setValue(subtitle).setFontColor('#486581');
}

function marketDataRows_(records) {
  return [['Market', 'Price', 'Currency', 'Unit', 'Source', 'Source timestamp']].concat(
    records.map((record) => [record.code, record.price, record.currency, record.unit, record.source, record.timestamp])
  );
}

function normalizeGasMarkets_(henryHubUsdMmbtu, ttfEurMwh, jkmUsdMmbtu, eurUsd) {
  const values = [henryHubUsdMmbtu, ttfEurMwh, jkmUsdMmbtu, eurUsd].map(Number);
  if (!values.every(Number.isFinite) || values[3] <= 0) throw new Error('Gas-market inputs and FX must be finite positive values.');
  const ttfUsdMmbtu = (values[1] * values[3]) / 3.412141633;
  return {
    henryHub: values[0],
    ttf: ttfUsdMmbtu,
    jkm: values[2],
    ttfHenryHub: ttfUsdMmbtu - values[0],
    jkmHenryHub: values[2] - values[0],
    jkmTtf: values[2] - ttfUsdMmbtu
  };
}

function buildGasSpreadWorkbook() {
  const records = latestPrices_(OPA_PRODUCT.allowedCodes);
  const price = Object.fromEntries(records.map((record) => [record.code, record.price]));
  const normalized = normalizeGasMarkets_(
    price.NATURAL_GAS_USD,
    price.DUTCH_TTF_EUR,
    price.JKM_LNG_USD,
    price.EUR_USD
  );

  const data = sheet_('Market Data');
  writeTitle_(data, 'Global Gas Market Data', 'Native prices and units returned by OilPriceAPI.');
  writeTable_(data, 4, 1, marketDataRows_(records));

  const audit = sheet_('Conversion Audit');
  writeTitle_(audit, 'TTF Conversion Audit', 'TTF EUR/MWh × EUR/USD ÷ 3.412141633 MMBtu/MWh.');
  writeTable_(audit, 4, 1, [
    ['Input', 'Value', 'Unit'],
    ['TTF native', price.DUTCH_TTF_EUR, 'EUR/MWh'],
    ['EUR/USD', price.EUR_USD, 'USD per EUR'],
    ['Energy conversion', 3.412141633, 'MMBtu per MWh'],
    ['TTF normalized', normalized.ttf, 'USD/MMBtu']
  ]);
  audit.getRange('B5:B8').setNumberFormat('0.0000');

  const monitor = sheet_('Gas Spread Monitor');
  writeTitle_(monitor, 'Natural Gas Spread Monitor', 'Cross-market comparison after explicit currency and energy-unit normalization.');
  writeTable_(monitor, 4, 1, [
    ['Market', 'Normalized price', 'Unit'],
    ['Henry Hub', normalized.henryHub, 'USD/MMBtu'],
    ['Dutch TTF', normalized.ttf, 'USD/MMBtu'],
    ['JKM LNG', normalized.jkm, 'USD/MMBtu']
  ]);
  writeTable_(monitor, 10, 1, [
    ['Spread', 'Value', 'Unit'],
    ['TTF minus Henry Hub', normalized.ttfHenryHub, 'USD/MMBtu'],
    ['JKM minus Henry Hub', normalized.jkmHenryHub, 'USD/MMBtu'],
    ['JKM minus TTF', normalized.jkmTtf, 'USD/MMBtu']
  ]);
  monitor.getRange('B5:B7').setNumberFormat('$0.000');
  monitor.getRange('B11:B13').setNumberFormat('$0.000');

  activateProduct_();
  monitor.activate();
  return { success: true, message: 'Gas spread monitor built with Henry Hub, TTF, and JKM normalized to USD/MMBtu.' };
}
