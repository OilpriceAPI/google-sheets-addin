const OPA_PRODUCT = Object.freeze({
  "id": "fuel-surcharge-studio",
  "name": "Fuel Surcharge Studio by OilPriceAPI",
  "menu": "Fuel Surcharge Studio",
  "version": "1.0.0",
  "cloudProjectId": "oilpriceapi-fuel-surcharge",
  "iconMark": "FS",
  "brandColor": "#15803D",
  "builder": "buildFuelSurchargeWorkbook",
  "tagline": "Turn diesel indexes into auditable carrier surcharge schedules.",
  "landingPath": "/integrations/fuel-surcharge-studio",
  "signupCampaign": "fuel_surcharge_studio",
  "activationHeader": "fuel-surcharge-studio",
  "allowedCodes": [
    "DIESEL_RETAIL_USD",
    "DIESEL_RETAIL_EAST_COAST_USD",
    "DIESEL_RETAIL_MIDWEST_USD",
    "DIESEL_RETAIL_GULF_COAST_USD",
    "DIESEL_RETAIL_ROCKY_MOUNTAIN_USD",
    "DIESEL_RETAIL_WEST_COAST_USD"
  ],
  "sheets": [
    "Surcharge Calculator",
    "Regional Indexes",
    "Surcharge Table"
  ],
  "workflow": "Creates a freight surcharge workbook with regional diesel indexes, a transparent base-price formula, and a publishable band schedule.",
  "category": "Accounting and Finance"
});

/**
 * Shared security and workbook runtime for the OilPriceAPI Workspace portfolio.
 * Product behavior is supplied by exactly one Product.gs file at build time.
 */

const OPA_API_BASE_URL = 'https://api.oilpriceapi.com/v1';
const OPA_KEY_PROPERTY = 'OILPRICEAPI_KEY';
const OPA_ACTIVATED_PROPERTY = 'OILPRICEAPI_ACTIVATED_AT';
const OPA_VERSION = OPA_PRODUCT.version;
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
  html.version = OPA_VERSION;
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

function activeSpreadsheetId_() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet && typeof spreadsheet.getId === 'function'
      ? spreadsheet.getId()
      : null;
  } catch (error) {
    return null;
  }
}

function spreadsheetKeyProperty_() {
  const spreadsheetId = activeSpreadsheetId_();
  return spreadsheetId ? `${OPA_KEY_PROPERTY}:${spreadsheetId}` : null;
}

function getApiKey_() {
  const documentProperties = documentProperties_();
  const documentKey = documentProperties
    ? documentProperties.getProperty(OPA_KEY_PROPERTY)
    : null;
  if (documentKey) return documentKey;

  // Installed add-on custom functions can run in a separate Apps Script
  // authorization context where document properties are unavailable. Keep a
  // compatibility copy in owner user properties, namespaced by spreadsheet ID,
  // so the key cannot cross into another workbook.
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetKeyProperty = spreadsheetKeyProperty_();
  const spreadsheetKey = spreadsheetKeyProperty
    ? userProperties.getProperty(spreadsheetKeyProperty)
    : null;
  if (spreadsheetKey) return spreadsheetKey;

  // Migration fallback for the two prototype releases. Saving again removes
  // this unscoped value.
  return userProperties.getProperty(OPA_KEY_PROPERTY);
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
  const documentProperties = documentProperties_();
  const spreadsheetKeyProperty = spreadsheetKeyProperty_();
  if (!documentProperties || !spreadsheetKeyProperty) {
    throw new Error('Open the add-on from a spreadsheet before saving a key.');
  }
  documentProperties.setProperty(OPA_KEY_PROPERTY, value);
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(spreadsheetKeyProperty, value);
  userProperties.deleteProperty(OPA_KEY_PROPERTY);
  return { success: true, configured: true };
}

function deleteApiKey() {
  const documentProperties = documentProperties_();
  if (documentProperties) {
    documentProperties.deleteProperty(OPA_KEY_PROPERTY);
    documentProperties.deleteProperty(OPA_ACTIVATED_PROPERTY);
  }
  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetKeyProperty = spreadsheetKeyProperty_();
  if (spreadsheetKeyProperty) userProperties.deleteProperty(spreadsheetKeyProperty);
  userProperties.deleteProperty(OPA_KEY_PROPERTY);
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
  let response;
  try {
    response = UrlFetchApp.fetch(`${OPA_API_BASE_URL}${normalizedPath}`, {
      method: 'get',
      headers: {
        Authorization: `Token ${apiKey}`,
        Accept: 'application/json',
        'X-OilPriceAPI-Client': `${OPA_PRODUCT.activationHeader}/${OPA_VERSION}`
      },
      muteHttpExceptions: true
    });
  } catch (error) {
    throw new Error('OilPriceAPI could not be reached. Check the connection and retry.');
  }
  const status = response.getResponseCode();
  if (status === 401) throw new Error('The OilPriceAPI key is invalid or revoked. Replace it in the sidebar.');
  if (status === 402 || status === 403) throw new Error('This dataset is not enabled for the account. Review OilPriceAPI pricing or use an entitled key.');
  if (status === 408) throw new Error('The OilPriceAPI request timed out. Retry in a moment.');
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
  const normalized = records.map((record) => {
    const rawPrice = record && record.price;
    return {
      code: String(record && record.code || '').toUpperCase(),
      price: rawPrice === null || rawPrice === undefined || rawPrice === ''
        ? NaN
        : Number(rawPrice),
      currency: String(record && record.currency || ''),
      unit: String(record && record.unit || ''),
      source: String(record && record.source || ''),
      timestamp: String(record && (record.created_at || record.as_of || record.timestamp) || '')
    };
  });
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
  return {
    success: true,
    code: probe.code || probe.contract || 'curve',
    timestamp: probe.timestamp || '',
    message: 'Connection and response schema verified.'
  };
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

function surchargePerMile_(indexPrice, basePrice, milesPerGallon) {
  const index = Number(indexPrice);
  const base = Number(basePrice);
  const mpg = Number(milesPerGallon);
  if (![index, base, mpg].every(Number.isFinite)) throw new Error('Surcharge inputs must be finite.');
  if (index < 0 || base < 0) throw new Error('Diesel index and base price must be non-negative.');
  if (mpg <= 0) throw new Error('MPG must be positive.');
  return Math.max(0, index - base) / mpg;
}

function surchargeBands_(basePrice, milesPerGallon, startPrice, step, bandCount) {
  const count = Number(bandCount);
  if (!Number.isInteger(count) || count < 1 || count > 100) throw new Error('Band count must be an integer from 1 to 100.');
  if (![basePrice, milesPerGallon, startPrice, step].map(Number).every(Number.isFinite)) {
    throw new Error('Band inputs must be finite.');
  }
  if (Number(startPrice) < 0) throw new Error('Band start price must be non-negative.');
  if (Number(step) <= 0) throw new Error('Band step must be positive.');
  const rows = [['Diesel index from', 'Diesel index through', 'Surcharge per mile']];
  for (let index = 0; index < count; index += 1) {
    const lower = Number(startPrice) + (index * Number(step));
    const upper = lower + Number(step) - 0.001;
    rows.push([lower, upper, surchargePerMile_(lower, basePrice, milesPerGallon)]);
  }
  return rows;
}

function buildFuelSurchargeWorkbook() {
  const records = latestPrices_(OPA_PRODUCT.allowedCodes);
  const national = records.find((record) => record.code === 'DIESEL_RETAIL_USD');
  const basePrice = Math.max(0, national.price - 0.75);
  const mpg = 6.5;

  const indexes = sheet_('Regional Indexes');
  writeTitle_(indexes, 'Regional Diesel Indexes', 'Latest available retail diesel values returned by OilPriceAPI.');
  writeTable_(indexes, 4, 1, marketDataRows_(records));
  indexes.getRange(5, 2, records.length, 1).setNumberFormat('$0.000');

  const calculator = sheet_('Surcharge Calculator');
  writeTitle_(calculator, 'Fuel Surcharge Studio', 'Transparent index-minus-base calculation; confirm the method in each carrier agreement.');
  writeTable_(calculator, 4, 1, [
    ['Input', 'Value', 'Unit'],
    ['Current national diesel index', national.price, 'USD/gallon'],
    ['Contract base diesel price', basePrice, 'USD/gallon'],
    ['Fleet fuel economy', mpg, 'miles/gallon'],
    ['Calculated surcharge', surchargePerMile_(national.price, basePrice, mpg), 'USD/mile']
  ]);
  calculator.getRange('B8').setFormula('=MAX(0,B5-B6)/B7');
  calculator.getRange('B5:B6').setNumberFormat('$0.000');
  calculator.getRange('B8').setNumberFormat('$0.000');

  const schedule = sheet_('Surcharge Table');
  writeTitle_(schedule, 'Publishable Surcharge Bands', 'Twenty five-cent diesel-index bands using the calculator assumptions.');
  const bands = surchargeBands_(basePrice, mpg, Math.max(0, basePrice), 0.25, 20);
  writeTable_(schedule, 4, 1, bands);
  const bandFormulas = Array.from({ length: 20 }, (_, index) => {
    const row = index + 5;
    return [
      `='Surcharge Calculator'!$B$6+(ROW()-5)*0.25`,
      `=A${row}+0.249`,
      `=MAX(0,A${row}-'Surcharge Calculator'!$B$6)/'Surcharge Calculator'!$B$7`
    ];
  });
  schedule.getRange(5, 1, 20, 3).setFormulas(bandFormulas);
  schedule.getRange(5, 1, 20, 3).setNumberFormat('$0.000');

  activateProduct_();
  calculator.activate();
  return { success: true, message: 'Fuel surcharge workbook built with live regional indexes and a 20-band schedule.' };
}
