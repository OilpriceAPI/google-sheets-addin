const OPA_PRODUCT = Object.freeze({
  "id": "energy-curve-builder",
  "name": "Energy Curve Builder by OilPriceAPI",
  "menu": "Energy Curve Builder",
  "version": "1.0.0",
  "cloudProjectId": "oilpriceapi-energy-curve",
  "iconMark": "EC",
  "brandColor": "#7C3AED",
  "builder": "buildEnergyCurveWorkbook",
  "tagline": "Create futures curves, calendar spreads, and structure signals.",
  "landingPath": "/integrations/energy-curve-builder",
  "signupCampaign": "energy_curve_builder",
  "activationHeader": "energy-curve-builder",
  "allowedCodes": [],
  "sheets": [
    "Curve Dashboard",
    "WTI Curve",
    "Brent Curve",
    "Calendar Spreads"
  ],
  "workflow": "Creates WTI and Brent term-structure tables, month-on-month calendar spreads, and contango/backwardation signals.",
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

  // Prototype releases used an unscoped user property. It cannot be tied to a
  // source spreadsheet safely, so require the user to save the key again.
  return null;
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
        'X-API-Client': `${OPA_PRODUCT.activationHeader}/${OPA_VERSION}`
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

function curveContracts_(payload) {
  const data = payload && payload.data;
  const records = payload && Array.isArray(payload.contracts)
    ? payload.contracts
    : Array.isArray(data)
      ? data
      : data && Array.isArray(data.contracts)
        ? data.contracts
        : [];
  const normalized = records.map((record) => ({
    month: String(record.contract_month || record.month || record.symbol || ''),
    price: Number(record.settlement_price ?? record.settlement ?? record.price),
    sourceTimestamp: String(record.trading_date || record.created_at || record.as_of || record.timestamp || '')
  }));
  if (normalized.length < 2 || normalized.some((record) => !record.month || !Number.isFinite(record.price))) {
    throw new Error('Futures curve response needs at least two dated contracts with finite prices.');
  }
  return normalized;
}

function fetchCurve_(contract) {
  const code = String(contract || '').trim().toLowerCase();
  if (!['ice-wti', 'ice-brent'].includes(code)) {
    throw new Error('Supported curve contracts are ice-wti and ice-brent.');
  }
  return curveContracts_(requestJson_(`/futures/${encodeURIComponent(code)}/curve`, requireApiKey_()));
}

function curveTable_(records) {
  const rows = [['Contract month', 'Price', 'Calendar spread', 'Structure']];
  records.forEach((record, index) => {
    const spread = index === 0 ? '' : record.price - records[index - 1].price;
    const structure = index === 0 ? 'Front month' : spread > 0 ? 'Contango step' : spread < 0 ? 'Backwardation step' : 'Flat step';
    rows.push([record.month, record.price, spread, structure]);
  });
  return rows;
}

function productConnectionProbe_() {
  const record = fetchCurve_('ice-wti')[0];
  return { contract: record.month, timestamp: record.sourceTimestamp };
}

function buildEnergyCurveWorkbook() {
  const wti = fetchCurve_('ice-wti');
  const brent = fetchCurve_('ice-brent');

  const wtiSheet = sheet_('WTI Curve');
  writeTitle_(wtiSheet, 'WTI Futures Curve', 'Contract prices and month-on-month calendar spreads.');
  writeTable_(wtiSheet, 4, 1, curveTable_(wti));
  wtiSheet.getRange(5, 2, wti.length, 2).setNumberFormat('$0.00');

  const brentSheet = sheet_('Brent Curve');
  writeTitle_(brentSheet, 'Brent Futures Curve', 'Contract prices and month-on-month calendar spreads.');
  writeTable_(brentSheet, 4, 1, curveTable_(brent));
  brentSheet.getRange(5, 2, brent.length, 2).setNumberFormat('$0.00');

  const spreads = sheet_('Calendar Spreads');
  writeTitle_(spreads, 'Front Calendar Spreads', 'Positive back-month minus front-month values indicate an upward-sloping step.');
  writeTable_(spreads, 4, 1, [
    ['Market', 'Front month', 'Second month', 'Second minus front', 'Signal'],
    ['WTI', wti[0].month, wti[1].month, wti[1].price - wti[0].price, wti[1].price > wti[0].price ? 'Contango' : 'Backwardation'],
    ['Brent', brent[0].month, brent[1].month, brent[1].price - brent[0].price, brent[1].price > brent[0].price ? 'Contango' : 'Backwardation']
  ]);
  spreads.getRange('D5:E6').setFormulas([
    ['=\'WTI Curve\'!B6-\'WTI Curve\'!B5', '=IF(D5>0,"Contango",IF(D5<0,"Backwardation","Flat"))'],
    ['=\'Brent Curve\'!B6-\'Brent Curve\'!B5', '=IF(D6>0,"Contango",IF(D6<0,"Backwardation","Flat"))']
  ]);
  spreads.getRange('D5:D6').setNumberFormat('$0.00');

  const dashboard = sheet_('Curve Dashboard');
  writeTitle_(dashboard, 'Energy Futures Curve Builder', 'A term-structure workbook, not a generic quote downloader.');
  writeTable_(dashboard, 4, 1, [
    ['Market', 'Front month', 'Front price', 'Curve points', 'Front structure'],
    ['WTI', wti[0].month, wti[0].price, wti.length, wti[1].price > wti[0].price ? 'Contango' : 'Backwardation'],
    ['Brent', brent[0].month, brent[0].price, brent.length, brent[1].price > brent[0].price ? 'Contango' : 'Backwardation']
  ]);
  dashboard.getRange('C5:C6').setNumberFormat('$0.00');

  activateProduct_();
  dashboard.activate();
  return { success: true, message: 'Energy curve workbook built with WTI and Brent curves, calendar spreads, and structure signals.' };
}
