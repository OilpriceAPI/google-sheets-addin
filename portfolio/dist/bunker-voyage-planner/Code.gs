const OPA_PRODUCT = Object.freeze({
  "id": "bunker-voyage-planner",
  "name": "Bunker Voyage Planner by OilPriceAPI",
  "menu": "Bunker Voyage Planner",
  "builder": "buildBunkerVoyageWorkbook",
  "tagline": "Compare port fuel choices and calculate voyage bunker cost.",
  "landingPath": "/integrations/bunker-voyage-planner",
  "signupCampaign": "bunker_voyage_planner",
  "activationHeader": "bunker-voyage-planner",
  "allowedCodes": [
    "VLSFO_SGSIN_USD",
    "MGO_05S_SGSIN_USD",
    "VLSFO_NLRTM_USD",
    "MGO_05S_NLRTM_USD",
    "VLSFO_USHOU_USD",
    "MGO_05S_USHOU_USD"
  ],
  "sheets": [
    "Voyage Plan",
    "Port Prices",
    "Scenario Compare"
  ],
  "workflow": "Creates a voyage-level fuel budget using vessel consumption, sailing days, port prices, fuel mix, and scenario comparisons.",
  "category": "Productivity"
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

function calculateVoyageFuelCost_(seaDays, seaConsumption, portDays, portConsumption, vlsfoShare, vlsfoPrice, mgoPrice) {
  const values = [seaDays, seaConsumption, portDays, portConsumption, vlsfoShare, vlsfoPrice, mgoPrice].map(Number);
  if (!values.every(Number.isFinite)) throw new Error('Voyage-cost inputs must be finite numbers.');
  if (values[4] < 0 || values[4] > 1) throw new Error('VLSFO share must be between 0 and 1.');
  const tonnes = (values[0] * values[1]) + (values[2] * values[3]);
  const blendedPrice = (values[4] * values[5]) + ((1 - values[4]) * values[6]);
  return { tonnes, blendedPrice, totalCost: tonnes * blendedPrice };
}

function bunkerPortRows_(records) {
  const labels = {
    VLSFO_SGSIN_USD: ['Singapore', 'VLSFO'],
    MGO_05S_SGSIN_USD: ['Singapore', 'MGO 0.5%'],
    VLSFO_NLRTM_USD: ['Rotterdam', 'VLSFO'],
    MGO_05S_NLRTM_USD: ['Rotterdam', 'MGO 0.5%'],
    VLSFO_USHOU_USD: ['Houston', 'VLSFO'],
    MGO_05S_USHOU_USD: ['Houston', 'MGO 0.5%']
  };
  return [['Port', 'Fuel', 'Price', 'Currency', 'Unit', 'Source timestamp']].concat(
    records.map((record) => [labels[record.code][0], labels[record.code][1], record.price, record.currency, record.unit, record.timestamp])
  );
}

function buildBunkerVoyageWorkbook() {
  const records = latestPrices_(OPA_PRODUCT.allowedCodes);
  const price = Object.fromEntries(records.map((record) => [record.code, record.price]));
  const singapore = calculateVoyageFuelCost_(12, 28, 2, 4, 0.9, price.VLSFO_SGSIN_USD, price.MGO_05S_SGSIN_USD);
  const rotterdam = calculateVoyageFuelCost_(12, 28, 2, 4, 0.9, price.VLSFO_NLRTM_USD, price.MGO_05S_NLRTM_USD);
  const houston = calculateVoyageFuelCost_(12, 28, 2, 4, 0.9, price.VLSFO_USHOU_USD, price.MGO_05S_USHOU_USD);

  const ports = sheet_('Port Prices');
  writeTitle_(ports, 'Bunker Port Prices', 'Latest available marine-fuel values returned by OilPriceAPI.');
  writeTable_(ports, 4, 1, bunkerPortRows_(records));
  ports.getRange(5, 3, records.length, 1).setNumberFormat('$0.00');

  const plan = sheet_('Voyage Plan');
  writeTitle_(plan, 'Bunker Voyage Cost Planner', 'Edit vessel and voyage assumptions; compare live port pricing in Scenario Compare.');
  writeTable_(plan, 4, 1, [
    ['Assumption', 'Value', 'Unit'],
    ['Sea days', 12, 'days'],
    ['Sea consumption', 28, 'metric tonnes/day'],
    ['Port days', 2, 'days'],
    ['Port consumption', 4, 'metric tonnes/day'],
    ['VLSFO share', 0.9, 'fraction'],
    ['Total fuel', singapore.tonnes, 'metric tonnes']
  ]);
  plan.getRange('B9').setNumberFormat('0.0%');

  const compare = sheet_('Scenario Compare');
  writeTitle_(compare, 'Bunkering Scenario Comparison', 'Same voyage and fuel mix, different bunkering ports.');
  writeTable_(compare, 4, 1, [
    ['Port', 'Fuel tonnes', 'Blended price', 'Estimated voyage fuel cost', 'Difference vs lowest'],
    ['Singapore', singapore.tonnes, singapore.blendedPrice, singapore.totalCost, singapore.totalCost - Math.min(singapore.totalCost, rotterdam.totalCost, houston.totalCost)],
    ['Rotterdam', rotterdam.tonnes, rotterdam.blendedPrice, rotterdam.totalCost, rotterdam.totalCost - Math.min(singapore.totalCost, rotterdam.totalCost, houston.totalCost)],
    ['Houston', houston.tonnes, houston.blendedPrice, houston.totalCost, houston.totalCost - Math.min(singapore.totalCost, rotterdam.totalCost, houston.totalCost)]
  ]);
  compare.getRange(5, 3, 3, 3).setNumberFormat('$#,##0.00');

  activateProduct_();
  compare.activate();
  return { success: true, message: 'Bunker voyage workbook built with three live port scenarios and vessel-level fuel economics.' };
}
