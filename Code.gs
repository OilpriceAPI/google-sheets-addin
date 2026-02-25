/**
 * OilPriceAPI Google Sheets Add-on
 * Real-time commodity price data
 */

// API Configuration
const API_BASE_URL = 'https://api.oilpriceapi.com/v1';
const RATE_LIMIT_CACHE_KEY = 'historical_fetch_timestamps';

// Commodity mapping
const COMMODITY_MAP = {
  'BRENT_CRUDE_USD': { type: 'BRENT_CRUDE_OIL', unit: 'barrel' },
  'WTI_USD': { type: 'WTI_CRUDE_OIL', unit: 'barrel' },
  'NATURAL_GAS_USD': { type: 'NATURAL_GAS', unit: 'MBtu' },
  'NATURAL_GAS_GBP': { type: 'NATURAL_GAS', unit: 'therm' },
  'DUTCH_TTF_EUR': { type: 'NATURAL_GAS', unit: 'MWh' },
  'COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  // US Spot Coal
  'CAPP_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'PRB_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'ILLINOIS_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  // International Coal Futures
  'NEWCASTLE_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  'COKING_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'tonne' },
  'CME_COAL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  // NYMEX Historical (Discontinued)
  'NYMEX_APPALACHIAN_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' },
  'NYMEX_WESTERN_RAIL_USD': { type: 'COAL_BITUMINOUS', unit: 'short_ton' }
};

// Heat content factors (MMBtu per unit)
const HEAT_CONTENT = {
  'BRENT_CRUDE_OIL': 5.8,
  'WTI_CRUDE_OIL': 5.8,
  'NATURAL_GAS': 1.037,  // For Mcf
  'COAL_BITUMINOUS': 24.0
};

/**
 * Add menu to Google Sheets UI
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OilPriceAPI')
    .addItem('Configure API Key', 'showSidebar')
    .addItem('Fetch Latest Prices', 'showFetchDialog')
    .addItem('Convert to $/MMBtu', 'convertToMBtu')
    .addSeparator()
    .addItem('Fetch Bunker Prices (Data Connector)', 'fetchDataConnectorPrices')
    .addItem('Fetch Futures Data', 'showFuturesInfo')
    .addItem('Fetch Rig Counts', 'showRigCountInfo')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Show sidebar for configuration
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('OilPriceAPI Configuration')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Show about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'OilPriceAPI for Google Sheets',
    'Version 1.0.0\n\n' +
    'Real-time oil & commodity price data\n\n' +
    'Website: https://oilpriceapi.com\n' +
    'Support: support@oilpriceapi.com\n\n' +
    '© 2025 OilPriceAPI',
    ui.ButtonSet.OK
  );
}

/**
 * Get API key from user properties
 */
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('OILPRICEAPI_KEY');
}

/**
 * Save API key to user properties
 */
function saveApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('OILPRICEAPI_KEY', apiKey);
  return { success: true, message: 'API key saved successfully!' };
}

/**
 * Test API connection
 */
function testConnection() {
  const apiKey = getApiKey();

  if (!apiKey) {
    return { success: false, message: 'No API key found. Please configure your API key first.' };
  }

  try {
    const url = `${API_BASE_URL}/prices/latest?by_code=BRENT_CRUDE_USD`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      return { success: true, message: 'Connection successful! ✓' };
    } else if (statusCode === 401) {
      return { success: false, message: 'Invalid API key. Please check and try again.' };
    } else {
      return { success: false, message: `Error: HTTP ${statusCode}` };
    }
  } catch (error) {
    return { success: false, message: `Connection failed: ${error.message}` };
  }
}

/**
 * Get user tier information
 */
function getUserInfo() {
  const apiKey = getApiKey();

  if (!apiKey) {
    return { tier: 'none', limit: 0, used: 0 };
  }

  try {
    const url = `${API_BASE_URL}/users/me`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    return {
      tier: data.tier || 'free',
      limit: data.request_limit || 1000,
      used: data.requests_this_month || 0
    };
  } catch (error) {
    return { tier: 'unknown', limit: 0, used: 0 };
  }
}

/**
 * Fetch latest prices for selected commodities
 */
function fetchLatestPrices(commodityCodes) {
  const apiKey = getApiKey();

  if (!apiKey) {
    throw new Error('No API key configured. Please set your API key first.');
  }

  try {
    const codes = commodityCodes.join(',');
    const url = `${API_BASE_URL}/prices/latest?by_code=${codes}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`API request failed: HTTP ${statusCode}`);
    }

    const data = JSON.parse(response.getContentText());
    const prices = data.data.prices || [];

    // Write to Data sheet
    writeToDataSheet(prices);

    return { success: true, count: prices.length };
  } catch (error) {
    throw new Error(`Failed to fetch prices: ${error.message}`);
  }
}

/**
 * Write price data to Data sheet
 */
function writeToDataSheet(prices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data');

  // Create sheet if doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Data');
  }

  // Clear existing data
  sheet.clear();

  // Add headers
  const headers = [['Commodity Code', 'Price', 'Currency', 'Unit', 'Timestamp', 'Last Updated']];
  sheet.getRange(1, 1, 1, 6).setValues(headers);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f0f0f0');

  // Add data rows
  const rows = prices.map(p => [
    p.code,
    p.price,
    p.currency || 'USD',
    COMMODITY_MAP[p.code]?.unit || 'unknown',
    p.created_at,
    new Date().toISOString()
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);

    // Format price column
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0.00');
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);

  // Activate the sheet
  sheet.activate();
}

/**
 * Convert prices to $/MMBtu
 */
function convertToMBtu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('No Data sheet found. Please fetch prices first.');
    return;
  }

  // Get data from Data sheet
  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No price data found. Please fetch prices first.');
    return;
  }

  // Fetch exchange rates
  const rates = fetchExchangeRates();

  // Create or get Process sheet
  let processSheet = ss.getSheetByName('Process');
  if (!processSheet) {
    processSheet = ss.insertSheet('Process');
  }

  processSheet.clear();

  // Add headers
  const headers = [['Commodity', 'Original Price', 'Currency', 'Unit', 'USD Price', 'Heat Content (MMBtu)', 'Price per MBtu (USD)']];
  processSheet.getRange(1, 1, 1, 7).setValues(headers);
  processSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e3f2fd');

  // Process data rows
  const processRows = [];

  for (let i = 1; i < data.length; i++) {
    const [code, price, currency, unit] = data[i];

    if (!code || !COMMODITY_MAP[code]) continue;

    const commodityInfo = COMMODITY_MAP[code];

    // Get heat content based on unit
    let heatContent;
    if (commodityInfo.unit === 'therm') {
      heatContent = 0.1;  // 1 therm = 0.1 MMBtu
    } else if (commodityInfo.unit === 'MWh') {
      heatContent = 3.412;  // 1 MWh = 3.412 MMBtu
    } else if (commodityInfo.unit === 'MBtu') {
      heatContent = 1.0;  // Already in MMBtu
    } else {
      heatContent = HEAT_CONTENT[commodityInfo.type] || 1.0;
    }

    // Convert price to USD
    let usdPrice = parseFloat(price);
    if (currency === 'GBP' || currency === 'GBp') {
      // UK Natural Gas is in pence, convert: pence → pounds → USD
      usdPrice = (parseFloat(price) / 100) * rates.gbpUsd;
    } else if (currency === 'EUR') {
      usdPrice = parseFloat(price) * rates.eurUsd;
    }

    // Calculate price per MMBtu
    let pricePerMBtu;
    if (commodityInfo.unit === 'MBtu') {
      pricePerMBtu = usdPrice;
    } else {
      pricePerMBtu = usdPrice / heatContent;
    }

    processRows.push([
      code,
      parseFloat(price),
      currency || 'USD',
      unit,
      usdPrice,
      heatContent,
      pricePerMBtu
    ]);
  }

  // Write converted data
  if (processRows.length > 0) {
    processSheet.getRange(2, 1, processRows.length, 7).setValues(processRows);

    // Format columns
    processSheet.getRange(2, 2, processRows.length, 1).setNumberFormat('#,##0.00');  // Original price
    processSheet.getRange(2, 5, processRows.length, 1).setNumberFormat('$#,##0.00'); // USD price
    processSheet.getRange(2, 6, processRows.length, 1).setNumberFormat('0.000');     // Heat content
    processSheet.getRange(2, 7, processRows.length, 1).setNumberFormat('$#,##0.00'); // Price per MBtu
  }

  // Auto-resize columns
  processSheet.autoResizeColumns(1, 7);

  // Activate Process sheet
  processSheet.activate();

  SpreadsheetApp.getUi().alert(`Converted ${processRows.length} commodities to $/MMBtu`);
}

/**
 * Fetch exchange rates
 */
function fetchExchangeRates() {
  const apiKey = getApiKey();

  try {
    const url = `${API_BASE_URL}/prices/latest?by_code=GBP_USD,EUR_USD`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const prices = data.data.prices || [];

    const gbpRate = prices.find(p => p.code === 'GBP_USD')?.price || 1.30;
    const eurRate = prices.find(p => p.code === 'EUR_USD')?.price || 1.10;

    return { gbpUsd: gbpRate, eurUsd: eurRate };
  } catch (error) {
    // Return fallback rates
    return { gbpUsd: 1.30, eurUsd: 1.10 };
  }
}

/**
 * Custom function: Get latest price for a commodity
 * @param {string} commodityCode The commodity code (e.g., "BRENT_CRUDE_USD")
 * @return {number} The latest price
 * @customfunction
 */
function OILPRICE(commodityCode) {
  if (!commodityCode) {
    throw new Error('Commodity code is required');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured. Use OilPriceAPI menu to configure.');
  }

  // Check cache first
  const cache = CacheService.getUserCache();
  const cacheKey = 'price_' + commodityCode;
  const cached = cache.get(cacheKey);
  if (cached) {
    return parseFloat(cached);
  }

  try {
    const url = `${API_BASE_URL}/prices/latest?by_code=${commodityCode}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data.data && data.data.prices && data.data.prices.length > 0) {
      const price = data.data.prices[0].price;
      // Cache for 5 minutes (live data)
      cache.put(cacheKey, price.toString(), 300);
      return price;
    }

    throw new Error('No price data found');
  } catch (error) {
    throw new Error(`Failed to fetch price: ${error.message}`);
  }
}

/**
 * Custom function: Get historical prices
 * @param {string} commodityCode The commodity code
 * @param {number} days Number of days of history
 * @return {Array} Array of prices
 * @customfunction
 */
function OILPRICE_HISTORY(commodityCode, days) {
  if (!commodityCode) {
    throw new Error('Commodity code is required');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured');
  }

  days = days || 30;

  // Check cache first
  const cache = CacheService.getUserCache();
  const cacheKey = 'history_' + commodityCode + '_' + days;
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  try {
    let endpoint;
    if (days <= 1) {
      endpoint = 'past_day';
    } else if (days <= 7) {
      endpoint = 'past_week';
    } else if (days <= 30) {
      endpoint = 'past_month';
    } else {
      endpoint = 'past_year';
    }
    const url = `${API_BASE_URL}/prices/${endpoint}?by_code=${commodityCode}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data.data && data.data.prices) {
      const result = data.data.prices.map(p => [p.created_at, p.price]);
      // Cache for 1 hour
      cache.put(cacheKey, JSON.stringify(result), 3600);
      return result;
    }

    throw new Error('No price data found');
  } catch (error) {
    throw new Error(`Failed to fetch history: ${error.message}`);
  }
}

/**
 * Custom function: Convert commodity price to $/MMBtu
 * @param {string} commodityCode The commodity code
 * @return {number} Price in $/MMBtu
 * @customfunction
 */
function OILPRICE_CONVERT(commodityCode) {
  if (!commodityCode) {
    throw new Error('Commodity code is required');
  }

  const price = OILPRICE(commodityCode);
  const commodityInfo = COMMODITY_MAP[commodityCode];

  if (!commodityInfo) {
    throw new Error('Unknown commodity code');
  }

  // Get heat content based on unit
  let heatContent;
  if (commodityInfo.unit === 'therm') {
    heatContent = 0.1;
  } else if (commodityInfo.unit === 'MWh') {
    heatContent = 3.412;
  } else if (commodityInfo.unit === 'MBtu') {
    return price;  // Already in MMBtu
  } else {
    heatContent = HEAT_CONTENT[commodityInfo.type] || 1.0;
  }

  return price / heatContent;
}

/**
 * Show fetch prices dialog
 */
function showFetchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('FetchDialog')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fetch Latest Prices');
}

// ============================================================================
// DATA CONNECTOR (BYOS) - Bunker Fuel Prices
// ============================================================================

/**
 * Fetch bunker fuel prices from Data Connector (BYOS)
 * Requires Data Connector feature enabled on your organization
 */
function fetchDataConnectorPrices() {
  const apiKey = getApiKey();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('No API key configured. Please set your API key first.');
    return;
  }

  try {
    const url = `${API_BASE_URL}/prices/data-connector`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 403) {
      SpreadsheetApp.getUi().alert(
        'Data Connector not enabled',
        'Your organization does not have Data Connector enabled.\n\n' +
        'Contact sales@oilpriceapi.com to enable this feature.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    if (statusCode !== 200) {
      throw new Error(`API request failed: HTTP ${statusCode}`);
    }

    const data = JSON.parse(response.getContentText());
    const prices = data.data?.prices || [];

    if (prices.length === 0) {
      SpreadsheetApp.getUi().alert('No bunker prices found. Check your Data Connector configuration.');
      return;
    }

    // Write to Bunker Prices sheet
    writeToDataConnectorSheet(prices);

    SpreadsheetApp.getUi().alert(`Fetched ${prices.length} bunker fuel prices.`);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Failed to fetch bunker prices: ${error.message}`);
  }
}

/**
 * Write Data Connector prices to Bunker Prices sheet
 */
function writeToDataConnectorSheet(prices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bunker Prices');

  // Create sheet if doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Bunker Prices');
  }

  // Clear existing data
  sheet.clear();

  // Add headers
  const headers = [['Port', 'Fuel Type', 'Price', 'Currency', 'Unit', 'Region', 'Source', 'Timestamp', 'Last Updated']];
  sheet.getRange(1, 1, 1, 9).setValues(headers);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#e8f5e9');  // Light green for marine theme

  // Add data rows
  const rows = prices.map(p => [
    p.port,
    p.fuel_type,
    p.price,
    p.currency,
    p.unit,
    p.region || '',
    p.source,
    p.timestamp,
    new Date().toISOString()
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 9).setValues(rows);

    // Format price column as currency
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('$#,##0.00');
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 9);

  // Activate the sheet
  sheet.activate();
}

/**
 * Custom function: Get bunker fuel price for a port
 * @param {string} port The port name (e.g., "SINGAPORE", "ROTTERDAM")
 * @param {string} fuelType The fuel type (e.g., "VLSFO", "MGO", "IFO380")
 * @return {number} The latest bunker price in USD/MT
 * @customfunction
 */
function BUNKER_PRICE(port, fuelType) {
  if (!port) {
    throw new Error('Port is required');
  }
  if (!fuelType) {
    throw new Error('Fuel type is required (VLSFO, MGO, or IFO380)');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured. Use OilPriceAPI menu to configure.');
  }

  try {
    const url = `${API_BASE_URL}/prices/data-connector?port=${encodeURIComponent(port.toUpperCase())}&fuel_type=${encodeURIComponent(fuelType.toUpperCase())}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 403) {
      throw new Error('Data Connector not enabled for your organization');
    }

    const data = JSON.parse(response.getContentText());
    const prices = data.data?.prices || [];

    if (prices.length === 0) {
      throw new Error(`No price found for ${fuelType} at ${port}`);
    }

    return prices[0].price;
  } catch (error) {
    throw new Error(`Failed to fetch bunker price: ${error.message}`);
  }
}

/**
 * Custom function: Get all bunker prices for a port
 * Returns array with fuel type and price for all fuel grades
 * @param {string} port The port name (e.g., "SINGAPORE")
 * @return {Array} Array of [fuel_type, price] pairs
 * @customfunction
 */
function BUNKER_PORT_PRICES(port) {
  if (!port) {
    throw new Error('Port is required');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured');
  }

  try {
    const url = `${API_BASE_URL}/prices/data-connector?port=${encodeURIComponent(port.toUpperCase())}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const prices = data.data?.prices || [];

    if (prices.length === 0) {
      return [['No prices found']];
    }

    // Return header + prices
    return [['Fuel Type', 'Price (USD/MT)']].concat(
      prices.map(p => [p.fuel_type, p.price])
    );
  } catch (error) {
    throw new Error(`Failed to fetch: ${error.message}`);
  }
}

// ============================================================================
// FUTURES DATA
// ============================================================================

/**
 * Custom function: Get latest futures price for a contract
 * @param {string} contract The futures contract code ("BZ" for Brent, "CL" for WTI)
 * @return {number} The front-month futures price
 * @customfunction
 */
function FUTURES_PRICE(contract) {
  if (!contract) {
    throw new Error('Contract code is required (BZ for Brent, CL for WTI)');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured. Use OilPriceAPI menu to configure.');
  }

  // Check cache first
  const cache = CacheService.getUserCache();
  const cacheKey = 'futures_price_' + contract.toUpperCase();
  const cached = cache.get(cacheKey);
  if (cached) {
    return parseFloat(cached);
  }

  try {
    const url = `${API_BASE_URL}/futures/latest?contract=${encodeURIComponent(contract.toUpperCase())}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`API request failed: HTTP ${statusCode}`);
    }

    const data = JSON.parse(response.getContentText());

    if (data.data && data.data.contracts && data.data.contracts.length > 0) {
      const price = data.data.contracts[0].price;
      // Cache for 5 minutes (matches spot price TTL)
      cache.put(cacheKey, price.toString(), 300);
      return price;
    }

    throw new Error('No futures data found');
  } catch (error) {
    throw new Error(`Failed to fetch futures price: ${error.message}`);
  }
}

/**
 * Custom function: Get futures forward curve
 * Returns array of contract months and prices
 * @param {string} contract The futures contract code ("BZ" for Brent, "CL" for WTI)
 * @return {Array} Array of [month, price, change] rows
 * @customfunction
 */
function FUTURES_CURVE(contract) {
  if (!contract) {
    throw new Error('Contract code is required (BZ for Brent, CL for WTI)');
  }

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured. Use OilPriceAPI menu to configure.');
  }

  // Check cache first
  const cache = CacheService.getUserCache();
  const cacheKey = 'futures_curve_' + contract.toUpperCase();
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  try {
    const url = `${API_BASE_URL}/futures/curve?contract=${encodeURIComponent(contract.toUpperCase())}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Token ${apiKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`API request failed: HTTP ${statusCode}`);
    }

    const data = JSON.parse(response.getContentText());

    if (data.data && data.data.contracts && data.data.contracts.length > 0) {
      const result = [['Month', 'Price', 'Change']];
      data.data.contracts.forEach(c => {
        result.push([
          c.month,
          c.price,
          c.change !== undefined ? c.change : ''
        ]);
      });

      // Cache for 5 minutes (matches spot price TTL)
      cache.put(cacheKey, JSON.stringify(result), 300);
      return result;
    }

    throw new Error('No futures curve data found');
  } catch (error) {
    throw new Error(`Failed to fetch futures curve: ${error.message}`);
  }
}

/**
 * Show futures info dialog
 */
function showFuturesInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Futures Data Functions',
    'Use these custom functions in your spreadsheet:\n\n' +
    '=FUTURES_PRICE("BZ")\n  Get Brent front-month futures price\n\n' +
    '=FUTURES_PRICE("CL")\n  Get WTI front-month futures price\n\n' +
    '=FUTURES_CURVE("BZ")\n  Get full Brent forward curve\n\n' +
    'Note: Requires Reservoir Mastery subscription.',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// RIG COUNT DATA
// ============================================================================

/**
 * Custom function: Get latest rig count data
 * @param {string} type The rig type to return: "oil", "gas", "total", or "all"
 * @return {number|Array} Rig count number, or array if type is "all"
 * @customfunction
 */
function RIG_COUNT(type) {
  type = (type || 'total').toLowerCase();

  const apiKey = getApiKey();
  if (!apiKey) {
    throw new Error('API key not configured. Use OilPriceAPI menu to configure.');
  }

  // Check cache first
  const cache = CacheService.getUserCache();
  const cacheKey = 'rig_count_data';
  const cached = cache.get(cacheKey);
  let rigData;

  if (cached) {
    rigData = JSON.parse(cached);
  } else {
    try {
      const url = `${API_BASE_URL}/rig-counts/latest`;
      const options = {
        method: 'get',
        headers: {
          'Authorization': `Token ${apiKey}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode !== 200) {
        throw new Error(`API request failed: HTTP ${statusCode}`);
      }

      const data = JSON.parse(response.getContentText());

      if (!data.data) {
        throw new Error('No rig count data found');
      }

      rigData = data.data;
      // Cache for 1 hour
      cache.put(cacheKey, JSON.stringify(rigData), 3600);
    } catch (error) {
      throw new Error(`Failed to fetch rig counts: ${error.message}`);
    }
  }

  switch (type) {
    case 'oil':
      return rigData.oil;
    case 'gas':
      return rigData.gas;
    case 'total':
      return rigData.total;
    case 'all':
      return [
        ['Type', 'Count'],
        ['Oil', rigData.oil],
        ['Gas', rigData.gas],
        ['Total', rigData.total],
        ['Date', rigData.date]
      ];
    default:
      throw new Error('Invalid type. Use "oil", "gas", "total", or "all"');
  }
}

/**
 * Show rig count info dialog
 */
function showRigCountInfo() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Rig Count Functions',
    'Use these custom functions in your spreadsheet:\n\n' +
    '=RIG_COUNT("oil")\n  Get oil rig count\n\n' +
    '=RIG_COUNT("gas")\n  Get gas rig count\n\n' +
    '=RIG_COUNT("total")\n  Get total rig count\n\n' +
    '=RIG_COUNT("all")\n  Get full rig count breakdown',
    ui.ButtonSet.OK
  );
}
