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
      return data.data.prices[0].price;
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

  try {
    const url = `${API_BASE_URL}/prices/past_day?by_code=${commodityCode}&days=${days}`;
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
      return data.data.prices.map(p => [p.created_at, p.price]);
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
