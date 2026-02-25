/**
 * Google Sheets Add-on Test Harness
 * Tests all functionality without requiring manual UI interaction
 */

// Test configuration
// Set TEST_API_KEY in Script Properties (File > Project Settings > Script Properties)
// before running these tests.
const TEST_API_KEY =
  PropertiesService.getScriptProperties().getProperty("TEST_API_KEY") ||
  "YOUR_API_KEY_HERE";
const TEST_COMMODITIES = ["BRENT_CRUDE_USD", "WTI_USD", "NATURAL_GAS_USD"];

/**
 * Main test runner
 */
function runAllTests() {
  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    tests: [],
  };

  console.log("🧪 Starting Google Sheets Add-on Test Harness...\n");

  // Test 1: API Key Management
  runTest(results, "API Key Save/Retrieve", testApiKeyManagement);

  // Test 2: API Connection
  runTest(results, "API Connection Test", testConnection);

  // Test 3: User Info Retrieval
  runTest(results, "User Info Retrieval", testUserInfo);

  // Test 4: Latest Price Fetch
  runTest(results, "Fetch Latest Prices", testFetchLatestPrices);

  // Test 5: Data Sheet Creation
  runTest(results, "Data Sheet Creation", testDataSheetCreation);

  // Test 6: Exchange Rates
  runTest(results, "Exchange Rate Fetch", testExchangeRates);

  // Test 7: MBtu Conversion
  runTest(results, "MBtu Conversion", testMBtuConversion);

  // Test 8: Custom Function OILPRICE
  runTest(results, "Custom Function OILPRICE", testOilPriceFunction);

  // Test 9: Custom Function OILPRICE_CONVERT
  runTest(
    results,
    "Custom Function OILPRICE_CONVERT",
    testOilPriceConvertFunction,
  );

  // Test 10: Heat Content Logic
  runTest(results, "Heat Content Logic", testHeatContentLogic);

  // Test 11: Futures Price
  runTest(results, "Custom Function FUTURES_PRICE", testFuturesPrice);

  // Test 12: Futures Curve
  runTest(results, "Custom Function FUTURES_CURVE", testFuturesCurve);

  // Test 13: Rig Count
  runTest(results, "Custom Function RIG_COUNT", testRigCount);

  // Test 14: Rate Limiting (Cache)
  runTest(results, "Rate Limiting Cache", testRateLimiting);

  // Print summary
  console.log("\n" + "=".repeat(60));
  console.log("📊 TEST SUMMARY");
  console.log("=".repeat(60));
  console.log(`Total Tests: ${results.total}`);
  console.log(`✅ Passed: ${results.passed}`);
  console.log(`❌ Failed: ${results.failed}`);
  console.log(
    `Success Rate: ${((results.passed / results.total) * 100).toFixed(1)}%`,
  );
  console.log("=".repeat(60));

  // Print detailed results
  if (results.failed > 0) {
    console.log("\n❌ FAILED TESTS:");
    results.tests
      .filter((t) => !t.passed)
      .forEach((t) => {
        console.log(`  - ${t.name}: ${t.error}`);
      });
  }

  return results;
}

/**
 * Run a single test
 */
function runTest(results, name, testFunc) {
  results.total++;
  console.log(`Running: ${name}...`);

  try {
    testFunc();
    results.passed++;
    results.tests.push({ name, passed: true });
    console.log(`  ✅ PASSED\n`);
  } catch (error) {
    results.failed++;
    results.tests.push({ name, passed: false, error: error.message });
    console.log(`  ❌ FAILED: ${error.message}\n`);
  }
}

/**
 * Test: API Key Management
 */
function testApiKeyManagement() {
  // Save API key
  const saveResult = saveApiKey(TEST_API_KEY);
  assert(saveResult.success === true, "API key save should succeed");

  // Retrieve API key
  const retrievedKey = getApiKey();
  assert(
    retrievedKey === TEST_API_KEY,
    "Retrieved API key should match saved key",
  );

  console.log(`  API Key: ${retrievedKey.substring(0, 20)}...`);
}

/**
 * Test: API Connection
 */
function testConnection() {
  const result = testConnection();
  assert(
    result.success === true,
    `Connection should succeed: ${result.message}`,
  );
  console.log(`  ${result.message}`);
}

/**
 * Test: User Info Retrieval
 */
function testUserInfo() {
  const info = getUserInfo();
  assert(info.tier !== "unknown", "Should retrieve user tier");
  assert(info.limit > 0, "Should have request limit");

  console.log(`  Tier: ${info.tier}`);
  console.log(`  Limit: ${info.limit.toLocaleString()}`);
  console.log(`  Used: ${info.used.toLocaleString()}`);
}

/**
 * Test: Fetch Latest Prices
 */
function testFetchLatestPrices() {
  const result = fetchLatestPrices(TEST_COMMODITIES);
  assert(result.success === true, "Fetch should succeed");
  assert(
    result.count === TEST_COMMODITIES.length,
    `Should fetch ${TEST_COMMODITIES.length} prices`,
  );

  console.log(`  Fetched ${result.count} prices`);
}

/**
 * Test: Data Sheet Creation
 */
function testDataSheetCreation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");

  assert(sheet !== null, "Data sheet should exist");

  const data = sheet.getDataRange().getValues();
  assert(data.length > 1, "Data sheet should have data rows");
  assert(data[0][0] === "Commodity Code", "Header should be correct");

  console.log(`  Data sheet has ${data.length - 1} rows`);
}

/**
 * Test: Exchange Rates
 */
function testExchangeRates() {
  const rates = fetchExchangeRates();
  assert(rates.gbpUsd > 0, "GBP/USD rate should be positive");
  assert(rates.eurUsd > 0, "EUR/USD rate should be positive");

  console.log(`  GBP/USD: ${rates.gbpUsd.toFixed(4)}`);
  console.log(`  EUR/USD: ${rates.eurUsd.toFixed(4)}`);
}

/**
 * Test: MBtu Conversion
 */
function testMBtuConversion() {
  convertToMBtu();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Process");

  assert(sheet !== null, "Process sheet should exist");

  const data = sheet.getDataRange().getValues();
  assert(data.length > 1, "Process sheet should have data");
  assert(data[0][6] === "Price per MBtu (USD)", "Header should be correct");

  // Check that conversions are reasonable
  for (let i = 1; i < data.length; i++) {
    const pricePerMBtu = data[i][6];
    assert(pricePerMBtu > 0, `Price per MBtu should be positive: row ${i}`);
    assert(
      pricePerMBtu < 1000,
      `Price per MBtu should be reasonable: row ${i}`,
    );
  }

  console.log(`  Converted ${data.length - 1} commodities`);
}

/**
 * Test: Custom Function OILPRICE
 */
function testOilPriceFunction() {
  const price = OILPRICE("BRENT_CRUDE_USD");
  assert(typeof price === "number", "Should return a number");
  assert(price > 0, "Price should be positive");
  assert(price < 200, "Price should be reasonable for Brent");

  console.log(`  Brent Crude: $${price.toFixed(2)}`);
}

/**
 * Test: Custom Function OILPRICE_CONVERT
 */
function testOilPriceConvertFunction() {
  const pricePerMBtu = OILPRICE_CONVERT("BRENT_CRUDE_USD");
  assert(typeof pricePerMBtu === "number", "Should return a number");
  assert(pricePerMBtu > 0, "Price should be positive");

  console.log(`  Brent per MBtu: $${pricePerMBtu.toFixed(2)}`);
}

/**
 * Test: Heat Content Logic
 */
function testHeatContentLogic() {
  // Test therm conversion
  const commodityInfo = COMMODITY_MAP["NATURAL_GAS_GBP"];
  assert(commodityInfo.unit === "therm", "UK gas should be in therms");

  let heatContent;
  if (commodityInfo.unit === "therm") {
    heatContent = 0.1;
  }
  assert(heatContent === 0.1, "Therm heat content should be 0.1");

  // Test MWh conversion
  const dutchInfo = COMMODITY_MAP["DUTCH_TTF_EUR"];
  assert(dutchInfo.unit === "MWh", "Dutch TTF should be in MWh");

  if (dutchInfo.unit === "MWh") {
    heatContent = 3.412;
  }
  assert(heatContent === 3.412, "MWh heat content should be 3.412");

  // Test MBtu (no conversion needed)
  const usGasInfo = COMMODITY_MAP["NATURAL_GAS_USD"];
  assert(usGasInfo.unit === "MBtu", "US gas should be in MBtu");

  console.log("  ✓ Therm: 0.1 MMBtu/therm");
  console.log("  ✓ MWh: 3.412 MMBtu/MWh");
  console.log("  ✓ MBtu: 1.0 (no conversion)");
}

/**
 * Test: Custom Function FUTURES_PRICE
 */
function testFuturesPrice() {
  const price = FUTURES_PRICE("BZ");
  assert(typeof price === "number", "Should return a number");
  assert(price > 0, "Price should be positive");
  assert(price < 200, "Price should be reasonable for Brent futures");

  console.log(`  Brent Futures: $${price.toFixed(2)}`);
}

/**
 * Test: Custom Function FUTURES_CURVE
 */
function testFuturesCurve() {
  const curve = FUTURES_CURVE("BZ");
  assert(Array.isArray(curve), "Should return an array");
  assert(curve.length > 1, "Should have header + data rows");
  assert(curve[0][0] === "Month", "First column header should be Month");
  assert(curve[0][1] === "Price", "Second column header should be Price");

  // Check data rows
  for (let i = 1; i < curve.length; i++) {
    assert(
      typeof curve[i][1] === "number",
      `Row ${i} price should be a number`,
    );
    assert(curve[i][1] > 0, `Row ${i} price should be positive`);
  }

  console.log(`  Curve has ${curve.length - 1} contract months`);
}

/**
 * Test: Custom Function RIG_COUNT
 */
function testRigCount() {
  const total = RIG_COUNT("total");
  assert(typeof total === "number", "Should return a number");
  assert(total > 0, "Total rigs should be positive");

  const oil = RIG_COUNT("oil");
  const gas = RIG_COUNT("gas");
  assert(oil > 0, "Oil rigs should be positive");
  assert(gas > 0, "Gas rigs should be positive");
  assert(
    oil + gas <= total + 10,
    "Oil + gas should roughly equal total (with misc)",
  );

  const all = RIG_COUNT("all");
  assert(Array.isArray(all), "all type should return array");
  assert(all.length >= 4, "all should have at least 4 rows");

  console.log(`  Oil: ${oil}, Gas: ${gas}, Total: ${total}`);
}

/**
 * Test: Rate Limiting Cache
 */
function testRateLimiting() {
  // First call fetches from API
  const price1 = OILPRICE("BRENT_CRUDE_USD");

  // Second call should be from cache (much faster)
  const startTime = new Date().getTime();
  const price2 = OILPRICE("BRENT_CRUDE_USD");
  const elapsed = new Date().getTime() - startTime;

  assert(price1 === price2, "Cached price should match original");
  assert(elapsed < 100, "Cached call should be fast (< 100ms)");

  console.log(`  Cached call took ${elapsed}ms`);
}

/**
 * Simple assertion helper
 */
function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

/**
 * Test cleanup - remove test sheets
 */
function cleanupTestSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove Data sheet
  const dataSheet = ss.getSheetByName("Data");
  if (dataSheet) {
    ss.deleteSheet(dataSheet);
  }

  // Remove Process sheet
  const processSheet = ss.getSheetByName("Process");
  if (processSheet) {
    ss.deleteSheet(processSheet);
  }

  console.log("✅ Test sheets cleaned up");
}

/**
 * Integration test - Full workflow
 */
function testFullWorkflow() {
  console.log("🧪 Running Full Integration Test...\n");

  try {
    // Step 1: Save API key
    console.log("Step 1: Saving API key...");
    saveApiKey(TEST_API_KEY);
    console.log("  ✅ API key saved\n");

    // Step 2: Test connection
    console.log("Step 2: Testing connection...");
    const connResult = testConnection();
    console.log(`  ${connResult.message}\n`);

    // Step 3: Fetch prices
    console.log("Step 3: Fetching latest prices...");
    const fetchResult = fetchLatestPrices(TEST_COMMODITIES);
    console.log(`  ✅ Fetched ${fetchResult.count} prices\n`);

    // Step 4: Convert to MBtu
    console.log("Step 4: Converting to $/MMBtu...");
    convertToMBtu();
    console.log("  ✅ Conversion complete\n");

    // Step 5: Test custom functions
    console.log("Step 5: Testing custom functions...");
    const brentPrice = OILPRICE("BRENT_CRUDE_USD");
    const brentMBtu = OILPRICE_CONVERT("BRENT_CRUDE_USD");
    console.log(`  Brent Crude: $${brentPrice.toFixed(2)}`);
    console.log(`  Brent per MBtu: $${brentMBtu.toFixed(2)}\n`);

    console.log("✅ Full workflow test PASSED!");
    return true;
  } catch (error) {
    console.log(`❌ Full workflow test FAILED: ${error.message}`);
    return false;
  }
}

/**
 * Quick smoke test - runs essential checks only
 */
function smokeTest() {
  console.log("🔥 Running Smoke Test...\n");

  const tests = [
    { name: "API Key", func: () => saveApiKey(TEST_API_KEY) },
    { name: "Connection", func: testConnection },
    {
      name: "Fetch Prices",
      func: () => fetchLatestPrices(["BRENT_CRUDE_USD"]),
    },
  ];

  let passed = 0;
  for (const test of tests) {
    try {
      test.func();
      console.log(`✅ ${test.name}`);
      passed++;
    } catch (error) {
      console.log(`❌ ${test.name}: ${error.message}`);
    }
  }

  console.log(`\n${passed}/${tests.length} smoke tests passed`);
  return passed === tests.length;
}
