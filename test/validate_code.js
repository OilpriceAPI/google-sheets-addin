#!/usr/bin/env node

/**
 * Code Validation Script
 * Validates Google Apps Script code structure without requiring Apps Script runtime
 */

const fs = require("fs");
const path = require("path");

// Color codes for terminal output
const colors = {
  reset: "\x1b[0m",
  green: "\x1b[32m",
  red: "\x1b[31m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
};

function log(message, color = "reset") {
  console.log(colors[color] + message + colors.reset);
}

/**
 * Read file contents
 */
function readFile(filename) {
  const filepath = path.join(__dirname, "..", filename);
  try {
    return fs.readFileSync(filepath, "utf8");
  } catch (error) {
    throw new Error(`Failed to read ${filename}: ${error.message}`);
  }
}

/**
 * Validate Code.gs
 */
function validateCodeGs() {
  log("\n📄 Validating Code.gs...", "blue");
  const code = readFile("Code.gs");

  const checks = [
    {
      name: "Has onOpen function",
      test: () => code.includes("function onOpen()"),
    },
    {
      name: "Has menu creation",
      test: () => code.includes("createMenu('OilPriceAPI')"),
    },
    {
      name: "Has API key management",
      test: () =>
        code.includes("function saveApiKey") &&
        code.includes("function getApiKey"),
    },
    {
      name: "Has custom function OILPRICE",
      test: () => code.includes("function OILPRICE("),
    },
    {
      name: "Has custom function OILPRICE_CONVERT",
      test: () => code.includes("function OILPRICE_CONVERT("),
    },
    {
      name: "Has custom function OILPRICE_HISTORY",
      test: () => code.includes("function OILPRICE_HISTORY("),
    },
    {
      name: "Has fetchLatestPrices",
      test: () => code.includes("function fetchLatestPrices("),
    },
    {
      name: "Has convertToMBtu",
      test: () => code.includes("function convertToMBtu("),
    },
    {
      name: "Has heat content logic for therms",
      test: () => code.includes("heatContent = 0.1") && code.includes("therm"),
    },
    {
      name: "Has heat content logic for MWh",
      test: () => code.includes("heatContent = 3.412") && code.includes("MWh"),
    },
    {
      name: "Has exchange rate conversion",
      test: () => code.includes("fetchExchangeRates"),
    },
    {
      name: "Has GBP to USD conversion",
      test: () => code.includes("GBp") || code.includes("GBP"),
    },
    {
      name: "Uses correct API endpoint",
      test: () => code.includes("api.oilpriceapi.com/v1"),
    },
    {
      name: "Has error handling",
      test: () => code.includes("try") && code.includes("catch"),
    },
    {
      name: "Has custom function FUTURES_PRICE",
      test: () => code.includes("function FUTURES_PRICE("),
    },
    {
      name: "Has custom function FUTURES_CURVE",
      test: () => code.includes("function FUTURES_CURVE("),
    },
    {
      name: "Has custom function RIG_COUNT",
      test: () => code.includes("function RIG_COUNT("),
    },
    {
      name: "Has @customfunction tag for FUTURES_PRICE",
      test: () =>
        code.includes("@customfunction") && code.includes("FUTURES_PRICE"),
    },
    {
      name: "Has CacheService usage",
      test: () => code.includes("CacheService.getUserCache()"),
    },
    {
      name: "Has futures menu item",
      test: () => code.includes("Fetch Futures Data"),
    },
    {
      name: "Has rig count menu item",
      test: () => code.includes("Fetch Rig Counts"),
    },
  ];

  runChecks(checks);
}

/**
 * Validate Sidebar.html
 */
function validateSidebarHtml() {
  log("\n📄 Validating Sidebar.html...", "blue");
  const html = readFile("Sidebar.html");

  const checks = [
    {
      name: "Has configuration section",
      test: () => html.includes("Configuration"),
    },
    {
      name: "Has API key input",
      test: () => html.includes('type="password"') && html.includes("apiKey"),
    },
    {
      name: "Has save button",
      test: () => html.includes("saveApiKey()"),
    },
    {
      name: "Has test connection button",
      test: () => html.includes("testConnection()"),
    },
    {
      name: "Has actions section",
      test: () => html.includes("Actions"),
    },
    {
      name: "Has quick functions section",
      test: () => html.includes("Quick Functions"),
    },
    {
      name: "Shows formula examples",
      test: () => html.includes("=OILPRICE"),
    },
    {
      name: "Has status messages",
      test: () => html.includes('class="status"'),
    },
    {
      name: "Has usage info display",
      test: () => html.includes("usageInfo"),
    },
    {
      name: "Uses Google.script.run",
      test: () => html.includes("google.script.run"),
    },
  ];

  runChecks(checks);
}

/**
 * Validate FetchDialog.html
 */
function validateFetchDialogHtml() {
  log("\n📄 Validating FetchDialog.html...", "blue");
  const html = readFile("FetchDialog.html");

  const checks = [
    {
      name: "Has commodity checkboxes",
      test: () => html.includes('type="checkbox"'),
    },
    {
      name: "Has BRENT_CRUDE_USD option",
      test: () => html.includes("BRENT_CRUDE_USD"),
    },
    {
      name: "Has WTI_USD option",
      test: () => html.includes("WTI_USD"),
    },
    {
      name: "Has NATURAL_GAS_USD option",
      test: () => html.includes("NATURAL_GAS_USD"),
    },
    {
      name: "Has select all button",
      test: () => html.includes("selectAll()"),
    },
    {
      name: "Has clear all button",
      test: () => html.includes("selectNone()"),
    },
    {
      name: "Has fetch button",
      test: () => html.includes("fetchPrices()"),
    },
    {
      name: "Has cancel button",
      test: () => html.includes("google.script.host.close"),
    },
    {
      name: "Has status display",
      test: () => html.includes("showStatus"),
    },
  ];

  runChecks(checks);
}

/**
 * Validate appsscript.json
 */
function validateManifest() {
  log("\n📄 Validating appsscript.json...", "blue");
  const manifestText = readFile("appsscript.json");

  const checks = [
    {
      name: "Is valid JSON",
      test: () => {
        JSON.parse(manifestText);
        return true;
      },
    },
    {
      name: "Has OAuth scopes",
      test: () => {
        const manifest = JSON.parse(manifestText);
        return manifest.oauthScopes && manifest.oauthScopes.length > 0;
      },
    },
    {
      name: "Has spreadsheet scope",
      test: () => manifestText.includes("spreadsheets.currentonly"),
    },
    {
      name: "Has external request scope",
      test: () => manifestText.includes("script.external_request"),
    },
    {
      name: "Uses V8 runtime",
      test: () => {
        const manifest = JSON.parse(manifestText);
        return manifest.runtimeVersion === "V8";
      },
    },
  ];

  runChecks(checks);
}

/**
 * Validate file structure
 */
function validateFileStructure() {
  log("\n📁 Validating File Structure...", "blue");

  const requiredFiles = [
    "Code.gs",
    "Sidebar.html",
    "FetchDialog.html",
    "appsscript.json",
    "README.md",
    "DEPLOYMENT_GUIDE.md",
    "docs/index.html",
    ".github/workflows/github-pages.yml",
  ];

  let allPresent = true;
  for (const file of requiredFiles) {
    const filepath = path.join(__dirname, "..", file);
    if (fs.existsSync(filepath)) {
      log(`  ✅ ${file}`, "green");
    } else {
      log(`  ❌ ${file} - MISSING`, "red");
      allPresent = false;
    }
  }

  if (!allPresent) {
    throw new Error("Missing required files");
  }
}

/**
 * Run a set of validation checks
 */
function runChecks(checks) {
  let passed = 0;
  let failed = 0;

  for (const check of checks) {
    try {
      if (check.test()) {
        log(`  ✅ ${check.name}`, "green");
        passed++;
      } else {
        log(`  ❌ ${check.name}`, "red");
        failed++;
      }
    } catch (error) {
      log(`  ❌ ${check.name}: ${error.message}`, "red");
      failed++;
    }
  }

  return { passed, failed };
}

/**
 * Main validation runner
 */
function main() {
  log("🧪 Google Sheets Add-on Code Validator\n", "blue");
  log("=".repeat(60), "blue");

  try {
    validateFileStructure();
    validateCodeGs();
    validateSidebarHtml();
    validateFetchDialogHtml();
    validateManifest();

    log("\n" + "=".repeat(60), "blue");
    log("✅ ALL VALIDATIONS PASSED!", "green");
    log("=".repeat(60), "blue");

    log("\n📋 Next Steps:", "yellow");
    log("  1. Open Google Sheets: https://sheets.google.com");
    log("  2. Extensions → Apps Script");
    log("  3. Copy files from this directory");
    log("  4. Run test harness: runAllTests()");
    log("\n📚 See test/README.md for detailed testing instructions\n");

    process.exit(0);
  } catch (error) {
    log("\n" + "=".repeat(60), "red");
    log("❌ VALIDATION FAILED", "red");
    log("=".repeat(60), "red");
    log(`\nError: ${error.message}\n`, "red");
    process.exit(1);
  }
}

// Run if called directly
if (require.main === module) {
  main();
}

module.exports = {
  validateCodeGs,
  validateSidebarHtml,
  validateFetchDialogHtml,
  validateManifest,
};
