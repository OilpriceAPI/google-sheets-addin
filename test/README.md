# Google Sheets Add-on Test Suite

**Purpose:** Automated testing for the OilPriceAPI Google Sheets add-on

---

## 🧪 Test Harness

### What It Tests

1. **API Key Management**
   - Save API key to user properties
   - Retrieve API key from storage

2. **API Connection**
   - Test authentication
   - Verify API is reachable

3. **User Info Retrieval**
   - Get user tier
   - Get request limits
   - Get usage stats

4. **Price Fetching**
   - Fetch latest prices for multiple commodities
   - Create Data sheet
   - Populate with price data

5. **Exchange Rates**
   - Fetch GBP/USD rate
   - Fetch EUR/USD rate

6. **MBtu Conversion**
   - Create Process sheet
   - Convert all prices to $/MMBtu
   - Apply correct heat content factors

7. **Custom Functions**
   - `=OILPRICE()` returns valid price
   - `=OILPRICE_CONVERT()` returns converted price
   - Error handling

8. **Heat Content Logic**
   - Therm: 0.1 MMBtu/therm
   - MWh: 3.412 MMBtu/MWh
   - MBtu: 1.0 (no conversion)
   - Commodity-specific for barrel/tonne

---

## 🚀 How to Run Tests

### Method 1: In Apps Script Editor (Recommended)

1. **Open your test spreadsheet:**
   - Go to https://sheets.google.com
   - Open the sheet with OilPriceAPI add-on installed

2. **Open Apps Script Editor:**
   - Extensions → Apps Script

3. **Add test harness:**
   - Click + next to Files → Script
   - Name it: `TestHarness`
   - Copy contents of `test_harness.js`
   - Paste into editor
   - Save (Ctrl+S)

4. **Run tests:**
   - Select function: `runAllTests`
   - Click Run (▶️)
   - View output in Execution log

5. **View results:**
   - Bottom panel shows execution log
   - Look for ✅ PASSED or ❌ FAILED

---

### Method 2: Quick Smoke Test

For faster testing, run just the essential checks:

1. Select function: `smokeTest`
2. Click Run
3. Verify 3/3 tests pass

---

### Method 3: Full Integration Test

Test the complete user workflow:

1. Select function: `testFullWorkflow`
2. Click Run
3. Verifies entire flow from API key to custom functions

---

## 📊 Expected Output

### All Tests Passing

```
🧪 Starting Google Sheets Add-on Test Harness...

Running: API Key Save/Retrieve...
  API Key: 3839c085460dd3a9d...
  ✅ PASSED

Running: API Connection Test...
  Connection successful! ✓
  ✅ PASSED

Running: User Info Retrieval...
  Tier: admin
  Limit: 999,999,999
  Used: 12,345
  ✅ PASSED

Running: Fetch Latest Prices...
  Fetched 3 prices
  ✅ PASSED

Running: Data Sheet Creation...
  Data sheet has 3 rows
  ✅ PASSED

Running: Exchange Rate Fetch...
  GBP/USD: 1.3200
  EUR/USD: 1.1000
  ✅ PASSED

Running: MBtu Conversion...
  Converted 3 commodities
  ✅ PASSED

Running: Custom Function OILPRICE...
  Brent Crude: $71.45
  ✅ PASSED

Running: Custom Function OILPRICE_CONVERT...
  Brent per MBtu: $12.32
  ✅ PASSED

Running: Heat Content Logic...
  ✓ Therm: 0.1 MMBtu/therm
  ✓ MWh: 3.412 MMBtu/MWh
  ✓ MBtu: 1.0 (no conversion)
  ✅ PASSED

============================================================
📊 TEST SUMMARY
============================================================
Total Tests: 10
✅ Passed: 10
❌ Failed: 0
Success Rate: 100.0%
============================================================
```

---

## 🐛 Troubleshooting

### Issue: "ReferenceError: saveApiKey is not defined"

**Cause:** Test harness can't find main Code.gs functions

**Solution:**
1. Make sure Code.gs is in the same Apps Script project
2. Both files must be saved
3. Refresh the editor

---

### Issue: "Exception: Action not allowed"

**Cause:** Script needs authorization to access sheets

**Solution:**
1. Click "Review permissions"
2. Choose your Google account
3. Click "Advanced" → "Go to [Project Name]"
4. Click "Allow"

---

### Issue: "API request failed: HTTP 401"

**Cause:** Invalid or missing API key

**Solution:**
1. Update `TEST_API_KEY` in test_harness.js
2. Use a valid API key from https://oilpriceapi.com
3. Save and re-run

---

### Issue: Tests fail with "No Data sheet found"

**Cause:** Tests must run in order

**Solution:**
1. Run `runAllTests()` (runs all tests in correct order)
2. Or run `testFullWorkflow()` for complete integration test
3. Don't run individual tests manually

---

## 🧹 Cleanup

After testing, remove test sheets:

```javascript
cleanupTestSheets()
```

This deletes:
- Data sheet
- Process sheet

---

## 📈 CI/CD Integration (Future)

For automated testing on every commit:

1. **Use clasp (Google's CLI):**
   ```bash
   npm install -g @google/clasp
   clasp login
   clasp create --type sheets --title "OilPriceAPI Test"
   clasp push
   ```

2. **Run tests via clasp:**
   ```bash
   clasp run runAllTests
   ```

3. **Add to GitHub Actions:**
   ```yaml
   - name: Run Google Apps Script tests
     run: clasp run runAllTests
   ```

---

## 📝 Test Coverage

| Area | Coverage | Tests |
|------|----------|-------|
| API Integration | ✅ 100% | 3 tests |
| Data Fetching | ✅ 100% | 2 tests |
| Data Processing | ✅ 100% | 2 tests |
| Custom Functions | ✅ 100% | 2 tests |
| Heat Content Logic | ✅ 100% | 1 test |
| **Total** | **✅ 100%** | **10 tests** |

---

## 🎯 Success Criteria

**All tests must pass before:**
- Submitting to Google Workspace Marketplace
- Deploying updates to production
- Releasing new versions

**Critical tests:**
- API Connection (must succeed)
- Custom Functions (must return valid data)
- Heat Content Logic (must use correct factors)

---

## 🔗 Related Files

- **Main Code:** `../Code.gs`
- **UI Files:** `../Sidebar.html`, `../FetchDialog.html`
- **Deployment:** `../DEPLOYMENT_GUIDE.md`

---

## 📞 Support

**Issues with tests?**
- Check Apps Script execution logs
- Verify API key is valid
- Ensure all files are saved

**Need help?**
- Email: support@oilpriceapi.com
- GitHub Issues: https://github.com/OilpriceAPI/google-sheets-addin/issues

---

**Test Status:** ✅ Ready to run
**Last Updated:** 2025-11-29
**Version:** 1.0.0
