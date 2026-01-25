# Google Sheets Add-on Deployment Guide

**Status:** ✅ Ready for Development Testing
**Target:** Google Workspace Marketplace Launch

---

## 📋 Project Structure

```
google-sheets-energy-addin/
├── Code.gs                   # Main Apps Script code
├── Sidebar.html              # Configuration sidebar UI
├── FetchDialog.html          # Commodity selection dialog
├── appsscript.json           # Apps Script manifest
├── README.md                 # Project documentation
├── DEPLOYMENT_GUIDE.md       # This file
├── .github/
│   └── workflows/
│       └── github-pages.yml  # Auto-deploy docs to GitHub Pages
└── docs/
    └── index.html            # Landing page for users
```

---

## 🚀 Development Testing (Do This First)

### Step 1: Create New Apps Script Project

1. **Open Google Sheets:**
   - Go to https://sheets.google.com
   - Create a new blank spreadsheet

2. **Open Apps Script Editor:**
   - Extensions → Apps Script
   - You'll see a new editor window

3. **Copy Code Files:**

   **File 1: Code.gs**
   - Delete default `function myFunction() {}`
   - Copy entire contents of `Code.gs`
   - Paste into the editor

   **File 2: Sidebar.html**
   - Click + next to Files
   - Select "HTML"
   - Name it: `Sidebar`
   - Copy contents of `Sidebar.html`
   - Paste into the editor

   **File 3: FetchDialog.html**
   - Click + next to Files
   - Select "HTML"
   - Name it: `FetchDialog`
   - Copy contents of `FetchDialog.html`
   - Paste into the editor

   **File 4: appsscript.json**
   - Click on existing `appsscript.json` in sidebar
   - Replace contents with our `appsscript.json`

4. **Save the Project:**
   - Click disk icon or Ctrl+S
   - Name it: "OilPriceAPI Energy Prices"

5. **Test the Add-on:**
   - Go back to your Google Sheet
   - Refresh the page (Ctrl+R or Cmd+R)
   - You should see "OilPriceAPI" menu appear

---

### Step 2: Test All Features

**Test 1: Configure API Key**
1. Click OilPriceAPI → Configure API Key
2. Sidebar opens on right
3. Enter test key: `3839c085460dd3a9dac1291f937f5a6d1740e8c668c766bc9f95e166af59cb11`
4. Click Save → Should see success message
5. Click Test Connection → Should see "Connection successful! ✓"

**Test 2: Fetch Latest Prices**
1. Click OilPriceAPI → Fetch Latest Prices
2. Dialog opens with commodity list
3. Select a few commodities (Brent, WTI, Natural Gas)
4. Click Fetch Prices
5. Should see new "Data" sheet with price table

**Test 3: Convert to $/MMBtu**
1. Click OilPriceAPI → Convert to $/MMBtu
2. Should see new "Process" sheet
3. All prices standardized to $/MMBtu
4. Verify UK Natural Gas shows heat content = 0.1

**Test 4: Custom Functions**
1. In any cell, type: `=OILPRICE("BRENT_CRUDE_USD")`
2. Press Enter
3. Should see latest Brent price (e.g., 71.45)
4. Try: `=OILPRICE_CONVERT("WTI_USD")`
5. Should see price in $/MMBtu

**All tests pass?** ✅ Ready for deployment!

---

## 📦 Production Deployment

### Step 1: Create Cloud Project

1. **Enable Google Cloud Platform:**
   - In Apps Script editor: Project Settings (gear icon)
   - Click "Change project"
   - Click "Create GCP project"
   - Name: "OilPriceAPI Sheets Add-on"

2. **Configure OAuth Screen:**
   - Go to https://console.cloud.google.com
   - APIs & Services → OAuth consent screen
   - User Type: External
   - App name: "OilPriceAPI Energy Prices"
   - Support email: support@oilpriceapi.com
   - Developer contact: support@oilpriceapi.com
   - Scopes: Add required scopes (already in appsscript.json)

---

### Step 2: Deploy as Add-on

1. **In Apps Script Editor:**
   - Click Deploy → New deployment
   - Type: Add-on
   - Configuration:
     - Name: "OilPriceAPI Energy Prices"
     - Description: "Real-time oil & commodity price data"
     - Version: v1.0.0

2. **Set Deployment Settings:**
   - Execute as: User accessing the add-on
   - Who has access: Only myself (for testing)

3. **Click Deploy**
   - Copy deployment ID (you'll need this)

---

### Step 3: Test Installed Add-on

1. **Open a NEW Google Sheet**
   - Must be different from development sheet
   - Refresh the page

2. **Verify Menu Appears:**
   - Should see "OilPriceAPI" menu
   - Click to test all features again

3. **Test with Different API Keys:**
   - Free tier key (100 requests (lifetime))
   - Paid tier key (historical data access)
   - Verify tier detection works

---

### Step 4: Prepare Marketplace Submission

**Required Materials:**

1. **Screenshots (5 required):**
   - Size: 1280x800 pixels
   - Format: PNG or JPEG
   - Screenshots needed:
     - ✅ Main menu showing OilPriceAPI options
     - ✅ Configuration sidebar with API key input
     - ✅ Fetch dialog with commodity selection
     - ✅ Data sheet with fetched prices
     - ✅ Process sheet with $/MMBtu conversions

2. **Demo Video (Optional but Recommended):**
   - Length: 60-90 seconds
   - Format: MP4 or YouTube link
   - Show: Install → Configure → Fetch → Convert workflow

3. **Privacy Policy:**
   - URL: https://oilpriceapi.com/privacy
   - Must mention data collection (API keys stored in user properties)
   - Must mention external API calls

4. **Terms of Service:**
   - URL: https://oilpriceapi.com/terms

5. **Support URL:**
   - URL: https://www.oilpriceapi.com/tools/sheets-support

---

### Step 5: Submit to Google Workspace Marketplace

1. **Go to Google Workspace Marketplace:**
   - https://console.cloud.google.com/marketplace
   - Click "Publish"

2. **Fill Out Listing:**

   **Basic Info:**
   - Name: "OilPriceAPI Energy Prices"
   - Tagline: "Real-time oil & commodity price data"
   - Detailed description:
     ```
     Get live commodity prices directly in Google Sheets.

     Features:
     - Custom functions: =OILPRICE("BRENT_CRUDE_USD")
     - Sidebar UI for easy price fetching
     - Auto-convert to $/MMBtu for comparison
     - 20+ commodities (Brent, WTI, Natural Gas, Coal, etc.)
     - 20 years of historical data (paid tiers)

     Perfect for:
     - Energy analysts
     - Commodity traders
     - Financial researchers
     - Market analysts

     Free tier: 100 requests (lifetime)
     No credit card required to start
     ```

   **Category:** Productivity → Data Management
   **Language:** English (United States)

3. **Upload Screenshots & Video:**
   - Upload all 5 screenshots
   - Add video URL if available

4. **Add URLs:**
   - Website: https://oilpriceapi.com
   - Support: https://www.oilpriceapi.com/tools/sheets-support
   - Privacy: https://oilpriceapi.com/privacy
   - Terms: https://oilpriceapi.com/terms

5. **Submit for Review:**
   - Click "Submit for review"
   - Wait 5-10 days for Google approval

---

## 🎯 Post-Launch Checklist

### Week 1 After Approval

- [ ] Verify listing is live on Workspace Marketplace
- [ ] Test installation from Marketplace (not manual)
- [ ] Monitor error logs in Apps Script dashboard
- [ ] Check API usage spikes
- [ ] Respond to any user reviews

### Week 2-4

- [ ] Send email campaign to existing users
- [ ] Post to r/googlesheets on Reddit
- [ ] Share on Twitter/X with demo
- [ ] Update main website with Sheets add-on link
- [ ] Create tutorial blog post

### Month 2-3

- [ ] Add more commodities based on user requests
- [ ] Implement user-requested features
- [ ] Create video tutorial for YouTube
- [ ] Build dashboard templates for common use cases

---

## 🐛 Troubleshooting

### Issue: Menu doesn't appear after refresh
**Solution:**
- Clear Google Sheets cache
- Close all sheets and reopen
- Check Apps Script triggers are enabled

### Issue: API calls failing
**Solution:**
- Verify API key is correct
- Check `UrlFetchApp` is allowed in scopes
- Look at Apps Script execution logs

### Issue: Custom functions not working
**Solution:**
- Functions must start with `=` in cell
- Commodity code must be in quotes: `=OILPRICE("BRENT_CRUDE_USD")`
- Refresh sheet after saving Code.gs changes

### Issue: "Authorization required" error
**Solution:**
- Click "Review permissions"
- Allow access to spreadsheets and external requests
- This is normal for first-time setup

---

## 📊 Success Metrics

### Week 1 Targets:
- 100+ installs from Marketplace
- 50+ active users (configured API key)
- <5% error rate

### Month 1 Targets:
- 500+ installs
- 250+ active users
- 25+ paid conversions ($375 MRR)

### Month 3 Targets:
- 2,000+ installs
- 1,000+ active users
- 100+ paid conversions ($1,500 MRR)

---

## 🔗 Related Links

- **GitHub Repository:** https://github.com/OilpriceAPI/google-sheets-addin
- **Landing Page:** https://oilpriceapi.github.io/google-sheets-addin/
- **API Documentation:** https://docs.oilpriceapi.com
- **Support:** support@oilpriceapi.com

---

## 📝 Version History

**v1.0.0 (Planned - 2026-01-15)**
- Initial release
- Custom functions (OILPRICE, OILPRICE_HISTORY, OILPRICE_CONVERT)
- Sidebar UI for configuration
- Fetch dialog for multiple commodities
- Auto-convert to $/MMBtu
- Tier detection

---

**Current Status:** ✅ Code complete, ready for testing
**Next Step:** Manual testing in Apps Script
**Target Launch:** January 2026
