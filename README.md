# OilPriceAPI Google Sheets Add-on

**Real-time oil & commodity price data directly in Google Sheets**

## 🎯 Overview

This Google Sheets add-on brings live commodity prices to your spreadsheets:
- **Custom Functions**: `=OILPRICE("BRENT_CRUDE_USD")`
- **Sidebar UI**: Fetch and manage prices easily
- **Auto-Convert**: Standardize all prices to $/MMBtu
- **Historical Data**: Access 20+ years of daily prices

## 📦 Features

### Custom Functions
```javascript
=OILPRICE("BRENT_CRUDE_USD")           // Returns latest Brent Crude price
=OILPRICE_HISTORY("WTI_USD", 30)       // Returns 30 days of WTI data
=OILPRICE_CONVERT("NATURAL_GAS_GBP")   // Returns price in $/MMBtu
```

### Sidebar Features
- API key management
- Commodity selection (20+ commodities)
- Fetch latest prices
- Convert to $/MMBtu
- Historical data charts
- Usage tracking

### Tier Detection
- Free: 1,000 requests/month
- Exploration: 10,000 requests/month + historical data
- Production+: Higher limits + webhooks

## 🚀 Installation

### From Google Workspace Marketplace (Coming Soon)
1. Open any Google Sheet
2. Extensions → Add-ons → Get add-ons
3. Search "OilPriceAPI Energy Prices"
4. Click Install

### For Development
1. Open Google Sheets
2. Extensions → Apps Script
3. Copy code from `Code.gs` and `Sidebar.html`
4. Save and refresh

## 📖 Quick Start

### Step 1: Get API Key
Visit https://oilpriceapi.com/auth/signup to get your free API key.

### Step 2: Open Add-on
1. Extensions → OilPriceAPI → Configure
2. Enter your API key
3. Click Save

### Step 3: Use Custom Functions
In any cell, type:
```
=OILPRICE("BRENT_CRUDE_USD")
```

### Step 4: Fetch Multiple Prices
1. Extensions → OilPriceAPI → Fetch Prices
2. Select commodities
3. Click Fetch
4. Data appears in "Data" sheet

## 🎨 Example Dashboards

### Energy Price Comparison
```
| Commodity       | Latest Price | $/MMBtu  | 30-Day Change |
|-----------------|--------------|----------|---------------|
| Brent Crude     | $71.45       | $12.32   | +2.3%         |
| WTI Crude       | $66.20       | $11.41   | +1.8%         |
| Natural Gas (US)| $4.84        | $4.84    | -3.2%         |
| UK Natural Gas  | £7.50        | $99.30   | +15.7%        |
```

### Historical Price Chart
1. Fetch past year data
2. Insert → Chart
3. Select date and price columns
4. Choose line chart

## 💰 Pricing

**Free Tier:**
- 1,000 requests/month
- Real-time prices
- 20+ commodities
- Custom functions

**Exploration ($15/mo):**
- 10,000 requests/month
- Historical data access
- Pandas integration (Python SDK)
- Email support

**Production ($45/mo):**
- 50,000 requests/month
- Webhooks
- Priority support
- 99.9% uptime SLA

[View Full Pricing](https://oilpriceapi.com/pricing)

## 🛠️ Development

### Project Structure
```
google-sheets-energy-addin/
├── Code.gs              # Main Apps Script code
├── Sidebar.html         # Sidebar UI
├── appsscript.json      # Project manifest
├── README.md            # This file
└── docs/
    └── index.html       # GitHub Pages landing page
```

### Testing Locally
1. Open Google Apps Script editor
2. Make changes to Code.gs or Sidebar.html
3. Click Run → test function
4. Debug in Apps Script console

### Deployment
1. Click Deploy → New deployment
2. Type: Add-on
3. Description: Version X.X.X
4. Click Deploy

## 📚 Documentation

- **API Reference**: https://docs.oilpriceapi.com
- **Custom Functions Guide**: https://www.oilpriceapi.com/tools/sheets-functions
- **Video Tutorials**: https://www.youtube.com/oilpriceapi
- **Support**: support@oilpriceapi.com

## 🐛 Known Issues

- None currently

## 🔒 Privacy & Security

- API keys stored securely in user properties
- No data shared with third parties
- HTTPS encryption for all API calls
- [Privacy Policy](https://oilpriceapi.com/privacy)
- [Terms of Service](https://oilpriceapi.com/terms)

## 📞 Support

**Having issues?**
- Email: support@oilpriceapi.com
- GitHub Issues: https://github.com/OilpriceAPI/google-sheets-addin/issues
- Documentation: https://www.oilpriceapi.com/tools/sheets-support

## 🎓 Examples

### Example 1: Build a Price Dashboard
1. Create new sheet named "Dashboard"
2. Add formulas:
   ```
   =OILPRICE("BRENT_CRUDE_USD")
   =OILPRICE("WTI_USD")
   =OILPRICE("NATURAL_GAS_USD")
   ```
3. Format as table
4. Add conditional formatting for price changes

### Example 2: Historical Analysis
1. Extensions → OilPriceAPI → Fetch Historical
2. Select "Brent Crude" and "Past Year"
3. Data appears in "Historical" sheet
4. Create line chart to visualize trends

### Example 3: Energy Cost Calculator
```
A1: Commodity       | B1: =OILPRICE("NATURAL_GAS_USD")
A2: Volume (MMBtu)  | B2: 1000
A3: Total Cost      | B3: =B1*B2
```

## 🏆 Why OilPriceAPI?

- **98% Less Cost**: Compared to Bloomberg Terminal
- **20+ Commodities**: Brent, WTI, Natural Gas, Coal, and more
- **20 Years of Data**: Daily historical prices since 2005
- **Production-Ready**: Used by analysts, traders, and developers worldwide
- **Free Tier**: 1,000 requests/month, no credit card required

## 📝 License

MIT License - See LICENSE file for details

## 🚀 Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create feature branch
3. Submit pull request

## 🔗 Related Tools

- **Python SDK**: https://github.com/OilpriceAPI/python-sdk
- **Node.js SDK**: https://github.com/OilpriceAPI/oilpriceapi-node
- **Excel Add-in**: https://github.com/OilpriceAPI/excel-energy-addin

---

**Built with ❤️ by OilPriceAPI**

[Website](https://oilpriceapi.com) • [Pricing](https://oilpriceapi.com/pricing) • [Docs](https://docs.oilpriceapi.com) • [Support](mailto:support@oilpriceapi.com)
