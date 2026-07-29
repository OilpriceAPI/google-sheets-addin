function normalizeGasMarkets_(henryHubUsdMmbtu, ttfEurMwh, jkmUsdMmbtu, eurUsd) {
  const values = [henryHubUsdMmbtu, ttfEurMwh, jkmUsdMmbtu, eurUsd].map(Number);
  if (!values.every(Number.isFinite) || values[3] <= 0) throw new Error('Gas-market inputs and FX must be finite positive values.');
  const ttfUsdMmbtu = (values[1] * values[3]) / 3.412141633;
  return {
    henryHub: values[0],
    ttf: ttfUsdMmbtu,
    jkm: values[2],
    ttfHenryHub: ttfUsdMmbtu - values[0],
    jkmHenryHub: values[2] - values[0],
    jkmTtf: values[2] - ttfUsdMmbtu
  };
}

function buildGasSpreadWorkbook() {
  const records = latestPrices_(OPA_PRODUCT.allowedCodes);
  const price = Object.fromEntries(records.map((record) => [record.code, record.price]));
  const normalized = normalizeGasMarkets_(
    price.NATURAL_GAS_USD,
    price.DUTCH_TTF_EUR,
    price.JKM_LNG_USD,
    price.EUR_USD
  );

  const data = sheet_('Market Data');
  writeTitle_(data, 'Global Gas Market Data', 'Native prices and units returned by OilPriceAPI.');
  writeTable_(data, 4, 1, marketDataRows_(records));

  const audit = sheet_('Conversion Audit');
  writeTitle_(audit, 'TTF Conversion Audit', 'TTF EUR/MWh × EUR/USD ÷ 3.412141633 MMBtu/MWh.');
  writeTable_(audit, 4, 1, [
    ['Input', 'Value', 'Unit'],
    ['TTF native', price.DUTCH_TTF_EUR, 'EUR/MWh'],
    ['EUR/USD', price.EUR_USD, 'USD per EUR'],
    ['Energy conversion', 3.412141633, 'MMBtu per MWh'],
    ['TTF normalized', normalized.ttf, 'USD/MMBtu']
  ]);
  audit.getRange('B5:B8').setNumberFormat('0.0000');

  const monitor = sheet_('Gas Spread Monitor');
  writeTitle_(monitor, 'Natural Gas Spread Monitor', 'Cross-market comparison after explicit currency and energy-unit normalization.');
  writeTable_(monitor, 4, 1, [
    ['Market', 'Normalized price', 'Unit'],
    ['Henry Hub', normalized.henryHub, 'USD/MMBtu'],
    ['Dutch TTF', normalized.ttf, 'USD/MMBtu'],
    ['JKM LNG', normalized.jkm, 'USD/MMBtu']
  ]);
  writeTable_(monitor, 10, 1, [
    ['Spread', 'Value', 'Unit'],
    ['TTF minus Henry Hub', normalized.ttfHenryHub, 'USD/MMBtu'],
    ['JKM minus Henry Hub', normalized.jkmHenryHub, 'USD/MMBtu'],
    ['JKM minus TTF', normalized.jkmTtf, 'USD/MMBtu']
  ]);
  monitor.getRange('B5:B7').setNumberFormat('$0.000');
  monitor.getRange('B11:B13').setNumberFormat('$0.000');

  activateProduct_();
  monitor.activate();
  return { success: true, message: 'Gas spread monitor built with Henry Hub, TTF, and JKM normalized to USD/MMBtu.' };
}
