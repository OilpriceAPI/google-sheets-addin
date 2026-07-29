function calculateCrackSpread_(crudeUsdPerBarrel, gasolineUsdPerGallon, distillateUsdPerGallon, ratio) {
  const crude = Number(crudeUsdPerBarrel);
  const gasoline = Number(gasolineUsdPerGallon);
  const distillate = Number(distillateUsdPerGallon);
  if (![crude, gasoline, distillate].every(Number.isFinite)) throw new Error('Crack-spread inputs must be finite numbers.');
  const selected = ratio || '3-2-1';
  if (selected === '2-1-1') return ((gasoline * 42) + (distillate * 42) - (2 * crude)) / 2;
  if (selected !== '3-2-1') throw new Error('Supported crack-spread ratios are 3-2-1 and 2-1-1.');
  return ((2 * gasoline * 42) + (distillate * 42) - (3 * crude)) / 3;
}

function crackSensitivity_(crude, gasoline, distillate) {
  const rows = [['Crude change', 'Product change', '3-2-1 spread (USD/bbl)']];
  [-10, -5, 0, 5, 10].forEach((crudeDelta) => {
    [-0.20, -0.10, 0, 0.10, 0.20].forEach((productDelta) => {
      rows.push([
        crudeDelta,
        productDelta,
        calculateCrackSpread_(crude + crudeDelta, gasoline + productDelta, distillate + productDelta, '3-2-1')
      ]);
    });
  });
  return rows;
}

function buildCrackSpreadWorkbook() {
  const records = latestPrices_(['WTI_USD', 'GASOLINE_USD', 'HEATING_OIL_USD']);
  const byCode = Object.fromEntries(records.map((record) => [record.code, record]));
  const crude = byCode.WTI_USD.price;
  const gasoline = byCode.GASOLINE_USD.price;
  const distillate = byCode.HEATING_OIL_USD.price;

  const marketSheet = sheet_('Market Data');
  writeTitle_(marketSheet, 'Crack Spread Market Data', 'Latest available source values returned by OilPriceAPI.');
  writeTable_(marketSheet, 4, 1, marketDataRows_(records));

  const model = sheet_('Crack Model');
  writeTitle_(model, 'Crack Spread Lab', 'Gross refinery-margin proxy; excludes operating, transport, inventory, and hedge costs.');
  writeTable_(model, 4, 1, [
    ['Input', 'Value', 'Unit', 'Source'],
    ['WTI crude', crude, 'USD/barrel', 'Market Data'],
    ['Gasoline', gasoline, 'USD/gallon', 'Market Data'],
    ['Heating oil / distillate', distillate, 'USD/gallon', 'Market Data'],
    ['Gallons per barrel', 42, 'gallons', 'Conversion constant']
  ]);
  writeTable_(model, 11, 1, [
    ['Benchmark', 'Spread (USD/barrel)', 'Formula'],
    ['3-2-1', calculateCrackSpread_(crude, gasoline, distillate, '3-2-1'), '((2 × gasoline × 42) + (distillate × 42) − (3 × crude)) ÷ 3'],
    ['2-1-1', calculateCrackSpread_(crude, gasoline, distillate, '2-1-1'), '((gasoline × 42) + (distillate × 42) − (2 × crude)) ÷ 2']
  ]);
  model.getRange('B12:B13').setNumberFormat('$0.00');

  const historyRows = [['Series', 'Source timestamp', 'Price', 'Unit']];
  [
    ['WTI_USD', 'USD/barrel'],
    ['GASOLINE_USD', 'USD/gallon'],
    ['HEATING_OIL_USD', 'USD/gallon']
  ].forEach(([code, unit]) => {
    history_(code, 30).forEach((record) => historyRows.push([code, record.timestamp, record.price, unit]));
  });
  const historySheet = sheet_('History');
  writeTitle_(historySheet, 'Thirty-Day Component History', 'Series remain separate when source timestamps do not align.');
  writeTable_(historySheet, 4, 1, historyRows);

  const sensitivity = sheet_('Sensitivity');
  writeTitle_(sensitivity, 'Crude and Product Sensitivity', 'Parallel product-price changes applied to gasoline and distillate.');
  writeTable_(sensitivity, 4, 1, crackSensitivity_(crude, gasoline, distillate));
  sensitivity.getRange(5, 3, 25, 1).setNumberFormat('$0.00');

  activateProduct_();
  model.activate();
  return { success: true, message: 'Crack Spread Lab built with live market data, history, and 25 sensitivity cases.' };
}
