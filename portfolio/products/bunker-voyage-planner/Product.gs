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
