function surchargePerMile_(indexPrice, basePrice, milesPerGallon) {
  const index = Number(indexPrice);
  const base = Number(basePrice);
  const mpg = Number(milesPerGallon);
  if (![index, base, mpg].every(Number.isFinite)) throw new Error('Surcharge inputs must be finite.');
  if (index < 0 || base < 0) throw new Error('Diesel index and base price must be non-negative.');
  if (mpg <= 0) throw new Error('MPG must be positive.');
  return Math.max(0, index - base) / mpg;
}

function surchargeBands_(basePrice, milesPerGallon, startPrice, step, bandCount) {
  const count = Number(bandCount);
  if (!Number.isInteger(count) || count < 1 || count > 100) throw new Error('Band count must be an integer from 1 to 100.');
  if (![basePrice, milesPerGallon, startPrice, step].map(Number).every(Number.isFinite)) {
    throw new Error('Band inputs must be finite.');
  }
  if (Number(startPrice) < 0) throw new Error('Band start price must be non-negative.');
  if (Number(step) <= 0) throw new Error('Band step must be positive.');
  const rows = [['Diesel index from', 'Diesel index through', 'Surcharge per mile']];
  for (let index = 0; index < count; index += 1) {
    const lower = Number(startPrice) + (index * Number(step));
    const upper = lower + Number(step) - 0.001;
    rows.push([lower, upper, surchargePerMile_(lower, basePrice, milesPerGallon)]);
  }
  return rows;
}

function buildFuelSurchargeWorkbook() {
  const records = latestPrices_(OPA_PRODUCT.allowedCodes);
  const national = records.find((record) => record.code === 'DIESEL_RETAIL_USD');
  const basePrice = Math.max(0, national.price - 0.75);
  const mpg = 6.5;

  const indexes = sheet_('Regional Indexes');
  writeTitle_(indexes, 'Regional Diesel Indexes', 'Latest available retail diesel values returned by OilPriceAPI.');
  writeTable_(indexes, 4, 1, marketDataRows_(records));
  indexes.getRange(5, 2, records.length, 1).setNumberFormat('$0.000');

  const calculator = sheet_('Surcharge Calculator');
  writeTitle_(calculator, 'Fuel Surcharge Studio', 'Transparent index-minus-base calculation; confirm the method in each carrier agreement.');
  writeTable_(calculator, 4, 1, [
    ['Input', 'Value', 'Unit'],
    ['Current national diesel index', national.price, 'USD/gallon'],
    ['Contract base diesel price', basePrice, 'USD/gallon'],
    ['Fleet fuel economy', mpg, 'miles/gallon'],
    ['Calculated surcharge', surchargePerMile_(national.price, basePrice, mpg), 'USD/mile']
  ]);
  calculator.getRange('B8').setFormula('=MAX(0,B5-B6)/B7');
  calculator.getRange('B5:B6').setNumberFormat('$0.000');
  calculator.getRange('B8').setNumberFormat('$0.000');

  const schedule = sheet_('Surcharge Table');
  writeTitle_(schedule, 'Publishable Surcharge Bands', 'Twenty five-cent diesel-index bands using the calculator assumptions.');
  const bands = surchargeBands_(basePrice, mpg, Math.max(0, basePrice), 0.25, 20);
  writeTable_(schedule, 4, 1, bands);
  const bandFormulas = Array.from({ length: 20 }, (_, index) => {
    const row = index + 5;
    return [
      `='Surcharge Calculator'!$B$6+(ROW()-5)*0.25`,
      `=A${row}+0.249`,
      `=MAX(0,A${row}-'Surcharge Calculator'!$B$6)/'Surcharge Calculator'!$B$7`
    ];
  });
  schedule.getRange(5, 1, 20, 3).setFormulas(bandFormulas);
  schedule.getRange(5, 1, 20, 3).setNumberFormat('$0.000');

  activateProduct_();
  calculator.activate();
  return { success: true, message: 'Fuel surcharge workbook built with live regional indexes and a 20-band schedule.' };
}
