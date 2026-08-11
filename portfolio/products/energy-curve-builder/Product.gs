function curveContracts_(payload) {
  const data = payload && payload.data;
  const records = payload && Array.isArray(payload.contracts)
    ? payload.contracts
    : Array.isArray(data)
      ? data
      : data && Array.isArray(data.contracts)
        ? data.contracts
        : [];
  const normalized = records.map((record) => ({
    month: String(record.contract_month || record.month || record.symbol || ''),
    price: Number(record.settlement_price ?? record.settlement ?? record.price),
    sourceTimestamp: String(record.trading_date || record.created_at || record.as_of || record.timestamp || '')
  }));
  if (normalized.length < 2 || normalized.some((record) => !record.month || !Number.isFinite(record.price))) {
    throw new Error('Futures curve response needs at least two dated contracts with finite prices.');
  }
  return normalized;
}

function fetchCurve_(contract) {
  const code = String(contract || '').trim().toLowerCase();
  if (!['ice-wti', 'ice-brent'].includes(code)) {
    throw new Error('Supported curve contracts are ice-wti and ice-brent.');
  }
  return curveContracts_(requestJson_(`/futures/${encodeURIComponent(code)}/curve`, requireApiKey_()));
}

function curveTable_(records) {
  const rows = [['Contract month', 'Price', 'Calendar spread', 'Structure']];
  records.forEach((record, index) => {
    const spread = index === 0 ? '' : record.price - records[index - 1].price;
    const structure = index === 0 ? 'Front month' : spread > 0 ? 'Contango step' : spread < 0 ? 'Backwardation step' : 'Flat step';
    rows.push([record.month, record.price, spread, structure]);
  });
  return rows;
}

function productConnectionProbe_() {
  const record = fetchCurve_('ice-wti')[0];
  return { contract: record.month, timestamp: record.sourceTimestamp };
}

function buildEnergyCurveWorkbook() {
  const wti = fetchCurve_('ice-wti');
  const brent = fetchCurve_('ice-brent');

  const wtiSheet = sheet_('WTI Curve');
  writeTitle_(wtiSheet, 'WTI Futures Curve', 'Contract prices and month-on-month calendar spreads.');
  writeTable_(wtiSheet, 4, 1, curveTable_(wti));
  wtiSheet.getRange(5, 2, wti.length, 2).setNumberFormat('$0.00');

  const brentSheet = sheet_('Brent Curve');
  writeTitle_(brentSheet, 'Brent Futures Curve', 'Contract prices and month-on-month calendar spreads.');
  writeTable_(brentSheet, 4, 1, curveTable_(brent));
  brentSheet.getRange(5, 2, brent.length, 2).setNumberFormat('$0.00');

  const spreads = sheet_('Calendar Spreads');
  writeTitle_(spreads, 'Front Calendar Spreads', 'Positive back-month minus front-month values indicate an upward-sloping step.');
  writeTable_(spreads, 4, 1, [
    ['Market', 'Front month', 'Second month', 'Second minus front', 'Signal'],
    ['WTI', wti[0].month, wti[1].month, wti[1].price - wti[0].price, wti[1].price > wti[0].price ? 'Contango' : 'Backwardation'],
    ['Brent', brent[0].month, brent[1].month, brent[1].price - brent[0].price, brent[1].price > brent[0].price ? 'Contango' : 'Backwardation']
  ]);
  spreads.getRange('D5:E6').setFormulas([
    ['=\'WTI Curve\'!B6-\'WTI Curve\'!B5', '=IF(D5>0,"Contango",IF(D5<0,"Backwardation","Flat"))'],
    ['=\'Brent Curve\'!B6-\'Brent Curve\'!B5', '=IF(D6>0,"Contango",IF(D6<0,"Backwardation","Flat"))']
  ]);
  spreads.getRange('D5:D6').setNumberFormat('$0.00');

  const dashboard = sheet_('Curve Dashboard');
  writeTitle_(dashboard, 'Energy Futures Curve Builder', 'A term-structure workbook, not a generic quote downloader.');
  writeTable_(dashboard, 4, 1, [
    ['Market', 'Front month', 'Front price', 'Curve points', 'Front structure'],
    ['WTI', wti[0].month, wti[0].price, wti.length, wti[1].price > wti[0].price ? 'Contango' : 'Backwardation'],
    ['Brent', brent[0].month, brent[0].price, brent.length, brent[1].price > brent[0].price ? 'Contango' : 'Backwardation']
  ]);
  dashboard.getRange('C5:C6').setNumberFormat('$0.00');

  activateProduct_();
  dashboard.activate();
  return { success: true, message: 'Energy curve workbook built with WTI and Brent curves, calendar spreads, and structure signals.' };
}
