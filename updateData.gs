function onOpen() {
  const ui = SpreadsheetApp.getUi();

    ui.createMenu('Update Data')
    .addItem('Update Data', 'updateData')
    .addToUi();

    ui.createMenu('Pull Data')
    .addItem('Import Data (Realm + Throne)', 'runAllImports')
    .addToUi();
}

function updateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = ss.getSheetByName('Data l MetaAds');
  const mappingSheet = ss.getSheetByName('Brand Mapping');
  const aoMapping = ss.getSheetByName('AO Spend Mapping');
  const aff = ss.getSheetByName('Data l My affiliates');
  const out = ss.getSheetByName('Resumen');

  if (!meta || !mappingSheet || !aoMapping || !aff || !out) throw new Error("Falta alguna hoja.");

  const cutoffMonth = "2025-07";

  function norm(s) {
    return s ? String(s).trim().toLowerCase() : '';
  }

  function capitalize(s) {
    return s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : '';
  }

  function extractMonthKey(val) {
    try {
      const date = new Date(val);
      return !isNaN(date) ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM") : '';
    } catch { return ''; }
  }

  function formatMonthToString(key) {
    const [year, month] = key.split("-");
    const names = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"];
    return names[parseInt(month) - 1] + " " + year;
  }

  const mapVals = mappingSheet.getDataRange().getValues();
  const brandToStakeholder = {}, stakeholders = { Benji: [], Mert: [] };
  const brandList = [];

  for (let i = 1; i < mapVals.length; i++) {
    const brandRaw = String(mapVals[i][13]).trim();
    const stake = capitalize(norm(mapVals[i][12]));
    if (brandRaw) {
      const brandKey = norm(brandRaw);
      brandToStakeholder[brandKey] = stake;
      if (!brandList.includes(brandRaw)) brandList.push(brandRaw);
      if (stake === "Benji" || stake === "Mert") {
        stakeholders[stake].push(brandRaw);
      }
    }
  }

  const aoVals = aoMapping.getDataRange().getValues();
  const stakeSpendMap = {};
  for (let i = 1; i < aoVals.length; i++) {
    const rawDate = aoVals[i][1];
    const spend = parseFloat(aoVals[i][2]) || 0;
    const stake = capitalize(norm(aoVals[i][3]));
    if (!rawDate || !stake || !spend) continue;
    const date = new Date(rawDate);
    const month = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM");
    const key = `${month}|${stake}`;
    stakeSpendMap[key] = (stakeSpendMap[key] || 0) + spend;
  }

  const metaVals = meta.getDataRange().getValues();
  const sharedPreJul = {}, ownSpend = {}, directSpend = {};
  const allMonths = new Set();

  for (let i = 1; i < metaVals.length; i++) {
    const date = new Date(metaVals[i][0]);
    const monthKey = extractMonthKey(date);
    const euro = parseFloat(metaVals[i][5]) || 0;
    const campaign = norm(metaVals[i][6]);
    const brandRaw = String(metaVals[i][7]).trim();
    const brandKey = norm(brandRaw);

    if (!monthKey || !brandRaw || !campaign) continue;
    allMonths.add(monthKey);
    if (!brandList.includes(brandRaw)) brandList.push(brandRaw);

    const ownKey = `${monthKey}|${brandRaw}`;
    const fullKey = `${monthKey}|${brandRaw}|${campaign}`;

    if (campaign === 'always on') {
      ownSpend[ownKey] = (ownSpend[ownKey] || 0) + euro;
      if (monthKey < cutoffMonth && ['global', 'experiments', 'benji brands', 'mert brands'].includes(brandKey)) {
        sharedPreJul[monthKey] = (sharedPreJul[monthKey] || 0) + euro;
      }
    } else {
      directSpend[fullKey] = (directSpend[fullKey] || 0) + euro;
    }
  }

  const affVals = aff.getDataRange().getValues();
  const metricsMap = {};
  for (let i = 1; i < affVals.length; i++) {
    const mesKey = extractMonthKey(affVals[i][0]);
    const brandRaw = String(affVals[i][2]).trim();
    const camp = norm(affVals[i][10]);
    if (!mesKey || !brandRaw || !camp) continue;
    const key = `${mesKey}|${brandRaw}|${camp}`;
    if (!metricsMap[key]) {
      metricsMap[key] = { clicks: 0, signups: 0, ndc: 0, ndcAmt: 0, deps: 0 };
    }

    metricsMap[key].clicks += parseFloat(affVals[i][4]) || 0;
    metricsMap[key].signups += parseFloat(affVals[i][5]) || 0;
    metricsMap[key].ndc += parseFloat(affVals[i][6]) || 0;
    metricsMap[key].ndcAmt += parseFloat(affVals[i][7]) || 0;
    metricsMap[key].deps += parseFloat(affVals[i][8]) || 0;
  }

  const hiddenBrands = ["benji brands", "mert brands", "global", "experiments"];

  const groups = {
    "bets 10": "Flagship", "jetbahis": "Flagship", "mobil bahis": "Flagship",
    "davegas": "Niche", "discount casino": "Niche",
    "casino maxi": "Conventional", "casino metropol": "Conventional",
    "betchip": "Generic", "betelli": "Generic", "betroad": "Generic",
    "genzobet": "Generic", "hovarda": "Generic", "intobet": "Generic",
    "milyar": "Generic", "rexbet": "Generic", "slotbon": "Generic",
    "winnit": "Generic", "jokera": "Generic"
  };

  if (out.getLastRow() > 2) {
    out.getRange(3, 1, out.getLastRow() - 2, out.getMaxColumns()).clearContent();
  }

  out.getRange(2, 1, 1, 15).setValues([[
    'Month', 'Stakeholder', 'Brand', 'Group', 'Campaign', 'Amount Spend (EUR)',
    'Clicks', 'Signups', 'NDC', 'NDC Amount', 'Deposits',
    'CPA', 'CPC', 'Conversion Rate', 'Avg First Deposit'
  ]]);

  let row = 3;
  for (const month of Array.from(allMonths)) {
    const isPostJuly = month >= cutoffMonth;

    for (const brandRaw of brandList) {
      const brandKey = norm(brandRaw);
      if (hiddenBrands.includes(brandKey)) continue;

      const stake = brandToStakeholder[brandKey] || "";
      const group = groups[brandKey] || "Other";

      // ALWAYS ON
      const ownKey = `${month}|${brandRaw}`;
      const metricsKey = `${month}|${brandRaw}|always on`;
      const metrics = metricsMap[metricsKey] || { clicks: 0, signups: 0, ndc: 0, ndcAmt: 0, deps: 0 };
      let spend = 0;

      if (isPostJuly && (stake === "Benji" || stake === "Mert")) {
        const stakeKey = `${month}|${stake}`;
        const count = stakeholders[stake]?.length || 1;
        if (stakeholders[stake]?.includes(brandRaw)) {
          spend = (stakeSpendMap[stakeKey] || 0) / count;
        }
        spend += ownSpend[ownKey] || 0;
      } else if (!isPostJuly) {
        const validBrands = brandList.filter(b => !hiddenBrands.includes(norm(b)));
        const count = validBrands.length || 1;
        spend = (sharedPreJul[month] || 0) / count + (ownSpend[ownKey] || 0);
      }

      if (spend > 0 || Object.values(metrics).some(v => v > 0)) {
        const cpa = metrics.ndc ? spend / metrics.ndc : 0;
        const cpc = metrics.clicks ? spend / metrics.clicks : 0;
        const cr = metrics.clicks ? metrics.ndc / metrics.clicks : 0;
        const afd = metrics.ndc ? metrics.ndcAmt / metrics.ndc : 0;

        out.getRange(row, 1, 1, 15).setValues([[
          formatMonthToString(month), stake, brandRaw, group, "always on", spend,
          metrics.clicks, metrics.signups, metrics.ndc, metrics.ndcAmt, metrics.deps,
          cpa, cpc, cr, afd
        ]]);
        row++;
      }

      // OTRAS CAMPAÑAS
      for (const campaignKey in directSpend) {
        if (!campaignKey.startsWith(`${month}|${brandRaw}|`)) continue;
        const campaign = campaignKey.split('|')[2];
        const spendOther = directSpend[campaignKey];
        const otherMetrics = metricsMap[`${month}|${brandRaw}|${campaign}`] || { clicks: 0, signups: 0, ndc: 0, ndcAmt: 0, deps: 0 };

        if (spendOther === 0 && Object.values(otherMetrics).every(v => v === 0)) continue;

        const cpa = otherMetrics.ndc ? spendOther / otherMetrics.ndc : 0;
        const cpc = otherMetrics.clicks ? spendOther / otherMetrics.clicks : 0;
        const cr = otherMetrics.clicks ? otherMetrics.ndc / otherMetrics.clicks : 0;
        const afd = otherMetrics.ndc ? otherMetrics.ndcAmt / otherMetrics.ndc : 0;

        out.getRange(row, 1, 1, 15).setValues([[
          formatMonthToString(month), stake, brandRaw, group, campaign, spendOther,
          otherMetrics.clicks, otherMetrics.signups, otherMetrics.ndc, otherMetrics.ndcAmt, otherMetrics.deps,
          cpa, cpc, cr, afd
        ]]);
        row++;
      }
    }
  }

  if (row > 3) {
    out.getRange(3, 6, row - 3, 10).setNumberFormat('0.00');
    out.getRange(3, 14, row - 3, 1).setNumberFormat("0.00%");
  }
}
