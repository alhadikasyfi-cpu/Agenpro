/**
 * Pricing Intelligence (AgenPRO) — weekly review, the effect of moving the price list to Home, and price coverage.
 * Sheet "Pricing Intelligence Agenpro" holds two tables side by side:
 *   A–G  daily usage: Date, App DAU, Feature DAU, First Time Feature Users, Returning Users, Weekly App (WAU), Weekly Feature (WAU)
 *   K–P  update frequency of the mill & ramp master list: category, count, % of master, min / max / median latest price
 */

const PI_CONFIG = {
  SHEET: 'Pricing Intelligence Agenpro',
  APP: 'agenpro',
  // The price list moved from a menu to the home screen on this date
  CHANGE_DATE: '2026-10-30',
  CHANGE_LABEL: 'Price list moved to Home',
  // Update-frequency categories counted as "regular" (updated at least monthly) and "within 90 days"
  REGULAR_PREFIXES: ['1.', '2.'],
  WITHIN_90_PREFIXES: ['1.', '2.', '3.']
};

function pricingIntelligenceFeature_() {
  return {
    id: 'pricing_intelligence',
    name: 'Pricing Intelligence',
    menuFunction: 'generatePricingIntelligenceSlides',
    buildSlides: (ss, issues) => {
      const app = CORE.APPS[PI_CONFIG.APP];
      const tag = app.label + ' ';
      const daily = readTable_(ss, PI_CONFIG.SHEET, {
        date: ['date', 'tanggal'],
        dau: ['app_dau', 'daily_active_users'],
        feature: ['feature_dau'],
        newUsers: ['first_time_feature_users', 'new_users'],
        returningUsers: ['returning_users'],
        appWau: ['weekly_app_wau', 'app_weekly_active_user'],
        featureWau: ['weekly_feature_wau', 'feature_weekly_active_user']
      }, ['date', 'dau', 'feature']);
      const series = toDailySeries_(daily.rows, ['feature'], tag, issues);
      const weekly = weeklyAnalysis_(series, ['feature'], tag, issues);
      if (!weekly) throw new Error('needs at least one complete week (Thu–Wed) of data');
      if (weekly.daysInProgress) issues.push(tag + weekly.daysInProgress + ' day(s) of the current week wait until the week is complete');

      const coverage = readCoverage_(ss, issues);
      const change = beforeAfter_(series, PI_CONFIG.CHANGE_DATE);
      const title = 'Pricing Intelligence · ' + app.label;
      const cmp = weekly.compare;

      return [{
        title: title,
        ai: {
          feature: 'Pricing Intelligence',
          app: app,
          description: 'Shows the latest FFB/TBS buying prices of mills and ramps (partner and non-partner) so users know where to ' +
            'deliver their fruit and at what price today. On ' + formatShortDate_(PI_CONFIG.CHANGE_DATE) + ' the price list moved from a ' +
            'menu to the home screen, so users now see prices without opening the detail page; detail interactions are expected to drop.',
          scope: [
            'Weekly review of Pricing Intelligence in AgenPRO: this week vs last week (Thursday–Wednesday), the before/after effect ' +
              'of moving the price list to Home, and price coverage (how many mills and ramps update their prices regularly).',
            'Do not treat the drop in detail users after the move as a failure by default: prices are now visible on Home without ' +
              'opening the detail. Recommend product changes that make the Home price list more useful and bring users to the detail when it adds value.'
          ],
          definitions: {
            appWau: 'AgenPRO weekly active users',
            clickUsers: 'distinct users who opened the Pricing Intelligence detail during the week (NOT WAU)',
            conversionPct: 'clickUsers ÷ appWau',
            'beforeAfter.before / after': 'equal-length periods right before and from the move to Home; avgFeatureUsers = detail users per day',
            'coverage.categories': 'mills & ramps in the master list by how often they update their price; prices are their latest price in Rp/kg',
            'coverage.regular': 'mills & ramps updating at least monthly (≤ 30 days)'
          },
          payload: {
            weeks: weekly.weeks,
            thisWeekVsLastWeek: cmp,
            beforeAfter: change,
            coverage: coverage,
            daily: series.map(r => ({ date: r.date, dau: r.dau, detailUsers: r.steps.feature, firstTime: r.newUsers, returning: r.returningUsers })),
            dataQualityIssues: issues.slice()
          },
          fallbackHeadline: title + ': ' + fmtNum_(cmp.current.clickUsers) + ' detail users this week.'
        },
        draw: (slide, ai) => drawPricingSlide_(slide, app, series, weekly, change, coverage, ai, issues.length)
      }];
    }
  };
}

function generatePricingIntelligenceSlides() {
  generateFeatures_(['pricing_intelligence']);
}

/** Update-frequency table (K–P): categories with count, share and latest-price range, plus regular / within-90-day totals. */
function readCoverage_(ss, issues) {
  let table;
  try {
    table = readTable_(ss, PI_CONFIG.SHEET, {
      category: ['kategori_frekuensi_update'],
      count: ['jumlah_mill_ramp'],
      pct: [/^persentase/],
      min: [/^min_harga/],
      max: [/^max_harga/],
      median: [/^median_harga/]
    }, ['category', 'count']);
  } catch (err) {
    issues.push('coverage table not found: ' + err.message);
    return null;
  }
  const cats = table.rows.map(r => ({
    category: String(r.category).trim(),
    count: toNumber_(r.count),
    min: numOrNull_(r.min), max: numOrNull_(r.max), median: numOrNull_(r.median)
  }));
  const total = cats.reduce((acc, c) => acc + c.count, 0);
  cats.forEach(c => c.pct = total ? round2_((c.count / total) * 100) : null);
  const sumWhere = prefixes => cats.filter(c => prefixes.some(p => c.category.indexOf(p) === 0)).reduce((acc, c) => acc + c.count, 0);
  const regular = sumWhere(PI_CONFIG.REGULAR_PREFIXES), within90 = sumWhere(PI_CONFIG.WITHIN_90_PREFIXES);
  return {
    total: total,
    categories: cats,
    regular: regular,
    regularPct: total ? round1_((regular / total) * 100) : null,
    within90: within90,
    within90Pct: total ? round1_((within90 / total) * 100) : null
  };
}

/** Equal-length periods right before and from the change date (the "after" period is all data from that date). */
function beforeAfter_(series, changeKey) {
  const after = series.filter(r => r.date >= changeKey);
  const before = series.filter(r => r.date < changeKey).slice(-Math.max(1, after.length));
  const agg = rows => {
    if (!rows.length) return null;
    const dau = rows.reduce((a, r) => a + r.dau, 0), users = rows.reduce((a, r) => a + r.steps.feature, 0);
    return {
      from: rows[0].date, to: rows[rows.length - 1].date, days: rows.length,
      avgDau: round1_(dau / rows.length),
      avgFeatureUsers: round1_(users / rows.length),
      ctrPct: dau ? round2_((users / dau) * 100) : null,
      firstTimeUsers: rows.reduce((a, r) => a + (r.newUsers || 0), 0),
      daysWithZeroUsers: rows.filter(r => r.steps.feature === 0).length
    };
  };
  return { changeDate: changeKey, before: agg(before), after: agg(after) };
}

function drawPricingSlide_(slide, app, series, weekly, change, coverage, ai, warningCount) {
  const C = CORE.COLORS, L = CORE.LAYOUT;
  drawTitle_(slide, L.title, { feature: 'Pricing Intelligence', app: app, suffix: 'Weekly Review', pill: weekPill_(weekly), warnings: warningCount });

  const cards = weeklyKpiCards_(weekly.compare, app, 'USERS OPEN PRICE DETAIL').slice(0, 2);
  const b = change.before, a = change.after;
  if (a && b) {
    const st = deltaStyle_({ diff: round2_(a.avgFeatureUsers - b.avgFeatureUsers), relPct: pctChange_(a.avgFeatureUsers, b.avgFeatureUsers) }, { kind: 'count', digits: 1 });
    cards.push({ label: 'DETAIL USERS/DAY SINCE ' + formatShortDate_(change.changeDate).toUpperCase(), value: fmtNum_(a.avgFeatureUsers, 1),
      sub: st.text + ' vs ' + fmtNum_(b.avgFeatureUsers, 1) + '/day before', subColor: st.ink, accent: C.marker });
  }
  if (coverage) {
    cards.push({ label: 'MILL & RAMP UPDATING ≤ 30 DAYS', value: fmtNum_(coverage.regular),
      sub: fmtPct_(coverage.regularPct, 1) + ' of ' + fmtNum_(coverage.total) + ' in the master list', accent: C.greenBar });
  }
  drawKpiCards_(slide, L.kpi, cards);

  if (coverage) drawCoveragePanel_(slide, L.main, coverage);

  // Whole period, with the move to Home marked
  drawDailyTrend_(slide, L.chart, {
    title: 'DAILY TREND · ' + app.label.toUpperCase(),
    rows: series.filter(r => r.date <= weekly.windows[weekly.windows.length - 1].end).map(r => ({
      date: r.date, line: r.dau,
      segs: r.newUsers != null && r.returningUsers != null ? [[r.newUsers, app.color], [r.returningUsers, C.returning]] : [[r.steps.feature, app.color]],
      total: r.steps.feature, pct: r.ctrPct
    })),
    legend: [{ label: 'DAU (left)', color: C.blue, kind: 'line' }, { label: 'First-time', color: app.color, kind: 'bar' },
      { label: 'Returning (detail users, right)', color: C.returning, kind: 'bar' }],
    rightColor: app.color,
    marker: { date: PI_CONFIG.CHANGE_DATE, label: PI_CONFIG.CHANGE_LABEL + ' (' + formatShortDate_(PI_CONFIG.CHANGE_DATE) + ')' }
  });
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

/** Mill & ramp master list by price-update frequency: share bars, counts and latest-price range per category. */
function drawCoveragePanel_(slide, box, cov) {
  const C = CORE.COLORS;
  panel_(slide, box);
  sectionTitle_(slide, 'PRICE UPDATE COVERAGE  ·  ' + fmtNum_(cov.total) + ' MILL & RAMP', box.x + 10, box.y + 8, 260);

  const col = { label: box.x + 10, labelW: 122, bar: box.x + 136, barW: 116, count: box.x + 258, countW: 56, price: box.x + 318, priceW: 92 };
  const headY = box.y + 23;
  [['UPDATE FREQUENCY', col.label, col.labelW, 'START'], ['SHARE OF MASTER LIST', col.bar, col.barW, 'START'],
    ['COUNT', col.count, col.countW, 'END'], ['LAST PRICE MIN–MED–MAX', col.price, col.priceW, 'END']]
    .forEach(h => addText_(slide, h[0], h[1], headY, h[2], 7, { fontSize: 5.5, bold: true, color: C.muted, align: h[3] }));
  hLine_(slide, box.x + 10, box.x + box.w - 10, headY + 9, C.border, 0.5);

  const palette = [C.greenDark, C.greenBar, C.returning, C.nonPartner, C.red];
  const top = box.y + 36, bottom = box.y + box.h - 20;
  const rowH = Math.min(16, (bottom - top) / Math.max(1, cov.categories.length));
  cov.categories.forEach((c, i) => {
    const y = top + i * rowH, cy = y + rowH / 2;
    const color = palette[i % palette.length];
    const label = c.category.replace(/^\d+\.\s*/, '');
    addText_(slide, label, col.label, y, col.labelW, rowH, { fontSize: fitFontSize_([label], col.labelW, 8, 6.5, 5.5, 0, true), bold: true, color: C.ink, valign: 'MIDDLE' });
    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.bar, cy - 4, col.barW, 8, C.track);
    if (c.pct) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.bar, cy - 4, Math.max(3, col.barW * c.pct / 100), 8, color);
    addRichText_(slide, [{ text: fmtNum_(c.count), bold: true, color: C.ink }, { text: '  ' + fmtPct_(c.pct, 1), color: C.muted }],
      col.count, y, col.countW, rowH, { fontSize: 6.5, align: 'END', valign: 'MIDDLE' });
    const price = c.min != null ? fmtNum_(c.min) + ' – ' + fmtNum_(c.median) + ' – ' + fmtNum_(c.max) : '–';
    addText_(slide, price, col.price, y, col.priceW, rowH, { fontSize: 6, color: C.muted, align: 'END', valign: 'MIDDLE' });
  });

  const fy = box.y + box.h - 15;
  hLine_(slide, box.x + 10, box.x + box.w - 10, fy - 3, C.border, 0.5);
  addRichText_(slide, [
    { text: 'Regular updates (≤ 30 days): ', color: C.muted },
    { text: fmtNum_(cov.regular) + ' (' + fmtPct_(cov.regularPct, 1) + ')', bold: true, color: C.okInk },
    { text: '   ·   within 90 days: ', color: C.muted },
    { text: fmtNum_(cov.within90) + ' (' + fmtPct_(cov.within90Pct, 1) + ')', bold: true, color: C.ink },
    { text: '   ·   older or inactive: ', color: C.muted },
    { text: fmtNum_(cov.total - cov.within90) + ' (' + fmtPct_(cov.total ? 100 - cov.within90Pct : null, 1) + ')', bold: true, color: C.red }
  ], box.x + 10, fy, box.w - 20, 10, { fontSize: 6, valign: 'MIDDLE' });
}
