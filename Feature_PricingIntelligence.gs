/**
 * Pricing Intelligence (AgenPRO) — weekly review, the effect of moving the price list to Home, and price coverage.
 * Sheet "Pricing Intelligence Agenpro" holds two tables side by side:
 *   A–G  daily usage: Date, App DAU, Feature DAU, First Time Feature Users, Returning Users, Weekly App (WAU), Weekly Feature (WAU)
 *   K–R  update frequency of the mill & ramp master list, one row per category × segment (PARTNER-MILLS, NON-PARTNER-MILLS,
 *        RAMP): count, % within category, % of master, min / max / median latest price
 */

const PI_CONFIG = {
  SHEET: 'Pricing Intelligence Agenpro',
  APP: 'agenpro',
  // The price list moved from a menu to the home screen on this date
  CHANGE_DATE: '2026-10-30',
  CHANGE_LABEL: 'Price list moved to Home',
  // Update-frequency categories counted as "regular" (updated at least monthly) and "within 90 days"
  REGULAR_PREFIXES: ['1.', '2.'],
  WITHIN_90_PREFIXES: ['1.', '2.', '3.'],
  // Segment values in the coverage table → label shown on the slide, in display order
  SEGMENTS: [
    { match: /^partner/i, key: 'partner', label: 'Partner mills', short: 'Partner' },
    { match: /non.?partner/i, key: 'nonpartner', label: 'Non-partner mills', short: 'Non-partner' },
    { match: /ramp/i, key: 'ramp', label: 'Ramp', short: 'Ramp' }
  ]
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
        // Exact headers first, then patterns so small renames ("Total App DAU", "Weekly App DAU (WAU)") still match
        date: ['date', 'tanggal'],
        dau: ['app_dau', 'total_app_dau', 'daily_active_users', /^(total_)?app_dau$/],
        feature: ['feature_dau', 'total_feature_dau', /^(total_)?feature_dau$/],
        newUsers: ['first_time_feature_users', 'new_users', /first_time/],
        returningUsers: ['returning_users', 'existing_returning_feature_users', /returning/],
        appWau: ['weekly_app_wau', 'weekly_app_dau_wau', 'app_weekly_active_user', /app.*wau$/],
        featureWau: ['weekly_feature_wau', 'weekly_feature_dau_wau', 'feature_weekly_active_user', /feature.*wau$/]
      }, ['date', 'dau', 'feature']);
      const series = toDailySeries_(daily.rows, ['feature'], tag, issues);
      const weekly = weeklyAnalysis_(series, ['feature'], tag, issues);
      if (!weekly) throw new Error('needs at least one complete week (Thu–Wed) of data');

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
            'menu to the home screen, so users now see prices without opening the detail page.',
          scope: [
            'Weekly review of Pricing Intelligence in AgenPRO: this week vs last week (Thursday–Wednesday), the before/after effect ' +
              'of moving the price list to Home, and price coverage: how many mills and ramps update their prices regularly, ' +
              'compared between partner mills, non-partner mills and ramps.',
            'Describe how detail users changed after the move to Home (prices are now visible on Home without opening the detail), ' +
              'and recommend product changes that make the Home price list more useful and bring users to the detail when it adds value.'
          ],
          definitions: {
            appWau: 'AgenPRO weekly active users',
            clickUsers: 'distinct users who opened the Pricing Intelligence detail during the week (NOT WAU)',
            conversionPct: 'clickUsers ÷ appWau',
            'beforeAfter.before / after': 'equal-length periods right before and from the move to Home; avgFeatureUsers = detail users per day',
            'coverage.segments': 'partner mills, non-partner mills and ramps in the master list; byCategory = how often they update ' +
              'their price (count, % of the segment, latest price range in Rp/kg)',
            'coverage.regular / regularPct': 'mills & ramps updating at least monthly (≤ 30 days), overall and per segment'
          },
          payload: {
            weeks: weekly.weeks,
            thisWeekVsLastWeek: cmp,
            estimatedDay: weekly.estimatedDay,
            beforeAfter: change,
            coverage: coverage,
            daily: series.map(r => ({ date: r.date, estimated: r.estimated || undefined, dau: r.dau, detailUsers: r.steps.feature, firstTime: r.newUsers, returning: r.returningUsers })),
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

/**
 * Update-frequency table: one row per category (× segment when the table has a segment column). Columns are found from
 * the "Kategori Frekuensi Update" header onwards: the segment column is the one holding text such as PARTNER-MILLS, the
 * count is the numeric "Jumlah …" column, and the price columns are matched by their Min / Max / Median headers.
 */
function readCoverage_(ss, issues) {
  const sheet = ss.getSheetByName(PI_CONFIG.SHEET);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = values[0].map(normalizeHeader_);
  const c0 = headers.indexOf('kategori_frekuensi_update');
  if (c0 < 0) {
    issues.push('coverage table not found (no "Kategori Frekuensi Update" column)');
    return null;
  }
  const rows = [];
  for (let r = 1; r < values.length && values[r][c0] !== '' && values[r][c0] != null; r++) rows.push(values[r]);
  const cols = headers.map((h, j) => j).filter(j => j > c0);
  const isText = j => rows.some(r => r[j] !== '' && r[j] != null && isNaN(Number(String(r[j]).replace(/,/g, ''))));
  const segCol = cols.find(isText);
  const countCol = cols.find(j => j !== segCol && /^jumlah/.test(headers[j]) && !isText(j));
  const priceCol = re => cols.find(j => re.test(headers[j]));
  const at = { min: priceCol(/^min_harga/), max: priceCol(/^max_harga/), median: priceCol(/^median_harga/) };
  if (countCol == null) {
    issues.push('coverage table: no count column');
    return null;
  }

  const segDef = v => PI_CONFIG.SEGMENTS.find(d => d.match.test(v)) ||
    { key: String(v).toLowerCase(), label: String(v), short: String(v) };
  const lines = rows.map(r => ({
    category: String(r[c0]).trim(),
    segment: segCol != null ? segDef(String(r[segCol]).trim()) : { key: 'all', label: 'Mill & ramp', short: 'All' },
    count: toNumber_(r[countCol]),
    min: at.min != null ? numOrNull_(r[at.min]) : null,
    max: at.max != null ? numOrNull_(r[at.max]) : null,
    median: at.median != null ? numOrNull_(r[at.median]) : null
  }));
  const categories = lines.map(l => l.category).filter((c, i, a) => a.indexOf(c) === i).sort();
  const isIn = (cat, prefixes) => prefixes.some(p => cat.indexOf(p) === 0);
  const summarise = list => {
    const total = list.reduce((a, l) => a + l.count, 0);
    const regular = list.filter(l => isIn(l.category, PI_CONFIG.REGULAR_PREFIXES)).reduce((a, l) => a + l.count, 0);
    const within90 = list.filter(l => isIn(l.category, PI_CONFIG.WITHIN_90_PREFIXES)).reduce((a, l) => a + l.count, 0);
    return {
      total: total,
      regular: regular,
      regularPct: total ? round1_((regular / total) * 100) : null,
      within90: within90,
      within90Pct: total ? round1_((within90 / total) * 100) : null
    };
  };

  const segKeys = lines.map(l => l.segment.key).filter((k, i, a) => a.indexOf(k) === i);
  const order = k => {
    const i = PI_CONFIG.SEGMENTS.findIndex(d => d.key === k);
    return i < 0 ? PI_CONFIG.SEGMENTS.length : i;
  };
  const segments = segKeys.sort((a, b) => order(a) - order(b)).map(k => {
    const list = lines.filter(l => l.segment.key === k);
    const sum = summarise(list);
    const weekly = list.find(l => l.category.indexOf(PI_CONFIG.REGULAR_PREFIXES[0]) === 0);
    return Object.assign({ key: k, label: list[0].segment.label, short: list[0].segment.short }, sum, {
      medianPriceWeeklyUpdaters: weekly ? weekly.median : null,
      byCategory: categories.map(c => {
        const l = list.find(x => x.category === c);
        const count = l ? l.count : 0;
        return { category: c, count: count, pctOfSegment: sum.total ? round1_((count / sum.total) * 100) : 0,
          minPrice: l ? l.min : null, medianPrice: l ? l.median : null, maxPrice: l ? l.max : null };
      })
    });
  });
  return Object.assign({ categories: categories, segments: segments }, summarise(lines));
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
  drawTitle_(slide, L.title, { feature: 'Pricing Intelligence', app: app, suffix: 'Weekly Review', pill: weekPill_(weekly), warnings: warningCount, estimatedDay: weekly.estimatedDay });

  const cards = weeklyKpiCards_(weekly.compare, app, 'USERS OPEN PRICE DETAIL').slice(0, 2);
  const b = change.before, a = change.after;
  if (a && b) {
    const st = deltaStyle_({ diff: round2_(a.avgFeatureUsers - b.avgFeatureUsers), relPct: pctChange_(a.avgFeatureUsers, b.avgFeatureUsers) }, { kind: 'count', digits: 1 });
    cards.push({ label: 'DETAIL USERS/DAY SINCE ' + formatShortDate_(change.changeDate).toUpperCase(), value: fmtNum_(a.avgFeatureUsers, 1),
      sub: st.text + ' vs ' + fmtNum_(b.avgFeatureUsers, 1) + '/day before', subColor: st.ink, accent: C.marker });
  }
  if (coverage) {
    cards.push({ label: 'MILL & RAMP UPDATING ≤ 30 DAYS', value: fmtPct_(coverage.regularPct, 1),
      // The two largest segments fit on the card; every segment is in the coverage panel
      sub: coverage.segments.length > 1
        ? coverage.segments.slice().sort((a, b) => b.total - a.total).slice(0, 2).sort((a, b) => b.regularPct - a.regularPct)
          .map(sg => sg.short + ' ' + fmtPct_(sg.regularPct, 0)).join(' vs ')
        : fmtNum_(coverage.regular) + ' of ' + fmtNum_(coverage.total) + ' in the master list',
      accent: C.greenBar });
  }
  drawKpiCards_(slide, L.kpi, cards);

  if (coverage) drawCoveragePanel_(slide, L.main, coverage);

  // Whole period, with the move to Home marked
  drawDailyTrend_(slide, L.chart, {
    title: 'DAILY TREND · ' + app.label.toUpperCase(),
    rows: series.filter(r => r.date <= weekly.windows[weekly.windows.length - 1].end).map(r => ({
      date: r.date, estimated: r.estimated, line: r.dau,
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

/**
 * Coverage by segment: for partner mills, non-partner mills and ramps, a 100% bar split by update frequency and the
 * share updating at least monthly, so the segments can be compared at a glance.
 */
function drawCoveragePanel_(slide, box, cov) {
  const C = CORE.COLORS;
  panel_(slide, box);
  sectionTitle_(slide, 'PRICE UPDATE COVERAGE  ·  ' + fmtNum_(cov.total) + ' MILL & RAMP', box.x + 10, box.y + 8, 260);

  const col = { label: box.x + 10, labelW: 88, bar: box.x + 104, barW: 206, right: box.x + 318, rightW: 92 };
  const headY = box.y + 23;
  [['SEGMENT', col.label, col.labelW, 'START'], ['UPDATE FREQUENCY · SHARE OF SEGMENT', col.bar, col.barW, 'START'],
    ['UPDATED ≤ 30 DAYS', col.right, col.rightW, 'END']]
    .forEach(h => addText_(slide, h[0], h[1], headY, h[2], 7, { fontSize: 5.5, bold: true, color: C.muted, align: h[3] }));
  hLine_(slide, box.x + 10, box.x + box.w - 10, headY + 9, C.border, 0.5);

  const palette = [C.greenDark, C.greenBar, C.returning, C.nonPartner, C.red];
  const colorOf = i => palette[i % palette.length];
  const top = box.y + 36;
  const rowH = Math.min(22, (box.y + box.h - 34 - top) / Math.max(1, cov.segments.length));

  cov.segments.forEach((sg, i) => {
    const y = top + i * rowH, cy = y + rowH / 2;
    addRichText_(slide, [{ text: sg.label, bold: true, color: C.ink }, { text: '\n' + fmtNum_(sg.total) + ' sites', color: C.muted, size: 5.5 }],
      col.label, y, col.labelW, rowH, { fontSize: 7, valign: 'MIDDLE' });

    // 100% bar: one piece per update-frequency category, count inside when it fits
    const barH = Math.min(12, rowH - 6);
    rect_(slide, SlidesApp.ShapeType.RECTANGLE, col.bar, cy - barH / 2, col.barW, barH, C.track);
    let x = col.bar;
    sg.byCategory.forEach((c, k) => {
      if (!c.count) return;
      const w = col.barW * c.count / sg.total;
      rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, cy - barH / 2, w, barH, colorOf(k), C.white, 0.5);
      if (w >= 18) addText_(slide, fmtNum_(c.count), x, cy - barH / 2, w, barH, { fontSize: 5.5, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });
      x += w;
    });

    const good = sg.regularPct >= 50, none = !sg.regular;
    addRichText_(slide, [
      { text: fmtNum_(sg.regular) + ' of ' + fmtNum_(sg.total) + '   ', color: C.muted, size: 5.5 },
      { text: fmtPct_(sg.regularPct, 1), bold: true, color: none ? C.red : good ? C.okInk : C.amber }
    ], col.right, y, col.rightW, rowH, { fontSize: 8.5, align: 'END', valign: 'MIDDLE' });
  });

  // Legend: update-frequency categories
  let lx = col.bar;
  const ly = top + cov.segments.length * rowH + 2;
  cov.categories.forEach((c, k) => {
    const label = c.replace(/^\d+\.\s*/, '').replace(/\s*\(.*$/, '');
    const range = (c.match(/\(([^)]*)\)/) || [])[1] || '';
    const text = label + (range ? ' ' + range.replace(/hari/i, 'd').replace(/tidak update/i, '').trim() : '');
    const tw = estimateTextWidth_(text, 5) + 2;
    rect_(slide, SlidesApp.ShapeType.RECTANGLE, lx, ly + 2, 6, 6, colorOf(k));
    addText_(slide, text, lx + 8, ly, tw, 10, { fontSize: 5, color: C.muted, valign: 'MIDDLE' });
    lx += tw + 11;
  });

  // Footer: overall coverage and the typical latest price of weekly updaters per segment
  const fy = box.y + box.h - 15;
  hLine_(slide, box.x + 10, box.x + box.w - 10, fy - 3, C.border, 0.5);
  const parts = [
    { text: 'All: ≤ 30 days ', color: C.muted },
    { text: fmtNum_(cov.regular) + ' (' + fmtPct_(cov.regularPct, 1) + ')', bold: true, color: C.ink },
    { text: '  ·  ≤ 90 days ', color: C.muted },
    { text: fmtNum_(cov.within90) + ' (' + fmtPct_(cov.within90Pct, 1) + ')', bold: true, color: C.ink }
  ];
  const priced = cov.segments.filter(sg => sg.medianPriceWeeklyUpdaters != null);
  if (priced.length) {
    parts.push({ text: '     Median price, weekly updaters: ', color: C.muted });
    priced.forEach((sg, k) => parts.push(
      { text: (k ? '  ·  ' : '') + sg.short + ' ', color: C.muted },
      { text: 'Rp ' + fmtNum_(sg.medianPriceWeeklyUpdaters), bold: true, color: C.ink }
    ));
  }
  addRichText_(slide, parts, box.x + 10, fy, box.w - 20, 10, { fontSize: 5.5, valign: 'MIDDLE' });
}
