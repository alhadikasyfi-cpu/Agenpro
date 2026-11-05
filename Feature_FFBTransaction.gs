/**
 * FFB Transaction (TimPRO) — Partner vs Non-Partner tonnage (and cutting ratio when the sheet has it).
 * Prices are not analysed: partner transactions are often entered without a price.
 * Sheet "FFB transaction Timpro (P, np)": one row per date with total_transactions and a details cell holding every
 * transaction of that day ("type; field; field; …" separated by ", "). The field order is read from the header text in
 * parentheses, so a new field such as cutting_ratio_pct is picked up once it is added to the header and the details.
 * The report month is the last complete month (in November, October is compared with September).
 */

const FFB_CONFIG = {
  SHEET: 'FFB transaction Timpro (P, np)',
  APP: 'timpro',
  TYPES: ['PARTNER', 'NON_PARTNER'],
  TYPE_LABELS: { PARTNER: 'Partner', NON_PARTNER: 'Non-partner' },
  MAX_MONTHS: 6
};

function ffbTransactionFeature_() {
  return {
    id: 'ffb_transaction',
    name: 'FFB Transaction',
    menuFunction: 'generateFfbTransactionSlides',
    buildSlides: (ss, issues) => {
      const app = CORE.APPS[FFB_CONFIG.APP];
      const tx = readFfbTransactions_(ss, issues);
      if (!tx.length) throw new Error('no transactions found');
      const months = monthlyFfb_(tx);
      const lastDate = tx[tx.length - 1].date;
      const report = reportMonth_(months, lastDate);
      if (!report) throw new Error('needs at least one complete month of data');
      const prev = months.find(m => m.month === previousMonthKey_(report.month)) || null;
      const weekly = weeklyFfb_(tx);
      // Cutting ratio is only analysed when the sheet has it
      const hasCutting = tx.some(t => t.cuttingRatioPct != null);
      const title = 'FFB Transaction · ' + app.label;

      return [{
        title: title,
        ai: {
          feature: 'FFB Transaction',
          app: app,
          description: 'FFB/TBS deliveries recorded by pengurus in TimPRO. PARTNER = delivered to a mill that partners with SawitPRO ' +
            '(earns the partner program benefits). NON_PARTNER = delivered to a non-partner mill; the pengurus uploads or enters the ' +
            'weighbridge ticket to earn the TimPRO non-partner program benefits.' + (hasCutting ? ' Cutting ratio = deduction weight ÷ ' +
            'gross weight; higher means more rejected fruit (unripe, long stalks, rotten, dirt).' : ''),
          scope: [
            'Monthly review of ' + monthLabel_(report.month) + ' vs ' + (prev ? monthLabel_(prev.month) : 'no previous month') +
              ': total, partner and non-partner tonnage and partner share' + (hasCutting ? ', cutting ratio' : '') +
              ', plus the weekly tonnage trend. Focus on tonnage.'
          ],
          definitions: Object.assign({
            kg: 'sum of weight_net_kg; tonnes = kg ÷ 1000',
            partnerSharePct: 'partner kg ÷ total kg',
            topMills: 'mills (buyer) with the most kg in the report month'
          }, hasCutting ? { avgCuttingRatioPct: 'average cutting_ratio_pct of the group' } : {}),
          extraRules: ['- Do not discuss prices; they are often not entered.'],
          payload: {
            reportMonth: report,
            previousMonth: prev,
            months: months,
            weeks: weekly,
            dataQualityIssues: issues.slice()
          },
          fallbackHeadline: title + ': ' + fmtTonnes_(report.total.kg) + ' delivered in ' + monthLabel_(report.month) + '.'
        },
        draw: (slide, ai) => drawFfbSlide_(slide, app, months, report, prev, weekly, hasCutting, ai, issues.length)
      }];
    }
  };
}

function generateFfbTransactionSlides() {
  generateFeatures_(['ffb_transaction']);
}

/** All transactions, oldest first: { date, type, kg, buyer, cuttingRatioPct }. */
function readFfbTransactions_(ss, issues) {
  const table = readTable_(ss, FFB_CONFIG.SHEET, {
    date: ['transaction_date', 'date', 'tanggal'],
    count: ['total_transactions'],
    details: [/details/]
  }, ['date', 'details']);

  // Field order from the header, e.g. "Transaction Date details(partner_type; business name; …)"
  const m = String(table.headers.details).match(/\(([^)]*)\)/);
  const fields = (m ? m[1] : 'partner_type; business name; buyer name; seller name; price per kg; weight_net_kg; short name; supplier name')
    .split(';').map(normalizeHeader_);
  const idx = name => fields.indexOf(name);
  const at = {
    type: idx('partner_type'),
    kg: idx('weight_net_kg'),
    buyer: idx('buyer_name'),
    cutting: fields.findIndex(f => /cutting_ratio/.test(f))
  };
  if (at.type < 0 || at.kg < 0) throw new Error('details header must list partner_type and weight_net_kg');

  const tx = [];
  table.rows.forEach(r => {
    const d = parseDate_(r.date);
    if (!d) {
      issues.push('row ' + r.rowNumber + ': unreadable date (row skipped)');
      return;
    }
    const date = dateKey_(d);
    // A new transaction starts at "PARTNER;" or "NON_PARTNER;" (other fields may contain commas)
    const items = String(r.details || '').split(/,\s*(?=(?:PARTNER|NON_PARTNER)\s*;)/).filter(s => s.trim());
    items.forEach(item => {
      const f = item.split(';').map(s => s.trim());
      const type = String(f[at.type] || '').toUpperCase();
      if (FFB_CONFIG.TYPES.indexOf(type) < 0) {
        issues.push(date + ': unknown partner_type "' + f[at.type] + '"');
        return;
      }
      tx.push({
        date: date,
        type: type,
        kg: toNumber_(f[at.kg]),
        buyer: at.buyer >= 0 ? String(f[at.buyer] || '').trim().toUpperCase() : '',
        cuttingRatioPct: at.cutting >= 0 ? numOrNull_(f[at.cutting]) : null
      });
    });
    const expected = r.count != null && r.count !== '' ? toNumber_(r.count) : null;
    if (expected != null && expected !== items.length) issues.push(date + ': total_transactions ' + expected + ' but ' + items.length + ' in the details');
  });
  return tx.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
}

/** Totals of a group of transactions. */
function ffbGroup_(list, type) {
  const cut = list.filter(t => t.cuttingRatioPct != null);
  const group = {
    type: type,
    tx: list.length,
    kg: Math.round(list.reduce((a, t) => a + t.kg, 0)),
    avgKgPerTx: list.length ? Math.round(list.reduce((a, t) => a + t.kg, 0) / list.length) : null
  };
  if (cut.length) group.avgCuttingRatioPct = round2_(cut.reduce((a, t) => a + t.cuttingRatioPct, 0) / cut.length);
  return group;
}

/** One entry per month (oldest first, last MAX_MONTHS): totals per type, partner share and top mills. */
function monthlyFfb_(tx) {
  const byMonth = {};
  tx.forEach(t => (byMonth[t.date.slice(0, 7)] = byMonth[t.date.slice(0, 7)] || []).push(t));
  return Object.keys(byMonth).sort().slice(-FFB_CONFIG.MAX_MONTHS).map(month => {
    const list = byMonth[month];
    const byType = FFB_CONFIG.TYPES.map(type => ffbGroup_(list.filter(t => t.type === type), type));
    const total = ffbGroup_(list, 'ALL');
    const topMills = {};
    FFB_CONFIG.TYPES.forEach(type => {
      const kg = {};
      list.filter(t => t.type === type && t.buyer).forEach(t => kg[t.buyer] = (kg[t.buyer] || 0) + t.kg);
      topMills[type] = Object.keys(kg).sort((a, b) => kg[b] - kg[a]).slice(0, 3).map(b => ({ mill: b, kg: Math.round(kg[b]) }));
    });
    return {
      month: month,
      lastDay: list[list.length - 1].date,
      days: new Set(list.map(t => t.date)).size,
      total: total,
      byType: byType,
      partnerSharePct: total.kg ? round1_((byType[0].kg / total.kg) * 100) : null,
      topMills: topMills
    };
  });
}

/** Last complete month: the latest month if the data reaches its last day, otherwise the month before. */
function reportMonth_(months, lastDate) {
  const d = parseDate_(lastDate);
  const monthComplete = addDays_(d, 1).getDate() === 1;
  const key = monthComplete ? lastDate.slice(0, 7) : previousMonthKey_(lastDate.slice(0, 7));
  return months.find(m => m.month === key) || null;
}

function previousMonthKey_(key) {
  const y = Number(key.slice(0, 4)), m = Number(key.slice(5, 7));
  return m === 1 ? (y - 1) + '-12' : y + '-' + String(m - 1).padStart(2, '0');
}

function monthLabel_(key) {
  return Utilities.formatDate(new Date(Number(key.slice(0, 4)), Number(key.slice(5, 7)) - 1, 1), Session.getScriptTimeZone(), 'MMM yyyy');
}

/** Complete Thursday–Wednesday weeks: kg and transactions per type. */
function weeklyFfb_(tx) {
  return weekWindows_(tx[0].date, tx[tx.length - 1].date).map(w => {
    const list = tx.filter(t => t.date >= w.start && t.date <= w.end);
    const out = { start: w.start, end: w.end, totalKg: Math.round(list.reduce((a, t) => a + t.kg, 0)) };
    FFB_CONFIG.TYPES.forEach(type => {
      const g = list.filter(t => t.type === type);
      out[type] = { kg: Math.round(g.reduce((a, t) => a + t.kg, 0)), tx: g.length };
    });
    out.partnerSharePct = out.totalKg ? round1_((out.PARTNER.kg / out.totalKg) * 100) : null;
    return out;
  });
}

/** KPI card: partner share of tonnage in the report month vs the previous month (in pp). */
function partnerShareCard_(report, prev) {
  const C = CORE.COLORS;
  const card = { label: 'PARTNER SHARE OF TONNAGE', value: fmtPct_(report.partnerSharePct, 1), accent: C.returning };
  if (prev && prev.partnerSharePct != null) {
    const st = deltaStyle_({ diff: round2_(report.partnerSharePct - prev.partnerSharePct) }, { kind: 'pct' });
    card.sub = st.text + ' vs ' + fmtPct_(prev.partnerSharePct, 1) + ' last month';
    card.subColor = st.ink;
  }
  return card;
}

function fmtTonnes_(kg) {
  return kg == null ? '–' : fmtNum_(kg / 1000, kg < 10000 ? 1 : 0) + ' t';
}

function drawFfbSlide_(slide, app, months, report, prev, weekly, hasCutting, ai, warningCount) {
  const C = CORE.COLORS, L = CORE.LAYOUT;
  const p = report.byType[0], np = report.byType[1];
  const pp = prev ? prev.byType[0] : null, pnp = prev ? prev.byType[1] : null;
  drawTitle_(slide, L.title, {
    feature: 'FFB Transaction', app: app, suffix: 'Partner vs Non-partner',
    pill: 'MONTH  ' + monthLabel_(report.month) + (prev ? '   vs   ' + monthLabel_(prev.month) : ''), warnings: warningCount
  });

  const tonCard = (label, cur, before, accent) => {
    const card = { label: label, value: fmtTonnes_(cur), accent: accent };
    if (before != null) {
      const diff = cur - before, rel = pctChange_(cur, before);
      card.sub = (diff > 0 ? '▲ +' : diff < 0 ? '▼ −' : '■ ') + fmtTonnes_(Math.abs(diff)) +
        (rel != null && diff !== 0 ? ' (' + (rel > 0 ? '+' : '−') + fmtNum_(Math.abs(rel), 0) + '%)' : '') + ' vs ' + fmtTonnes_(before);
      card.subColor = diff > 0 ? C.okInk : diff < 0 ? C.red : C.muted;
    }
    return card;
  };
  drawKpiCards_(slide, L.kpi, [
    tonCard('TOTAL TONNAGE · ' + monthLabel_(report.month).toUpperCase(), report.total.kg, prev && prev.total.kg, C.blue),
    tonCard('PARTNER TONNAGE', p.kg, pp && pp.kg, C.greenDark),
    tonCard('NON-PARTNER TONNAGE', np.kg, pnp && pnp.kg, C.nonPartner),
    partnerShareCard_(report, prev)
  ]);

  drawFfbMonthlyPanel_(slide, L.main, months, report, hasCutting);
  drawFfbWeeklyChart_(slide, L.chart, weekly);
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

/** Left: monthly tonnage stacked bars (partner / non-partner). Right: report-month comparison table. */
function drawFfbMonthlyPanel_(slide, box, months, report, hasCutting) {
  const C = CORE.COLORS;
  panel_(slide, box);
  sectionTitle_(slide, 'MONTHLY TONNAGE  ·  PARTNER vs NON-PARTNER', box.x + 10, box.y + 8, 260);

  // Monthly bars
  const chart = { x: box.x + 14, y: box.y + 34, w: 176, h: box.h - 34 - 22 };
  const max = niceMax_(Math.max(1, ...months.map(m => m.total.kg)));
  const band = chart.w / months.length;
  const barW = Math.min(30, band * 0.55);
  hLine_(slide, chart.x, chart.x + chart.w, chart.y + chart.h, C.border, 0.75);
  months.forEach((m, i) => {
    const cx = chart.x + (i + 0.5) * band;
    const partial = m.month > report.month;
    let y = chart.y + chart.h;
    [[m.byType[0].kg, C.greenDark], [m.byType[1].kg, C.nonPartner]].forEach(sg => {
      const h = (sg[0] / max) * chart.h;
      if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, cx - barW / 2, y - h, barW, h, sg[1], partial ? C.muted : null, 0.5);
      y -= h;
    });
    addText_(slide, fmtTonnes_(m.total.kg), cx - band / 2, y - 9, band, 8, { fontSize: 5.5, bold: true, color: C.ink, align: 'CENTER', valign: 'MIDDLE' });
    addRichText_(slide, [
      { text: monthLabel_(m.month).slice(0, 3) + (partial ? ' (MTD)' : ''), color: m.month === report.month ? C.ink : C.muted, bold: m.month === report.month },
      { text: '\nP ' + fmtPct_(m.partnerSharePct, 1), color: C.greenDark, bold: true }
    ], cx - band / 2, chart.y + chart.h + 2, band, 16, { fontSize: 5.5, align: 'CENTER' });
  });

  // Report-month comparison table
  const t = { x: box.x + 206, w: box.w - 216 };
  const colW = { label: 86, p: (t.w - 86) / 2 };
  const top = box.y + 26;
  addText_(slide, monthLabel_(report.month).toUpperCase(), t.x, top, colW.label, 9, { fontSize: 5.5, bold: true, color: C.muted, valign: 'MIDDLE' });
  pill_(slide, 'Partner', t.x + colW.label + 2, top, colW.p - 4, 10, C.greenDark, C.white, 5.5);
  pill_(slide, 'Non-partner', t.x + colW.label + colW.p + 2, top, colW.p - 4, 10, C.nonPartner, C.white, 5.5);
  const p = report.byType[0], np = report.byType[1];
  const rows = [
    ['Tonnage', fmtTonnes_(p.kg), fmtTonnes_(np.kg)],
    ['Share of tonnage', fmtPct_(report.partnerSharePct, 1), fmtPct_(report.partnerSharePct == null ? null : 100 - report.partnerSharePct, 1)],
    ['Transactions', fmtNum_(p.tx), fmtNum_(np.tx)],
    ['Avg kg / transaction', fmtNum_(p.avgKgPerTx), fmtNum_(np.avgKgPerTx)],
    ['Avg tonnes / day', fmtTonnes_(Math.round(p.kg / report.days)), fmtTonnes_(Math.round(np.kg / report.days))]
  ];
  if (hasCutting) rows.push(['Avg cutting ratio', fmtPct_(p.avgCuttingRatioPct, 2), fmtPct_(np.avgCuttingRatioPct, 2)]);
  const rowTop = top + 13, rowH = (box.y + box.h - 6 - rowTop) / rows.length;
  rows.forEach((r, i) => {
    const y = rowTop + i * rowH;
    if (i % 2 === 0) rect_(slide, SlidesApp.ShapeType.RECTANGLE, t.x - 3, y, t.w + 3, rowH, C.cardBg);
    addText_(slide, r[0], t.x, y, colW.label, rowH, { fontSize: 6, color: C.ink, valign: 'MIDDLE' });
    [r[1], r[2]].forEach((v, k) => {
      addText_(slide, v, t.x + colW.label + k * colW.p, y, colW.p - 4, rowH,
        { fontSize: 6.5, bold: true, color: k === 0 ? C.greenDark : C.ink, align: 'END', valign: 'MIDDLE' });
    });
  });
}

/** Weekly (Thu–Wed) tonnage, stacked partner / non-partner, with partner share under each week. */
function drawFfbWeeklyChart_(slide, box, weeks) {
  const C = CORE.COLORS;
  panel_(slide, box);
  sectionTitle_(slide, 'WEEKLY TONNAGE TREND (THU–WED)', box.x + 10, box.y + 8, 200);
  let lx = box.x + box.w - 10;
  [['Non-partner', C.nonPartner], ['Partner', C.greenDark]].forEach(item => {
    const tw = estimateTextWidth_(item[0], 6) + 3;
    lx -= tw;
    addText_(slide, item[0], lx, box.y + 8, tw, 11, { fontSize: 6, color: C.ink, valign: 'MIDDLE' });
    lx -= 10;
    rect_(slide, SlidesApp.ShapeType.RECTANGLE, lx, box.y + 10, 7, 7, item[1]);
    lx -= 9;
  });
  if (!weeks.length) return;

  const plot = { x: box.x + 34, y: box.y + 32, w: box.w - 44, h: box.h - 32 - 22 };
  const max = niceMax_(Math.max(1, ...weeks.map(w => w.totalKg)));
  [0, 0.5, 1].forEach(fr => {
    const gy = plot.y + plot.h - fr * plot.h;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, fr === 0 ? C.border : C.grid, fr === 0 ? 0.75 : 0.5);
    if (fr > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(max * fr / 1000, 0) + ' t', box.x + 4, gy - 4, plot.x - box.x - 8, 8, { fontSize: 5.5, color: C.muted, align: 'END', valign: 'MIDDLE' });
  });
  const band = plot.w / weeks.length;
  const barW = Math.min(28, band * 0.55);
  weeks.forEach((w, i) => {
    const cx = plot.x + (i + 0.5) * band;
    let y = plot.y + plot.h;
    [[w.PARTNER.kg, C.greenDark], [w.NON_PARTNER.kg, C.nonPartner]].forEach(sg => {
      const h = (sg[0] / max) * plot.h;
      if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, cx - barW / 2, y - h, barW, h, sg[1]);
      y -= h;
    });
    addText_(slide, fmtTonnes_(w.totalKg), cx - band / 2, y - 9, band, 8, { fontSize: 5.5, bold: true, color: C.ink, align: 'CENTER', valign: 'MIDDLE' });
    addRichText_(slide, [
      { text: formatRange_(w.start, w.end), color: C.muted },
      { text: '\nP ' + fmtPct_(w.partnerSharePct, 1), color: C.greenDark, bold: true }
    ], cx - band / 2, plot.y + plot.h + 3, band, 16, { fontSize: 5.5, align: 'CENTER' });
  });
}
