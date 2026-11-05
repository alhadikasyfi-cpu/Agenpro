/**
 * SawitPRO Feature Analysis — slide components shared by the Feature_*.gs files:
 * title row, KPI cards, step funnel, daily trend chart, AI panels, and createFunnelFeature_() for features whose
 * sheet is a daily funnel (app DAU → feature steps, with new / returning users, user names and weekly WAU).
 */

/* ------------------------------ FUNNEL FEATURES ------------------------------ */

/**
 * Builds a feature from a funnel sheet. def = {
 *   id, name, menuFunction, sheet, description,
 *   app: 'timpro' | 'agenpro'          — fixed app, or
 *   splitByColumn: true                — one slide per app found in the "segment" column (e.g. "Lapak Sawit Agenpro"),
 *   fields: { date, segment?, dau, <step keys>, newUsers?, returningUsers?, names?, appWau?, featureWau? } → header aliases,
 *   steps: [{ key, label }]            — feature steps in order; the first is "users of the feature",
 *   usersLabel                         — KPI label for the first step's users, e.g. 'USERS OPEN DASHBOARD'
 * }
 * One step → funnel App WAU → feature users → returning users. Several steps → weekly totals through every step.
 */
function createFunnelFeature_(def) {
  return {
    id: def.id,
    name: def.name,
    menuFunction: def.menuFunction,
    buildSlides: (ss, issues) => {
      const stepKeys = def.steps.map(s => s.key);
      const table = readTable_(ss, def.sheet, def.fields, ['date', 'dau', stepKeys[0]]);
      const groups = {};
      table.rows.forEach(r => {
        const appKey = def.splitByColumn ? appKeyFromText_(r.segment, def.name) : def.app;
        (groups[appKey] = groups[appKey] || []).push(r);
      });
      return Object.keys(groups).filter(k => CORE.APPS[k]).sort().map(appKey => {
        const app = CORE.APPS[appKey];
        const tag = app.label + ' ';
        const series = toDailySeries_(groups[appKey], stepKeys, tag, issues);
        const weekly = weeklyAnalysis_(series, stepKeys, tag, issues);
        if (!weekly) throw new Error(app.label + ': needs at least one complete week (Thu–Wed) of data');
        if (weekly.daysInProgress) issues.push(tag + weekly.daysInProgress + ' day(s) of the current week wait until the week is complete');
        const title = def.name + ' · ' + app.label;
        const cmp = weekly.compare;
        return {
          title: title,
          ai: {
            feature: def.name,
            app: app,
            description: def.description,
            scope: [
              'Weekly review of ' + def.name + ' in ' + app.label + ': this week vs last week (weeks run Thursday–Wednesday).',
              def.steps.length > 1
                ? 'Explain the funnel (app DAU → ' + def.steps.map(s => s.label).join(' → ') + '), where users drop off, and what changed week over week.'
                : 'Explain the funnel (app WAU → users of ' + def.name + ' → returning users) and what changed week over week.'
            ],
            definitions: weeklyDefinitions_(def),
            payload: {
              weeks: weekly.weeks,
              thisWeekVsLastWeek: cmp,
              daily: series.filter(r => r.date <= weekly.windows[weekly.windows.length - 1].end).slice(-14)
                .map(r => Object.assign({ date: r.date, dau: r.dau, newUsers: r.newUsers, returningUsers: r.returningUsers, ctrPct: r.ctrPct }, r.steps)),
              dataQualityIssues: issues.filter(i => i.indexOf(tag) === 0)
            },
            fallbackHeadline: title + ': ' + fmtNum_(cmp.current.clickUsers) + ' users this week.'
          },
          draw: (slide, ai) => drawFunnelFeatureSlide_(slide, def, app, series, weekly, ai, issues.filter(i => i.indexOf(tag) === 0).length)
        };
      });
    }
  };
}

function appKeyFromText_(text, featureName) {
  const rest = String(text || '').toLowerCase().replace(featureName.toLowerCase(), '').replace(/[^a-z0-9]+/g, '');
  return CORE.APPS[rest] ? rest : (/agen/.test(rest) ? 'agenpro' : /tim/.test(rest) ? 'timpro' : rest);
}

function weeklyDefinitions_(def) {
  const d = {
    appWau: 'app weekly active users',
    clickUsers: 'distinct users who used ' + def.name + ' (' + def.steps[0].label + ') during the week (NOT WAU)',
    conversionPct: 'clickUsers ÷ appWau',
    totalDau: 'app DAU summed over the week',
    ctrPct: 'daily ' + def.steps[0].label + ' users summed ÷ DAU summed',
    newUsers: 'users who used it for the first time ever during the week',
    returningUsers: 'clickUsers minus newUsers',
    retainedFromPrev: 'users active in both this and the previous week; retentionPct = retainedFromPrev ÷ previous week users',
    churnedFromPrev: 'previous-week users who did not come back this week',
    repeatInWeek: 'users active on 2+ days within the week',
    appStickinessPct: 'average app DAU ÷ app WAU'
  };
  if (def.steps.length > 1) {
    def.steps.forEach(s => d['step_' + s.key] = s.label + ': daily counts summed over the week');
    d.completionPct = def.steps[def.steps.length - 1].label + ' ÷ ' + def.steps[0].label + ' (weekly totals)';
  }
  return d;
}

/** Weekly slide for one app of a funnel feature. */
function drawFunnelFeatureSlide_(slide, def, app, series, weekly, ai, warningCount) {
  const L = CORE.LAYOUT;
  drawTitle_(slide, L.title, { feature: def.name, app: app, suffix: 'Weekly Review', pill: weekPill_(weekly), warnings: warningCount });
  drawKpiCards_(slide, L.kpi, weeklyKpiCards_(weekly.compare, app, def.usersLabel));
  drawStepFunnel_(slide, L.main, app, def.steps.length > 1 ? stepFunnelView_(weekly, app, def) : wauFunnelView_(weekly, app, def));

  // Daily trend of last week + this week
  const w = weekly.windows;
  const range = { from: w[Math.max(0, w.length - 2)].start, to: w[w.length - 1].end, weekStart: w.length > 1 ? w[w.length - 1].start : null };
  const rows = series.filter(r => r.date >= range.from && r.date <= range.to);
  const hasSplit = rows.every(r => r.newUsers != null && r.returningUsers != null);
  const lastStep = def.steps.length > 1 ? def.steps[def.steps.length - 1] : null;
  const C = CORE.COLORS;
  drawDailyTrend_(slide, L.chart, {
    title: 'DAILY TREND · ' + app.label.toUpperCase(),
    rows: rows.map(r => ({
      date: r.date,
      line: r.dau,
      segs: hasSplit ? [[r.newUsers, app.color], [r.returningUsers, C.returning]] : [[r.steps[def.steps[0].key], app.color]],
      total: r.steps[def.steps[0].key],
      inner: lastStep ? r.steps[lastStep.key] : null,
      pct: r.ctrPct
    })),
    legend: [{ label: 'DAU (left)', color: C.blue, kind: 'line' }]
      .concat(hasSplit ? [{ label: 'New', color: app.color, kind: 'bar' }, { label: 'Returning', color: C.returning, kind: 'bar' }]
        : [{ label: def.steps[0].label, color: app.color, kind: 'bar' }])
      .concat(lastStep ? [{ label: lastStep.label, color: C.ink, kind: 'bar' }] : [])
      .concat([{ label: '% = CTR', color: C.greenDark, kind: 'text' }]),
    rightColor: app.color,
    divider: range.weekStart ? { date: range.weekStart, before: 'Last week', after: 'This week' } : null
  });
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

function weekPill_(weekly) {
  const w = weekly.windows;
  const cur = w[w.length - 1], prev = w.length > 1 ? w[w.length - 2] : null;
  return 'THIS WEEK  ' + formatRange_(cur.start, cur.end) + (prev ? '   vs   ' + formatRange_(prev.start, prev.end) : '');
}

/** KPI cards: app WAU, users of the feature, WAU → feature conversion (or CTR), retained from last week. */
function weeklyKpiCards_(cmp, app, usersLabel) {
  const C = CORE.COLORS;
  const cur = cmp.current, prev = cmp.previous;
  const sub = (key, prevText) => {
    if (!prev) return { sub: 'first week in the data', subColor: null };
    const st = deltaStyle_(cmp.deltas[key], metricKind_(key));
    return { sub: st.text + ' vs ' + prevText + ' last week', subColor: st.ink };
  };
  const hasWau = cur.appWau != null;
  return [
    hasWau
      ? Object.assign({ label: app.label.toUpperCase() + ' · APP WAU', value: fmtNum_(cur.appWau), accent: C.blue }, sub('appWau', fmtNum_(prev && prev.appWau)))
      : { label: app.label.toUpperCase() + ' · APP WAU', value: '–', sub: '⚠ not in sheet · avg DAU ' + fmtNum_(cur.avgDau) + '/day', subColor: C.amber, accent: C.blue },
    Object.assign({ label: usersLabel, value: fmtNum_(cur.clickUsers), accent: app.color }, sub('clickUsers', fmtNum_(prev && prev.clickUsers))),
    hasWau
      ? Object.assign({ label: 'CONVERSION · WAU → FEATURE', value: fmtPct_(cur.conversionPct, 2), accent: C.greenBar }, sub('conversionPct', fmtPct_(prev && prev.conversionPct, 2)))
      : Object.assign({ label: 'CTR · VISITS ÷ DAU', value: fmtPct_(cur.ctrPct, 2), accent: C.greenBar }, sub('ctrPct', fmtPct_(prev && prev.ctrPct, 2))),
    prev && cur.retainedFromPrev != null ? {
      label: 'RETAINED FROM LAST WEEK',
      value: fmtPct_(cur.retentionPct, 0),
      sub: fmtNum_(cur.retainedFromPrev) + ' of ' + fmtNum_(cur.prevClickUsers) + ' last-week users came back',
      accent: C.returning
    } : Object.assign({ label: 'NEW USERS (FIRST TIME)', value: fmtNum_(cur.newUsers), accent: C.returning }, sub('newUsers', fmtNum_(prev && prev.newUsers)))
  ];
}

/** One-step feature: app WAU → users of the feature → returning users, this week with last week alongside. */
function wauFunnelView_(weekly, app, def) {
  const C = CORE.COLORS;
  const cmp = weekly.compare;
  const cur = cmp.current, prev = cmp.previous;
  const hasWau = cur.appWau != null;
  const top = hasWau ? cur.appWau : cur.avgDau;
  const conv = hasWau ? cur.conversionPct : cur.ctrPct;
  const convKey = hasWau ? 'conversionPct' : 'ctrPct';
  const lastWk = v => (prev ? 'last wk ' + v : null);
  const steps = [
    { name: hasWau ? 'Weekly Active Users' : 'App DAU (avg/day)', sub: lastWk(fmtNum_(prev && (hasWau ? prev.appWau : prev.avgDau))),
      value: top, ratio: 1, color: app.soft, valueColor: app.color, note: hasWau ? 'app users this week' : 'app WAU not in sheet' },
    { name: def.steps[0].label, sub: lastWk(fmtNum_(prev && prev.clickUsers)), value: cur.clickUsers,
      ratio: hasWau && top ? cur.clickUsers / top : (cur.ctrPct || 0) / 100, color: app.color,
      note: hasWau ? fmtPct_(conv, 2) + ' of WAU' : 'users · CTR ' + fmtPct_(cur.ctrPct, 2) }
  ];
  if (cur.returningUsers != null) {
    steps.push({ name: 'Kembali lagi', sub: lastWk(fmtNum_(prev && prev.returningUsers)), value: cur.returningUsers,
      ratio: top ? cur.returningUsers / top : 0, color: C.returning,
      note: fmtPct_(cur.clickUsers ? (cur.returningUsers / cur.clickUsers) * 100 : null, 0) + ' had used it before' });
  }
  const gaps = [];
  if (conv != null) gaps.push([{ text: '▼ ' + fmtPct_(100 - conv, 2), bold: true, color: C.red }, { text: ' of ' + (hasWau ? 'app WAU' : 'DAU') + ' did not use ' + def.name + ' this week', color: C.muted }]);
  if (cur.newUsers != null && cur.clickUsers) {
    gaps.push([{ text: fmtNum_(cur.newUsers) + ' new', bold: true, color: app.color }, { text: ' (' + fmtPct_((cur.newUsers / cur.clickUsers) * 100, 0) + ') used it for the first time', color: C.muted }]);
  }
  return {
    title: 'WEEKLY FUNNEL  ·  ' + (hasWau ? 'WAU' : 'DAU') + ' → ' + def.steps[0].label.toUpperCase() + ' → KEMBALI',
    subtitle: 'This week ' + formatRange_(cur.start, cur.end),
    steps: steps,
    gaps: gaps,
    stats: retentionStats_(cmp),
    callout: {
      label: hasWau ? 'WAU → FEATURE CONVERSION' : 'CLICK-THROUGH RATE',
      value: conv,
      sub: hasWau ? 'of ' + app.label + ' WAU used ' + def.name : 'visits ÷ ' + app.label + ' DAU',
      trendPp: cmp.deltas[convKey] ? cmp.deltas[convKey].diff : null,
      trendLabel: prev ? 'vs last week (' + fmtPct_(prev[convKey], 2) + ')' : ''
    }
  };
}

/** Multi-step feature: weekly totals (daily counts summed) from app DAU through every step, with step conversion. */
function stepFunnelView_(weekly, app, def) {
  const C = CORE.COLORS;
  const cmp = weekly.compare;
  const cur = cmp.current, prev = cmp.previous;
  const lastWk = v => (prev ? 'last wk ' + v : null);
  const total = cur.totalDau;
  const steps = [{ name: 'App DAU (7 days)', sub: lastWk(fmtNum_(prev && prev.totalDau)), value: total, ratio: 1, color: app.soft, valueColor: app.color, note: 'daily users summed' }];
  def.steps.forEach((s, i) => {
    const v = cur['step_' + s.key];
    const before = i === 0 ? total : cur['step_' + def.steps[i - 1].key];
    steps.push({
      name: s.label, sub: lastWk(fmtNum_(prev && prev['step_' + s.key])), value: v,
      ratio: total ? v / total : 0, color: i === def.steps.length - 1 ? C.ink : app.color, valueColor: C.white,
      note: fmtPct_(before ? (v / before) * 100 : null, 1) + ' of previous step'
    });
  });
  const gaps = def.steps.map((s, i) => {
    const before = i === 0 ? total : cur['step_' + def.steps[i - 1].key];
    const v = cur['step_' + s.key];
    return [{ text: '▼ ' + fmtPct_(before ? 100 - (v / before) * 100 : null, 1), bold: true, color: C.red }, { text: ' dropped before "' + s.label + '"', color: C.muted }];
  });
  const first = def.steps[0], last = def.steps[def.steps.length - 1];
  return {
    title: 'WEEKLY FUNNEL  ·  ' + ['DAU'].concat(def.steps.map(s => s.short || s.label)).join(' → ').toUpperCase(),
    subtitle: 'Weekly totals · ' + formatRange_(cur.start, cur.end),
    steps: steps,
    gaps: gaps,
    stats: retentionStats_(cmp),
    callout: {
      label: 'COMPLETION RATE',
      value: cur.completionPct,
      sub: (last.short || last.label) + ' ÷ ' + (first.short || first.label),
      trendPp: cmp.deltas.completionPct ? cmp.deltas.completionPct.diff : null,
      trendLabel: prev ? 'vs last week (' + fmtPct_(prev.completionPct, 2) + ')' : ''
    }
  };
}

function retentionStats_(cmp) {
  const C = CORE.COLORS;
  const cur = cmp.current, prev = cmp.previous;
  if (!prev || cur.retainedFromPrev == null) return null;
  return [
    { text: 'From last week\'s ' + fmtNum_(cur.prevClickUsers) + ' users:  ', color: C.muted },
    { text: fmtNum_(cur.retainedFromPrev) + ' came back (' + fmtPct_(cur.retentionPct, 0) + ')', bold: true, color: C.okInk },
    { text: '   ·   ', color: C.muted },
    { text: fmtNum_(cur.churnedFromPrev) + ' did not return', bold: true, color: C.red }
  ];
}

/* ------------------------------ SLIDE COMPONENTS ------------------------------ */

/** Title row: "<feature> · <app>  <suffix>", period pill on the right, red warning count when the data has issues. */
function drawTitle_(slide, box, t) {
  const C = CORE.COLORS;
  const parts = [{ text: t.feature, color: C.greenDark }];
  if (t.app) parts.push({ text: '  ·  ', color: C.muted }, { text: t.app.label, color: t.app.color });
  parts.push({ text: '   ' + t.suffix, color: C.muted, bold: false });
  addRichText_(slide, parts, box.x, box.y, 400, box.h, { fontSize: 12, bold: true, valign: 'MIDDLE' });

  const pillW = estimateTextWidth_(t.pill, 6.5, true) + 18;
  const pillX = box.x + box.w - pillW;
  pill_(slide, t.pill, pillX, box.y + 2, pillW, 14, C.greenSoft, C.greenDark, 6.5);
  if (t.warnings) {
    const warn = '⚠ ' + t.warnings + ' data warning' + (t.warnings > 1 ? 's' : '') + ' — see Run Log';
    const warnW = estimateTextWidth_(warn, 6, true) + 4;
    addText_(slide, warn, pillX - warnW - 8, box.y + 2, warnW, 14, { fontSize: 6, bold: true, color: C.red, align: 'END', valign: 'MIDDLE' });
  }
}

function drawKpiCards_(slide, box, cards) {
  const gap = 8;
  const w = (box.w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => drawKpiCard_(slide, box.x + i * (w + gap), box.y, w, box.h, card));
}

/**
 * Funnel panel: steps on one shared scale (each with an optional "last wk" line), drop-off notes between steps when
 * there is room, an optional stats line, and a callout on the right (value, trend pill, benchmark).
 */
function drawStepFunnel_(slide, box, app, view) {
  const C = CORE.COLORS;
  const steps = view.steps;
  const n = steps.length;

  panel_(slide, box);
  sectionTitle_(slide, view.title, box.x + 10, box.y + 8, 290);
  if (view.subtitle) addText_(slide, view.subtitle, box.x + box.w - 130, box.y + 8, 120, 11, { fontSize: 6, color: C.muted, align: 'END', valign: 'MIDDLE' });

  const callout = { x: box.x + box.w - 118, y: box.y + 26, w: 108, h: box.h - 34 };
  const col = { badge: box.x + 10, label: box.x + 28, track: box.x + 116 };
  col.labelW = col.track - col.label - 6;
  col.trackW = callout.x - 12 - col.track;

  const rowH = n >= 4 ? 14 : 16;
  const top = box.y + 28;
  const bottomReserve = view.stats ? 22 : 8;
  const step = n > 1 ? Math.min(36, (box.h - 28 - bottomReserve - rowH) / (n - 1)) : 0;
  const rowY = i => top + i * step;
  if (n > 1) {
    const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, col.badge + 6, rowY(0) + rowH / 2, col.badge + 6, rowY(n - 1) + rowH / 2);
    connector.getLineFill().setSolidFill(C.border);
    connector.setWeight(1);
  }

  steps.forEach((st, i) => {
    const y = rowY(i), cy = y + rowH / 2;
    rect_(slide, SlidesApp.ShapeType.ELLIPSE, col.badge, cy - 6, 12, 12, i === 0 ? app.color : st.color, C.white, 1);
    addText_(slide, String(i + 1), col.badge, cy - 6, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });
    const nameSize = fitFontSize_([st.name], col.labelW, 9, 7, 5.5, 0, true);
    if (st.sub) {
      addRichText_(slide, [{ text: st.name, bold: true, color: C.ink }, { text: '\n' + st.sub, color: C.muted, bold: false, size: 5.5 }],
        col.label, y - 2, col.labelW, rowH + 4, { fontSize: nameSize, valign: 'MIDDLE' });
    } else {
      addText_(slide, st.name, col.label, y, col.labelW, rowH, { fontSize: nameSize, bold: true, color: C.ink, valign: 'MIDDLE' });
    }

    const barH = rowH - 4;
    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - barH / 2, col.trackW, barH, C.track);
    const fillW = st.value > 0 ? Math.max(4, col.trackW * Math.min(1, st.ratio)) : 0;
    if (fillW) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - barH / 2, fillW, barH, st.color);
    if (fillW >= 40) {
      addText_(slide, fmtNum_(st.value, 1), col.track, y, fillW - 6, rowH, { fontSize: 7, bold: true, color: st.valueColor || C.white, align: 'END', valign: 'MIDDLE' });
      if (st.note && fillW < col.trackW - 60) addText_(slide, st.note, col.track + fillW + 5, y, col.trackW - fillW - 5, rowH, { fontSize: 6, color: C.muted, valign: 'MIDDLE' });
    } else {
      const ox = col.track + fillW + 5;
      addRichText_(slide, [{ text: fmtNum_(st.value, 1), bold: true, color: C.ink }, { text: '  ·  ' + st.note, color: C.muted }],
        ox, y, col.track + col.trackW - ox, rowH, { fontSize: 6.5, valign: 'MIDDLE' });
    }
  });

  // Drop-off notes between steps, when the spacing leaves room for a line of text
  if (step - rowH >= 8) {
    (view.gaps || []).forEach((parts, i) => {
      if (i < n - 1) addRichText_(slide, parts, col.track, rowY(i) + rowH + 1, col.trackW, step - rowH - 2, { fontSize: 5.5, valign: 'MIDDLE' });
    });
  }

  if (view.stats) {
    const statsY = box.y + box.h - 15;
    hLine_(slide, box.x + 10, callout.x - 12, statsY - 3, C.border, 0.5);
    addRichText_(slide, view.stats, box.x + 10, statsY, callout.x - 22 - box.x, 10, { fontSize: 6, valign: 'MIDDLE' });
  }

  const c = view.callout;
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, callout.x, callout.y, callout.w, callout.h, app.soft);
  addText_(slide, c.label, callout.x + 6, callout.y + 7, callout.w - 12, 8, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
  addText_(slide, fmtPct_(c.value, 2), callout.x + 6, callout.y + 17, callout.w - 12, 22, { fontSize: 19, bold: true, color: app.color, align: 'CENTER', valign: 'MIDDLE' });
  addText_(slide, c.sub, callout.x + 6, callout.y + 41, callout.w - 12, 8, { fontSize: 5.5, color: C.muted, align: 'CENTER' });
  if (c.trendPp != null) {
    const pp = c.trendPp, up = pp > 0, flat = pp === 0;
    pill_(slide, (flat ? '■ ' : up ? '▲ +' : '▼ −') + Math.abs(pp).toFixed(2) + ' pp', callout.x + 14, callout.y + 53, callout.w - 28, 12,
      flat ? C.cardBg : up ? C.okSoft : C.redSoft, flat ? C.muted : up ? C.okInk : C.red, 6.5);
    addText_(slide, c.trendLabel, callout.x + 6, callout.y + 67, callout.w - 12, 7, { fontSize: 5, color: C.muted, align: 'CENTER' });
  }
  if (c.benchmark) addText_(slide, c.benchmark, callout.x + 6, callout.y + callout.h - 12, callout.w - 12, 7, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
}

/**
 * Daily chart: a line on the left axis (app DAU) and stacked bars on the right axis, an optional dark inner bar (a later
 * funnel step) and % under each date. Optional week divider (shaded "this week") or event marker (vertical line).
 * spec = { title, rows: [{ date, line, segs: [[value, color]], total, inner, pct }], legend, rightColor, divider?, marker? }
 */
function drawDailyTrend_(slide, box, spec) {
  const C = CORE.COLORS;
  const rows = spec.rows;
  const n = rows.length;
  const detailed = n <= 14;

  panel_(slide, box);
  sectionTitle_(slide, spec.title, box.x + 10, box.y + 8, 150);

  let lx = box.x + box.w - 10;
  spec.legend.filter(item => item.kind !== 'text' || detailed).slice().reverse().forEach(item => {
    const tw = estimateTextWidth_(item.label, 6) + 3;
    lx -= tw;
    addText_(slide, item.label, lx, box.y + 8, tw, 11, { fontSize: 6, color: item.kind === 'text' ? item.color : C.ink, valign: 'MIDDLE' });
    if (item.kind === 'line') {
      lx -= 13;
      hLine_(slide, lx, lx + 10, box.y + 13.5, item.color, 2);
    } else if (item.kind === 'bar') {
      lx -= 10;
      rect_(slide, SlidesApp.ShapeType.RECTANGLE, lx, box.y + 10, 7, 7, item.color);
    }
    lx -= 9;
  });
  if (!n) return;

  const plot = { x: box.x + 34, y: box.y + 30, w: box.w - 34 - 34, h: box.h - 30 - (detailed ? 22 : 14) };
  const band = plot.w / n;
  const px = i => plot.x + (i + 0.5) * band;

  // Week divider / event marker, drawn first so the chart sits on top
  if (spec.divider) {
    const idx = rows.findIndex(r => r.date >= spec.divider.date);
    if (idx > 0) {
      const x = plot.x + idx * band;
      rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, box.y + 22, plot.x + plot.w - x, box.h - 26, C.greenSoft);
      addText_(slide, spec.divider.before, plot.x, box.y + 22, x - plot.x, 7, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
      addText_(slide, spec.divider.after, x, box.y + 22, plot.x + plot.w - x, 7, { fontSize: 5.5, bold: true, color: C.greenDark, align: 'CENTER' });
    }
  }
  if (spec.marker) {
    const idx = rows.findIndex(r => r.date >= spec.marker.date);
    if (idx >= 0) {
      const x = plot.x + idx * band;
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, box.y + 24, x, plot.y + plot.h);
      line.getLineFill().setSolidFill(C.marker);
      line.setWeight(1);
      line.setDashStyle(SlidesApp.DashStyle.DASH);
      addText_(slide, spec.marker.label, x + 3, box.y + 22, 120, 7, { fontSize: 5.5, bold: true, color: C.marker });
    }
  }

  const maxLine = niceMax_(Math.max(1, ...rows.map(r => r.line || 0)));
  const maxBar = niceMax_(Math.max(1, ...rows.map(r => r.total || 0)) * 1.6);
  const yOf = (v, max) => plot.y + plot.h - (v / max) * plot.h;
  [0, 0.5, 1].forEach(fr => {
    const gy = plot.y + plot.h - fr * plot.h;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, fr === 0 ? C.border : C.grid, fr === 0 ? 0.75 : 0.5);
    if (fr > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(maxLine * fr, maxLine < 10 ? 1 : 0), box.x + 6, gy - 4, plot.x - box.x - 10, 8, { fontSize: 5.5, color: C.blue, align: 'END', valign: 'MIDDLE' });
    addText_(slide, fmtNum_(maxBar * fr, 1), plot.x + plot.w + 4, gy - 4, 24, 8, { fontSize: 5.5, color: spec.rightColor, valign: 'MIDDLE' });
  });

  const barW = Math.max(2, Math.min(16, band * 0.6));
  const base = plot.y + plot.h;
  rows.forEach((r, i) => {
    let y = base;
    r.segs.forEach(sg => {
      const h = ((sg[0] || 0) / maxBar) * plot.h;
      if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, px(i) - barW / 2, y - h, barW, h, sg[1]);
      y -= h;
    });
    if (r.inner) {
      const h = (r.inner / maxBar) * plot.h;
      if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, px(i) - barW / 4, base - h, barW / 2, h, C.ink);
    }
    if (detailed) addText_(slide, fmtNum_(r.total), px(i) - 15, y - 9, 30, 8, { fontSize: 5.5, bold: true, color: C.ink, align: 'CENTER', valign: 'MIDDLE' });
  });

  for (let i = 1; i < n; i++) {
    const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, px(i - 1), yOf(rows[i - 1].line, maxLine), px(i), yOf(rows[i].line, maxLine));
    line.getLineFill().setSolidFill(C.blue);
    line.setWeight(1.75);
  }
  if (detailed) {
    rows.forEach((r, i) => {
      const y = yOf(r.line, maxLine);
      rect_(slide, SlidesApp.ShapeType.ELLIPSE, px(i) - 2.25, y - 2.25, 4.5, 4.5, C.blue, C.white, 0.75);
      addText_(slide, fmtNum_(r.line), px(i) - 15, y - 11, 30, 8, { fontSize: 5.5, bold: true, color: C.blue, align: 'CENTER', valign: 'MIDDLE' });
    });
  }

  const every = detailed ? 1 : Math.ceil(n / 8);
  const labelIdx = [];
  for (let i = 0; i < n; i += every) labelIdx.push(i);
  if (labelIdx[labelIdx.length - 1] !== n - 1) {
    if (n - 1 - labelIdx[labelIdx.length - 1] < every * 0.6) labelIdx.pop();
    labelIdx.push(n - 1);
  }
  labelIdx.forEach(i => {
    const parts = [{ text: formatShortDate_(rows[i].date), color: C.muted }];
    if (detailed && rows[i].pct != null) parts.push({ text: '\n' + fmtPct_(rows[i].pct, 2), color: C.greenDark, bold: true });
    addRichText_(slide, parts, px(i) - 22, plot.y + plot.h + 3, 44, detailed ? 16 : 8, { fontSize: 5.5, align: 'CENTER' });
  });
}

/* ------------------------------ DRAWING HELPERS ------------------------------ */

function rect_(slide, type, x, y, w, h, fill, border, weight) {
  const shape = slide.insertShape(type, x, y, w, h);
  shape.getFill().setSolidFill(fill);
  if (border) {
    shape.getBorder().getLineFill().setSolidFill(border);
    shape.getBorder().setWeight(weight || 0.5);
  } else {
    shape.getBorder().setTransparent();
  }
  return shape;
}

function hLine_(slide, x1, x2, y, color, weight) {
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y, x2, y);
  line.getLineFill().setSolidFill(color);
  line.setWeight(weight || 0.5);
  return line;
}

function panel_(slide, box) {
  const C = CORE.COLORS;
  return rect_(slide, SlidesApp.ShapeType.RECTANGLE, box.x, box.y, box.w, box.h, C.panelBg, C.border, 0.75);
}

function sectionTitle_(slide, title, x, y, w) {
  const C = CORE.COLORS;
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + 1.5, 3, 8, C.yellow);
  return addText_(slide, title, x + 7, y, w, 11, { fontSize: 8, bold: true, color: C.greenDark, valign: 'MIDDLE' });
}

function headerPanel_(slide, box, title) {
  const C = CORE.COLORS;
  panel_(slide, box);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, box.x, box.y, box.w, 20, C.greenDark);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, box.x, box.y, 3, 20, C.yellow);
  addText_(slide, title, box.x + 10, box.y, box.w - 70, 20, { fontSize: 8, bold: true, color: C.white, valign: 'MIDDLE' });
  addText_(slide, 'AI insight', box.x + box.w - 60, box.y, 52, 20, { fontSize: 5.5, color: C.greenMuted, align: 'END', valign: 'MIDDLE' });
}

function pill_(slide, text, x, y, w, h, fill, ink, fontSize) {
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h, fill);
  return addText_(slide, text, x, y, w, h, { fontSize: fontSize, bold: true, color: ink, align: 'CENTER', valign: 'MIDDLE' });
}

/** Adds a text box whose *text area* is exactly (x, y, w, h); the box is grown by the built-in insets. */
function addText_(slide, text, x, y, w, h, opts) {
  opts = opts || {};
  const ix = CORE.TEXT_INSET_X, iy = CORE.TEXT_INSET_Y;
  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - ix, y - iy, w + ix * 2, h + iy * 2);
  shape.getFill().setTransparent();
  shape.getBorder().setTransparent();
  shape.setContentAlignment(opts.valign === 'MIDDLE' ? SlidesApp.ContentAlignment.MIDDLE : SlidesApp.ContentAlignment.TOP);

  const str = String(text == null ? '' : text);
  if (!str) return shape;
  const range = shape.getText();
  range.setText(str);
  range.getTextStyle()
    .setFontFamily(CORE.FONT)
    .setFontSize(opts.fontSize || 8)
    .setForegroundColor(opts.color || CORE.COLORS.ink)
    .setBold(!!opts.bold);

  const alignments = {
    START: SlidesApp.ParagraphAlignment.START,
    CENTER: SlidesApp.ParagraphAlignment.CENTER,
    END: SlidesApp.ParagraphAlignment.END
  };
  const para = range.getParagraphStyle();
  para.setParagraphAlignment(alignments[opts.align || 'START']);
  if (opts.spaceBelow) para.setSpaceBelow(opts.spaceBelow);
  return shape;
}

/** Text box made of styled runs: parts = [{ text, color?, bold?, size? }]. */
function addRichText_(slide, parts, x, y, w, h, opts) {
  const shape = addText_(slide, parts.map(p => p.text).join(''), x, y, w, h, opts);
  let pos = 0;
  parts.forEach(p => {
    const end = pos + p.text.length;
    if (p.text.length && (p.bold != null || p.color || p.size)) {
      const style = shape.getText().getRange(pos, end).getTextStyle();
      if (p.bold != null) style.setBold(p.bold);
      if (p.color) style.setForegroundColor(p.color);
      if (p.size) style.setFontSize(p.size);
    }
    pos = end;
  });
  return shape;
}

function estimateTextWidth_(text, size, bold) {
  return String(text || '').length * size * CORE.CHAR_WIDTH * (bold ? 1.1 : 1) + 2;
}

function estimateLines_(text, w, size, bold) {
  const maxChars = Math.max(1, Math.floor(w / (size * CORE.CHAR_WIDTH * (bold ? 1.1 : 1))));
  let lines = 1, used = 0;
  String(text || '').split(/\s+/).filter(Boolean).forEach(word => {
    if (!used) used = word.length;
    else if (used + 1 + word.length <= maxChars) used += 1 + word.length;
    else { lines++; used = word.length; }
    while (used > maxChars) { lines++; used -= maxChars; }
  });
  return lines;
}

/** Largest font size (in 0.5pt steps) at which the paragraphs fit inside w x h. */
function fitFontSize_(paragraphs, w, h, maxSize, minSize, paragraphGap, bold) {
  for (let size = maxSize; size > minSize; size -= 0.5) {
    const height = paragraphs.reduce((acc, p) => acc + estimateLines_(p, w, size, bold) * size * CORE.LINE_HEIGHT, 0) +
      (paragraphGap || 0) * Math.max(0, paragraphs.length - 1);
    if (height <= h) return size;
  }
  return minSize;
}

function niceMax_(value) {
  const magnitude = Math.pow(10, Math.floor(Math.log10(value)));
  const steps = [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10];
  const n = value / magnitude;
  for (let i = 0; i < steps.length; i++) {
    if (n <= steps[i]) return steps[i] * magnitude;
  }
  return 10 * magnitude;
}

function drawKpiCard_(slide, x, y, w, h, card) {
  const C = CORE.COLORS;
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h, C.cardBg, C.border);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + 7, 3, h - 14, card.accent);

  const labelSize = fitFontSize_([card.label], w - 14, 8, 6, 5, 0, true);
  addText_(slide, card.label, x + 10, y + 5, w - 14, 8, { fontSize: labelSize, bold: true, color: C.muted });
  addText_(slide, card.value, x + 10, y + 14, w - 14, 16, { fontSize: 15, bold: true, color: C.ink, valign: 'MIDDLE' });
  const subSize = fitFontSize_([card.sub], w - 14, 7, 5.5, 4.5, 0);
  addText_(slide, card.sub, x + 10, y + 31, w - 14, 7, { fontSize: subSize, color: card.subColor || C.muted });
}

function drawObservations_(slide, ai, box) {
  const C = CORE.COLORS;
  headerPanel_(slide, box, 'EXECUTIVE OBSERVATIONS');

  const innerX = box.x + 10, innerW = box.w - 20;
  let y = box.y + 27;

  // Headline callout
  const hlSize = 7.5;
  const hlLines = Math.min(4, estimateLines_(ai.headline, innerW - 12, hlSize, true));
  const hlH = hlLines * hlSize * CORE.LINE_HEIGHT + 8;
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, innerX, y, innerW, hlH, C.greenSoft);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, innerX, y, 2.5, hlH, C.yellow);
  addText_(slide, ai.headline, innerX + 8, y + 4, innerW - 12, hlH - 8, { fontSize: hlSize, bold: true, color: C.greenDark });
  y += hlH + 6;

  if (!ai.observations.length) return;
  const avail = box.y + box.h - 7 - y;
  const listIndent = 14;
  const size = fitFontSize_(ai.observations, innerW - listIndent, avail, 8, 5.5, 4);
  const shape = addText_(slide, ai.observations.join('\n'), innerX, y, innerW, avail, { fontSize: size, color: C.ink, spaceBelow: 4 });
  shape.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
}

function drawNextSteps_(slide, ai, box) {
  const C = CORE.COLORS;
  headerPanel_(slide, box, 'ACTION PLAN / PIC');

  const steps = ai.nextSteps;
  if (!steps.length) return;

  const badgeX = box.x + 10;
  const textX = box.x + 28;
  const textW = box.x + box.w - 10 - textX;
  const top = box.y + 28;
  const avail = box.y + box.h - 6 - top;
  const gap = 6;

  const itemHeight = (step, size) => estimateLines_(step.action, textW, size) * size * CORE.LINE_HEIGHT + 2 + (size + 4);
  const totalHeight = size => steps.reduce((acc, step) => acc + itemHeight(step, size) + gap, -gap);
  let size = 7.5;
  while (size > 5.5 && totalHeight(size) > avail) size -= 0.5;

  let y = top;
  steps.forEach((step, i) => {
    if (i > 0) hLine_(slide, textX, box.x + box.w - 10, y - gap / 2, C.border, 0.5);

    rect_(slide, SlidesApp.ShapeType.ELLIPSE, badgeX, y, 12, 12, C.greenDark);
    addText_(slide, String(i + 1), badgeX, y, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });

    const actionH = estimateLines_(step.action, textW, size) * size * CORE.LINE_HEIGHT;
    addText_(slide, step.action, textX, y, textW, actionH, { fontSize: size, color: C.ink });

    const picText = 'PIC · ' + step.pic;
    const picSize = Math.max(5.5, size - 1);
    const picW = Math.min(textW, estimateTextWidth_(picText, picSize, true) + 12);
    pill_(slide, picText, textX, y + actionH + 2, picW, size + 4, C.greenSoft, C.greenDark, picSize);

    y += itemHeight(step, size) + gap;
  });
}

/** "▲ +3 (+12%)" / "▼ −1.19 pp" with matching colours; higher is better for every weekly metric. */
function deltaStyle_(delta, metric) {
  const C = CORE.COLORS;
  if (!delta) return { text: 'n/a', ink: C.muted, fill: C.cardBg };
  const d = delta.diff;
  const arrow = d > 0 ? '▲ +' : d < 0 ? '▼ −' : '■ ';
  let text = arrow + (metric.kind === 'pct' ? Math.abs(d).toFixed(2) + ' pp' : fmtNum_(Math.abs(d), metric.digits || 0));
  if (metric.kind !== 'pct' && delta.relPct != null && d !== 0) text += ' (' + (delta.relPct > 0 ? '+' : '−') + fmtNum_(Math.abs(delta.relPct), 0) + '%)';
  return d > 0 ? { text: text, ink: C.okInk, fill: C.okSoft } : d < 0 ? { text: text, ink: C.red, fill: C.redSoft } : { text: text, ink: C.muted, fill: C.cardBg };
}

