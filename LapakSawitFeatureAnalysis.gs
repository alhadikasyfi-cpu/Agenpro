/**
 * Lapak Sawit Feature Analysis — Google Apps Script
 * Duplicates the SawitPRO master template (slide 1) and renders a feature analysis snapshot:
 * KPI cards, DAU → feature funnel, daily trend chart and Gemini-generated insights.
 *
 * Source: sheet "Daily Metrics", one row per date (and feature). Rows are read from row 2 down to
 * the first empty row; columns are matched by header name, so their order does not matter.
 */

const LS_CONFIG = {
  GEMINI_MODEL: 'gemini-3.5-flash-lite',
  PRESENTATION_PROPERTY: 'LAPAK_SAWIT_PRESENTATION_ID',
  API_KEY_PROPERTY: 'GEMINI_API_KEY',
  SHEET_DAILY: 'Daily Metrics',
  SHEET_LOG: 'Run Log',
  FEATURE_NAME: 'Lapak Sawit',
  FONT: 'Arial',

  // Accepted header names per field (lower-case). Optional fields may be absent.
  DAILY_COLUMNS: {
    date: ['tanggal', 'date'],
    feature: ['feature_name', 'feature'],
    dau: ['daily_active_users', 'dau'],
    featureUsers: ['feature_unique_users', 'feature_users', 'unique_users'],
    penetration: ['penetration_rate_pct', 'penetration_pct', 'penetration']
  },
  REQUIRED_COLUMNS: ['date', 'dau', 'featureUsers'],

  // Slides text boxes carry built-in padding that SlidesApp cannot change.
  // addText_ expands every box by these insets so text lands exactly on the layout grid.
  TEXT_INSET_X: 7.2,
  TEXT_INSET_Y: 3.6,
  // Average Arial glyph width as a fraction of font size; used to estimate wrapping and auto-fit text.
  CHAR_WIDTH: 0.52,
  LINE_HEIGHT: 1.2,

  // Design tokens aligned with SawitPRO branding
  COLORS: {
    greenDark: '#31512B',
    greenMid: '#49663D',
    greenBar: '#6E9A5E',
    greenSoft: '#EEF4EA',
    greenMuted: '#C9DCC1',
    yellow: '#FFD400',
    blue: '#2563EB',
    greenLine: '#16A34A',
    orange: '#F97316',
    red: '#DC2626',
    redSoft: '#FEE2E2',
    redBorder: '#FCA5A5',
    amber: '#B45309',
    amberSoft: '#FEF3C7',
    okInk: '#166534',
    okSoft: '#DCFCE7',
    ink: '#1F2937',
    muted: '#6B7280',
    border: '#E5E7EB',
    grid: '#EDEFF2',
    track: '#F1F4EF',
    panelBg: '#FFFFFF',
    cardBg: '#F8FAFC',
    white: '#FFFFFF'
  },

  // Layout grid in points for a 16:9 slide (720 x 405 pt).
  // Content stays between the template header line (~50pt) and the footer bar (~370pt).
  LAYOUT: {
    title: { x: 20, y: 54, w: 680, h: 18 },
    kpi: { x: 20, y: 78, w: 420, h: 42 },
    funnel: { x: 20, y: 126, w: 420, h: 124 },
    chart: { x: 20, y: 256, w: 420, h: 108 },
    observations: { x: 452, y: 78, w: 248, h: 162 },
    actions: { x: 452, y: 246, w: 248, h: 118 }
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lapak Sawit')
    .addItem('1. Setup spreadsheet tabs', 'setupLapakSawit')
    .addItem('2. Set Gemini API key', 'setGeminiApiKey')
    .addItem('3. Generate slide snapshot', 'generateLapakSawitSlide')
    .addSeparator()
    .addItem('Open generated presentation', 'openLapakSawitPresentation')
    .addToUi();
}

function setupLapakSawit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daily = getOrCreateSheet_(ss, LS_CONFIG.SHEET_DAILY);
  const log = getOrCreateSheet_(ss, LS_CONFIG.SHEET_LOG);

  setupTab_(daily, ['tanggal', 'feature_name', 'daily_active_users', 'feature_unique_users', 'penetration_rate_pct'], [
    [20261018, 'Lapak Sawit', 172, 4, 2.33],
    [20261017, 'Lapak Sawit', 176, 3, 1.7],
    [20261016, 'Lapak Sawit', 234, 10, 4.27],
    [20261015, 'Lapak Sawit', 205, 5, 2.44]
  ]);
  setupTab_(log, ['run_at', 'status', 'message', 'presentation_url'], []);
  [daily, log].forEach(sh => {
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, sh.getLastColumn());
  });

  SpreadsheetApp.getUi().alert('Spreadsheet tabs are ready.');
}

function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Gemini API Key', 'Paste your Gemini API key:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const key = result.getResponseText().trim();
  if (key) {
    PropertiesService.getScriptProperties().setProperty(LS_CONFIG.API_KEY_PROPERTY, key);
    ui.alert('Gemini API key saved.');
  }
}

function openLapakSawitPresentation() {
  const id = PropertiesService.getScriptProperties().getProperty(LS_CONFIG.PRESENTATION_PROPERTY);
  if (!id) {
    SpreadsheetApp.getUi().alert('No presentation generated yet.');
    return;
  }
  const url = 'https://docs.google.com/presentation/d/' + id + '/edit';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<p>Open presentation:</p><p><a target="_blank" href="' + url + '">' + url + '</a></p>').setWidth(460).setHeight(130),
    'Lapak Sawit Presentation'
  );
}

function generateLapakSawitSlide() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Process is running. Please wait.');

  let presentationUrl = '';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rows = readDailyMetrics_(ss);
    const analysisInput = buildAnalysisInput_(rows);
    const ai = callGemini_(analysisInput);

    const presentation = getOrCreatePresentation_();
    const slides = presentation.getSlides();
    if (!slides.length) throw new Error('Presentation missing master template on Slide 1.');

    // 1. Duplicate Master Template Slide
    const newSlide = slides[0].duplicate();

    // 2. Render Slide UI Layout
    drawAnalysisSlide_(newSlide, analysisInput, ai);

    // 3. Inject Detailed Speaker Notes
    if (ai.speakerNotes) {
      newSlide.getNotesPage().getSpeakerNotesShape().getText().setText(ai.speakerNotes);
    }

    presentationUrl = 'https://docs.google.com/presentation/d/' + presentation.getId() + '/edit';
    const issues = analysisInput.dataQualityIssues;
    const message = 'Slide generated (' + analysisInput.dataPeriod.days + ' days).' +
      (issues.length ? ' Data warnings: ' + issues.join('; ') : '');
    logRun_(ss, issues.length ? 'SUCCESS_WITH_WARNINGS' : 'SUCCESS', message, presentationUrl);

    SpreadsheetApp.getUi().alert('Slide generated successfully!' +
      (issues.length ? '\n\n' + issues.length + ' data warning(s), see Run Log.' : '') +
      '\n\nOpen presentation:\n' + presentationUrl);
  } catch (err) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) logRun_(ss, 'ERROR', String(err && err.message ? err.message : err), presentationUrl);
    } catch (_) {}
    SpreadsheetApp.getUi().alert('Error: ' + (err && err.message ? err.message : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ DATA PARSING ------------------------------ */

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupTab_(sheet, headers, sampleRows) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (sampleRows.length) sheet.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  }
  sheet.getRange(1, 1, 1, headers.length).setBackground(LS_CONFIG.COLORS.greenDark).setFontColor('#FFFFFF').setFontWeight('bold');
}

/** Reads "Daily Metrics" from row 2 down to the first fully empty row; columns are located by header name. */
function readDailyMetrics_(ss) {
  const sheetName = LS_CONFIG.SHEET_DAILY;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet "' + sheetName + '".');
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('Sheet "' + sheetName + '" has no data rows.');

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const col = {};
  Object.keys(LS_CONFIG.DAILY_COLUMNS).forEach(field => {
    const names = LS_CONFIG.DAILY_COLUMNS[field];
    col[field] = headers.findIndex(h => names.indexOf(h) >= 0);
  });
  LS_CONFIG.REQUIRED_COLUMNS.forEach(field => {
    if (col[field] < 0) {
      throw new Error('Sheet "' + sheetName + '" is missing column "' + LS_CONFIG.DAILY_COLUMNS[field].join('" / "') + '".');
    }
  });

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === '' || v === null)) break; // data ends at the first empty row
    rows.push({
      rowNumber: r + 1,
      date: parseDate_(row[col.date]),
      feature: col.feature >= 0 ? String(row[col.feature] || '').trim() : '',
      dau: toNumber_(row[col.dau]),
      featureUsers: toNumber_(row[col.featureUsers]),
      penetration: col.penetration >= 0 && row[col.penetration] !== '' ? toNumber_(row[col.penetration]) : null
    });
  }
  if (!rows.length) throw new Error('Sheet "' + sheetName + '" has no data before the first empty row.');
  return rows;
}

/** Accepts real dates, yyyyMMdd (number or text) and yyyy-MM-dd. */
function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (value === '' || value == null) return null;
  const s = String(value).trim();
  let m = s.match(/^(\d{4})(\d{2})(\d{2})$/) || s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function toNumber_(value) {
  const n = Number(String(value == null ? '' : value).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}

/** Keeps the rows of the configured feature (or the only feature present). */
function selectFeatureRows_(rows) {
  const norm = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const names = rows.map(r => r.feature).filter(Boolean).filter((n, i, all) => all.findIndex(x => norm(x) === norm(n)) === i);
  const target = norm(LS_CONFIG.FEATURE_NAME);
  if (!names.length) return { feature: LS_CONFIG.FEATURE_NAME, rows: rows };
  const match = names.find(n => norm(n) === target);
  if (match) return { feature: match, rows: rows.filter(r => !r.feature || norm(r.feature) === target) };
  if (names.length === 1) return { feature: names[0], rows: rows };
  throw new Error('Feature "' + LS_CONFIG.FEATURE_NAME + '" not found. Features in sheet: ' + names.join(', '));
}

function buildAnalysisInput_(rawRows) {
  const issues = [];
  const selected = selectFeatureRows_(rawRows);

  // One row per date, validated
  const byDate = {};
  selected.rows.forEach(r => {
    if (!r.date) {
      issues.push('Row ' + r.rowNumber + ': unreadable date (row skipped)');
      return;
    }
    const key = dateKey_(r.date);
    if (byDate[key]) {
      issues.push('Duplicate date ' + key + ' (row ' + r.rowNumber + ' skipped)');
      return;
    }
    const penetrationPct = r.dau > 0 ? round2_((r.featureUsers / r.dau) * 100) : null;
    if (r.featureUsers > r.dau) issues.push(key + ': feature users (' + r.featureUsers + ') exceed DAU (' + r.dau + ')');
    if (r.penetration != null && penetrationPct != null && Math.abs(r.penetration - penetrationPct) > 0.05) {
      issues.push(key + ': penetration_rate_pct ' + r.penetration + ' differs from computed ' + penetrationPct);
    }
    byDate[key] = { date: key, dau: r.dau, featureUsers: r.featureUsers, penetrationPct: penetrationPct };
  });

  const series = Object.keys(byDate).sort().map(k => byDate[k]);
  if (!series.length) throw new Error('No valid rows for feature "' + selected.feature + '".');

  const days = series.length;
  const start = series[0].date, end = series[days - 1].date;
  const spanDays = Math.round((parseDate_(end).getTime() - parseDate_(start).getTime()) / 86400000) + 1;
  if (spanDays > days) issues.push((spanDays - days) + ' missing date(s) between ' + start + ' and ' + end);

  const sum = (rows, key) => rows.reduce((acc, r) => acc + r[key], 0);
  const totalDau = sum(series, 'dau');
  const totalFeatureUsers = sum(series, 'featureUsers');
  const penetrationPct = totalDau > 0 ? round2_((totalFeatureUsers / totalDau) * 100) : null;
  const withPen = series.filter(r => r.penetrationPct != null);
  const pick = (rows, better) => rows.reduce((best, r) => (best == null || better(r, best) ? r : best), null);

  return {
    feature: selected.feature,
    funnelDefinition: 'Step 1 = app Daily Active Users (DAU); Step 2 = users who clicked/opened the feature (feature_unique_users). Conversion = penetration = feature users / DAU.',
    unitsNote: 'Totals of feature users are user-days (sum of daily unique users), not unique users across the whole period.',
    dataPeriod: { start: start, end: end, days: days, calendarDays: spanDays },
    summary: {
      avgDau: Math.round(totalDau / days),
      avgFeatureUsers: round1_(totalFeatureUsers / days),
      totalDauUserDays: totalDau,
      totalFeatureUserDays: totalFeatureUsers,
      penetrationPct: penetrationPct,
      avgDailyPenetrationPct: withPen.length ? round2_(sum(withPen, 'penetrationPct') / withPen.length) : null,
      bestDay: pick(withPen, (a, b) => a.penetrationPct > b.penetrationPct),
      worstDay: pick(withPen, (a, b) => a.penetrationPct < b.penetrationPct),
      peakFeatureUsersDay: pick(series, (a, b) => a.featureUsers > b.featureUsers),
      latestDay: series[days - 1],
      previousDay: days > 1 ? series[days - 2] : null,
      trend: compareWindows_(series)
    },
    funnel: {
      steps: [
        { name: 'Daily Active Users', avgPerDay: Math.round(totalDau / days), total: totalDau },
        { name: 'Klik ' + selected.feature, avgPerDay: round1_(totalFeatureUsers / days), total: totalFeatureUsers }
      ],
      conversionPct: penetrationPct,
      dropOffPct: penetrationPct == null ? null : round2_(100 - penetrationPct)
    },
    patterns: buildPatterns_(series),
    dailySeries: series,
    dataQualityIssues: issues
  };
}

/** Recent window vs the window before it: 7 vs 7 days when there is enough data, otherwise half vs half. */
function compareWindows_(series) {
  if (series.length < 2) return null;
  const size = series.length >= 14 ? 7 : Math.floor(series.length / 2);
  const recent = series.slice(-size), prior = series.slice(-2 * size, -size);
  const agg = rows => {
    const dau = rows.reduce((a, r) => a + r.dau, 0), users = rows.reduce((a, r) => a + r.featureUsers, 0);
    return { avgDau: dau / rows.length, avgUsers: users / rows.length, penetration: dau > 0 ? (users / dau) * 100 : null };
  };
  const a = agg(recent), b = agg(prior);
  return {
    label: 'last ' + size + ' vs previous ' + size + ' days',
    windowDays: size,
    recentPenetrationPct: a.penetration == null ? null : round2_(a.penetration),
    priorPenetrationPct: b.penetration == null ? null : round2_(b.penetration),
    penetrationChangePp: a.penetration == null || b.penetration == null ? null : round2_(a.penetration - b.penetration),
    featureUsersChangePct: pctChange_(a.avgUsers, b.avgUsers),
    dauChangePct: pctChange_(a.avgDau, b.avgDau)
  };
}

/** Statistical patterns that are only reported once there is enough data to be meaningful. */
function buildPatterns_(series) {
  const pens = series.filter(r => r.penetrationPct != null);
  const out = {};
  if (pens.length >= 7) {
    const mean = pens.reduce((a, r) => a + r.penetrationPct, 0) / pens.length;
    const sd = Math.sqrt(pens.reduce((a, r) => a + Math.pow(r.penetrationPct - mean, 2), 0) / pens.length);
    out.penetrationStdDevPp = round2_(sd);
    out.anomalies = sd > 0 ? pens.filter(r => Math.abs(r.penetrationPct - mean) > 2 * sd).map(r => ({ date: r.date, penetrationPct: r.penetrationPct })) : [];
  }
  if (series.length >= 10) {
    out.dauVsFeatureUsersCorrelation = round2_(correlation_(series.map(r => r.dau), series.map(r => r.featureUsers)));
  }
  if (series.length >= 14) {
    const names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const buckets = names.map(() => ({ dau: 0, users: 0 }));
    series.forEach(r => {
      const b = buckets[parseDate_(r.date).getDay()];
      b.dau += r.dau;
      b.users += r.featureUsers;
    });
    out.penetrationByWeekday = names
      .map((n, i) => ({ day: n, penetrationPct: buckets[i].dau > 0 ? round2_((buckets[i].users / buckets[i].dau) * 100) : null }))
      .filter(d => d.penetrationPct != null);
  }
  return out;
}

function correlation_(xs, ys) {
  const n = xs.length;
  const mx = xs.reduce((a, v) => a + v, 0) / n, my = ys.reduce((a, v) => a + v, 0) / n;
  let num = 0, dx = 0, dy = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - mx) * (ys[i] - my);
    dx += Math.pow(xs[i] - mx, 2);
    dy += Math.pow(ys[i] - my, 2);
  }
  return dx > 0 && dy > 0 ? num / Math.sqrt(dx * dy) : 0;
}

function pctChange_(current, previous) {
  return previous > 0 ? round1_(((current - previous) / previous) * 100) : null;
}

function dateKey_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

/* ------------------------------ GEMINI API ------------------------------ */

function callGemini_(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(LS_CONFIG.API_KEY_PROPERTY);
  if (!apiKey) throw new Error('Gemini API key missing in Script Properties.');

  const prompt = [
    'You are a senior product analytics consultant analyzing the "' + data.feature + '" feature of the SawitPRO app (used by palm-oil farmers and agents).',
    'Funnel: ' + data.funnelDefinition,
    'Units: ' + data.unitsNote,
    'Output must be valid JSON matching this schema:',
    '{',
    '  "headline": "Clear single executive takeaway sentence",',
    '  "observations": ["Observation 1", "Observation 2", "Observation 3"],',
    '  "next_steps": [{"action": "Actionable item description", "pic": "Owner/Team"}],',
    '  "speaker_notes": "A detailed, professional presentation script. Explain the DAU-to-feature conversion, day-to-day variation, the trend comparison, any data quality warnings, and the reasoning behind each recommendation."',
    '}',
    '',
    'Rules (the slide has limited space, keep on-slide text short):',
    '- Ground every statement in DATA and quote numbers exactly as given.',
    '- Never describe summed feature users as unique users for the period.',
    '- If dataPeriod.days < 7, state that trends are preliminary.',
    '- If dataQualityIssues is not empty, mention the most important one in an observation.',
    '- headline: one sentence, max 20 words.',
    '- observations: exactly 3 items, each max 24 words, no leading bullet characters.',
    '- next_steps: exactly 3 items; "action" max 16 words; "pic" is a short team name (max 4 words).',
    '- speaker_notes: no length limit.',
    '',
    'DATA:',
    JSON.stringify(data)
  ].join('\n');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(LS_CONFIG.GEMINI_MODEL) + ':generateContent';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API error: ' + response.getContentText());
  }

  const envelope = JSON.parse(response.getContentText());
  const parsed = JSON.parse(envelope.candidates[0].content.parts[0].text);
  return normalizeAi_(parsed);
}

// Trims AI output to what fits on the slide; the full narrative lives in speaker notes.
function normalizeAi_(parsed) {
  const clean = s => String(s || '').replace(/^[\s•\-*]+/, '').trim();
  return {
    headline: clampWords_(clean(parsed.headline) || 'Feature Performance Executive Summary', 24),
    observations: (parsed.observations || []).map(clean).filter(Boolean).slice(0, 3).map(o => clampWords_(o, 28)),
    nextSteps: (parsed.next_steps || []).slice(0, 3).map(s => ({
      action: clampWords_(clean(typeof s === 'string' ? s : s.action), 20),
      pic: clampWords_(clean(typeof s === 'string' ? '' : s.pic) || 'TBD', 5)
    })).filter(s => s.action),
    speakerNotes: parsed.speaker_notes || ''
  };
}

function clampWords_(text, maxWords) {
  const words = String(text || '').split(/\s+/).filter(Boolean);
  return words.length <= maxWords ? words.join(' ') : words.slice(0, maxWords).join(' ').replace(/[,.;:]$/, '') + '…';
}

/* ------------------------------ SLIDE DESIGN & UI ---------------------------- */

function drawAnalysisSlide_(slide, data, ai) {
  const L = LS_CONFIG.LAYOUT;
  drawTitle_(slide, data, L.title);
  drawKpiCards_(slide, data, L.kpi);
  drawFunnel_(slide, data, L.funnel);
  drawDailyChart_(slide, data, L.chart);
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

function drawTitle_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  addRichText_(slide, [
    { text: data.feature, color: C.greenDark },
    { text: '   Feature Analysis Snapshot', color: C.muted, bold: false }
  ], box.x, box.y, 380, box.h, { fontSize: 12, bold: true, valign: 'MIDDLE' });

  const pillText = 'PERIOD  ' + formatPeriod_(data.dataPeriod.start, data.dataPeriod.end) + '  ·  ' + data.dataPeriod.days + ' days';
  const pillW = estimateTextWidth_(pillText, 6.5, true) + 18;
  const pillX = box.x + box.w - pillW;
  pill_(slide, pillText, pillX, box.y + 2, pillW, 14, C.greenSoft, C.greenDark, 6.5);

  const issues = data.dataQualityIssues.length;
  if (issues) {
    const warn = '⚠ ' + issues + ' data warning' + (issues > 1 ? 's' : '') + ' — see Run Log';
    const warnW = estimateTextWidth_(warn, 6, true) + 4;
    addText_(slide, warn, pillX - warnW - 8, box.y + 2, warnW, 14, { fontSize: 6, bold: true, color: C.red, align: 'END', valign: 'MIDDLE' });
  }
}

function drawKpiCards_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const s = data.summary;
  const latest = s.latestDay, prev = s.previousDay;

  let latestSub = fmtPct_(latest.penetrationPct, 2) + ' penetration';
  let latestColor = C.muted, latestAccent = C.greenDark;
  if (prev) {
    const diff = latest.featureUsers - prev.featureUsers;
    latestSub = (diff > 0 ? '▲ +' : diff < 0 ? '▼ −' : '■ ') + fmtNum_(Math.abs(diff)) + ' vs ' + formatShortDate_(prev.date) + ' · ' + fmtPct_(latest.penetrationPct, 2);
    latestColor = diff > 0 ? C.okInk : diff < 0 ? C.red : C.muted;
    latestAccent = diff > 0 ? C.greenLine : diff < 0 ? C.red : C.muted;
  }

  const cards = [
    { label: 'AVG DAILY ACTIVE (APP)', value: fmtNum_(s.avgDau), sub: 'users/day · ' + data.dataPeriod.days + ' days', accent: C.blue },
    {
      label: 'AVG FEATURE USERS',
      value: fmtNum_(s.avgFeatureUsers, 1),
      sub: s.peakFeatureUsersDay ? 'per day · peak ' + fmtNum_(s.peakFeatureUsersDay.featureUsers) + ' on ' + formatShortDate_(s.peakFeatureUsersDay.date) : 'per day',
      accent: C.greenBar
    },
    {
      label: 'PENETRATION',
      value: fmtPct_(s.penetrationPct, 2),
      sub: s.worstDay ? 'range ' + fmtPct_(s.worstDay.penetrationPct, 2) + ' – ' + fmtPct_(s.bestDay.penetrationPct, 2) : 'feature users / DAU',
      accent: C.greenDark
    },
    {
      label: 'LATEST · ' + formatShortDate_(latest.date).toUpperCase(),
      value: fmtNum_(latest.featureUsers) + ' users',
      sub: latestSub,
      subColor: latestColor,
      accent: latestAccent
    }
  ];
  const gap = 8;
  const w = (box.w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => drawKpiCard_(slide, box.x + i * (w + gap), box.y, w, box.h, card));
}

function drawKpiCard_(slide, x, y, w, h, card) {
  const C = LS_CONFIG.COLORS;
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h, C.cardBg, C.border);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + 7, 3, h - 14, card.accent);

  addText_(slide, card.label, x + 10, y + 5, w - 14, 8, { fontSize: 6, bold: true, color: C.muted });
  addText_(slide, card.value, x + 10, y + 14, w - 14, 16, { fontSize: 15, bold: true, color: C.ink, valign: 'MIDDLE' });
  addText_(slide, card.sub, x + 10, y + 31, w - 14, 7, { fontSize: 5.5, color: card.subColor || C.muted, bold: !!card.subColor && card.subColor !== C.muted });
}

/**
 * Two-step funnel: app DAU → users who clicked the feature (average per day).
 * Left: step rows on a shared track so the drop is visible to scale. Right: conversion callout + trend.
 */
function drawFunnel_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const f = data.funnel, s = data.summary;
  const steps = f.steps;

  panel_(slide, box);
  sectionTitle_(slide, 'FEATURE FUNNEL', box.x + 10, box.y + 8, 120);
  addText_(slide, 'Average per day · ' + data.dataPeriod.days + ' days', box.x + 140, box.y + 8, 150, 11, { fontSize: 6, color: C.muted, valign: 'MIDDLE' });

  const callout = { x: box.x + box.w - 120, y: box.y + 26, w: 110, h: box.h - 34 };
  const col = { badge: box.x + 10, label: box.x + 28, track: box.x + 112 };
  col.labelW = col.track - col.label - 6;
  col.trackW = callout.x - 12 - col.track;

  const rowH = 22;
  const rowY = [box.y + 30, box.y + 66];
  const colors = [C.greenDark, C.greenBar];

  const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, col.badge + 6, rowY[0] + rowH / 2, col.badge + 6, rowY[1] + rowH / 2);
  connector.getLineFill().setSolidFill(C.border);
  connector.setWeight(1);

  steps.forEach((step, i) => {
    const y = rowY[i], cy = y + rowH / 2;
    rect_(slide, SlidesApp.ShapeType.ELLIPSE, col.badge, cy - 6, 12, 12, colors[i], C.white, 1);
    addText_(slide, String(i + 1), col.badge, cy - 6, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });

    const labelSize = fitFontSize_([step.name], col.labelW, 9, 7, 6, 0, true);
    addText_(slide, step.name, col.label, y, col.labelW, rowH, { fontSize: labelSize, bold: true, color: C.ink, valign: 'MIDDLE' });

    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 7, col.trackW, 14, C.track);
    const ratio = i === 0 ? 1 : (f.conversionPct || 0) / 100;
    const fillW = step.avgPerDay > 0 ? Math.max(5, col.trackW * ratio) : 0;
    if (fillW) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 7, fillW, 14, colors[i]);

    if (fillW >= 34) {
      addText_(slide, fmtNum_(step.avgPerDay, 1), col.track, y, fillW - 6, rowH, { fontSize: 7.5, bold: true, color: C.white, align: 'END', valign: 'MIDDLE' });
    } else {
      const ox = col.track + fillW + 5;
      addRichText_(slide, [
        { text: fmtNum_(step.avgPerDay, 1), bold: true, color: C.ink },
        { text: '  ·  ' + fmtPct_(f.conversionPct, 2) + ' of DAU', color: C.muted }
      ], ox, y, col.track + col.trackW - ox, rowH, { fontSize: 7, valign: 'MIDDLE' });
    }
  });

  // Drop-off between the two steps
  if (f.dropOffPct != null) {
    addRichText_(slide, [
      { text: '▼ ' + fmtPct_(f.dropOffPct, 2), bold: true, color: C.red },
      { text: ' of DAU did not open ' + data.feature, color: C.muted }
    ], col.track, rowY[0] + rowH + 2, col.trackW, 10, { fontSize: 6, valign: 'MIDDLE' });
  }

  // Best / lowest day
  if (s.bestDay && s.worstDay) {
    const statsY = box.y + box.h - 22;
    hLine_(slide, box.x + 10, callout.x - 12, statsY - 3, C.border, 0.5);
    addRichText_(slide, [
      { text: 'Best day  ', color: C.muted },
      { text: formatShortDate_(s.bestDay.date) + ' · ' + fmtPct_(s.bestDay.penetrationPct, 2), bold: true, color: C.okInk },
      { text: '  (' + fmtNum_(s.bestDay.featureUsers) + ' of ' + fmtNum_(s.bestDay.dau) + ')', color: C.muted },
      { text: '      Lowest  ', color: C.muted },
      { text: formatShortDate_(s.worstDay.date) + ' · ' + fmtPct_(s.worstDay.penetrationPct, 2), bold: true, color: C.red },
      { text: '  (' + fmtNum_(s.worstDay.featureUsers) + ' of ' + fmtNum_(s.worstDay.dau) + ')', color: C.muted }
    ], box.x + 10, statsY, callout.x - 22 - box.x, 12, { fontSize: 6, valign: 'MIDDLE' });
  }

  // Conversion callout
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, callout.x, callout.y, callout.w, callout.h, C.greenSoft);
  addText_(slide, 'CLICK-THROUGH RATE', callout.x + 6, callout.y + 7, callout.w - 12, 8, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
  addText_(slide, fmtPct_(f.conversionPct, 2), callout.x + 6, callout.y + 17, callout.w - 12, 22, { fontSize: 19, bold: true, color: C.greenDark, align: 'CENTER', valign: 'MIDDLE' });
  addText_(slide, 'of DAU open ' + data.feature, callout.x + 6, callout.y + 41, callout.w - 12, 8, { fontSize: 5.5, color: C.muted, align: 'CENTER' });

  const t = s.trend;
  if (t && t.penetrationChangePp != null) {
    const pp = t.penetrationChangePp;
    const up = pp > 0, flat = pp === 0;
    const text = (flat ? '■ ' : up ? '▲ +' : '▼ −') + fmtNum_(Math.abs(pp), 2) + ' pp';
    pill_(slide, text, callout.x + 14, callout.y + 56, callout.w - 28, 12,
      flat ? C.cardBg : up ? C.okSoft : C.redSoft, flat ? C.muted : up ? C.okInk : C.red, 6.5);
    addText_(slide, t.label, callout.x + 6, callout.y + 71, callout.w - 12, 8, { fontSize: 5, color: C.muted, align: 'CENTER' });
  }
}

/** Feature users as bars (right axis) with app DAU as a line (left axis); penetration under each date when space allows. */
function drawDailyChart_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const rows = data.dailySeries;
  const n = rows.length;
  const detailed = n <= 14;

  panel_(slide, box);
  sectionTitle_(slide, 'DAILY TREND', box.x + 10, box.y + 8, 80);

  // Legend, right-aligned on the title row
  const legend = [
    { label: 'App DAU (left)', color: C.blue, kind: 'line' },
    { label: 'Feature users (right)', color: C.greenBar, kind: 'bar' }
  ];
  if (detailed) legend.push({ label: '% = penetration', color: C.greenDark, kind: 'text' });
  let lx = box.x + box.w - 10;
  legend.slice().reverse().forEach(item => {
    const tw = estimateTextWidth_(item.label, 6);
    lx -= tw;
    addText_(slide, item.label, lx, box.y + 8, tw, 11, { fontSize: 6, color: item.kind === 'text' ? C.greenDark : C.ink, valign: 'MIDDLE' });
    if (item.kind === 'line') {
      lx -= 13;
      hLine_(slide, lx, lx + 10, box.y + 13.5, item.color, 2);
    } else if (item.kind === 'bar') {
      lx -= 10;
      rect_(slide, SlidesApp.ShapeType.RECTANGLE, lx, box.y + 10, 7, 7, item.color);
    }
    lx -= 12;
  });

  const plot = { x: box.x + 34, y: box.y + 30, w: box.w - 34 - 34, h: box.h - 30 - (detailed ? 24 : 16) };
  const maxDau = niceMax_(Math.max(1, ...rows.map(r => r.dau)));
  // Bars use a taller scale so they stay in the lower part of the plot, clear of the DAU line
  const maxUsers = niceMax_(Math.max(1, ...rows.map(r => r.featureUsers)) * 1.6);
  const band = plot.w / n;
  const px = i => plot.x + (i + 0.5) * band;
  const yOf = (v, max) => plot.y + plot.h - (v / max) * plot.h;

  // Gridlines + dual axis labels
  [0, 0.5, 1].forEach(fr => {
    const gy = plot.y + plot.h - fr * plot.h;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, fr === 0 ? C.border : C.grid, fr === 0 ? 0.75 : 0.5);
    if (fr > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(maxDau * fr, maxDau < 10 ? 1 : 0), box.x + 6, gy - 4, plot.x - box.x - 10, 8, { fontSize: 5.5, color: C.blue, align: 'END', valign: 'MIDDLE' });
    addText_(slide, fmtNum_(maxUsers * fr, maxUsers < 10 ? 1 : 0), plot.x + plot.w + 4, gy - 4, 24, 8, { fontSize: 5.5, color: C.greenBar, valign: 'MIDDLE' });
  });

  // Feature-user bars
  const barW = Math.max(2, Math.min(16, band * 0.55));
  rows.forEach((r, i) => {
    const top = yOf(r.featureUsers, maxUsers);
    const h = plot.y + plot.h - top;
    if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, px(i) - barW / 2, top, barW, h, C.greenBar);
    if (detailed) {
      addText_(slide, fmtNum_(r.featureUsers), px(i) - 15, top - 9, 30, 8, { fontSize: 5.5, bold: true, color: C.greenDark, align: 'CENTER', valign: 'MIDDLE' });
    }
  });

  // DAU line + markers
  for (let i = 1; i < n; i++) {
    const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, px(i - 1), yOf(rows[i - 1].dau, maxDau), px(i), yOf(rows[i].dau, maxDau));
    line.getLineFill().setSolidFill(C.blue);
    line.setWeight(1.75);
  }
  if (detailed) {
    rows.forEach((r, i) => {
      const y = yOf(r.dau, maxDau);
      rect_(slide, SlidesApp.ShapeType.ELLIPSE, px(i) - 2.25, y - 2.25, 4.5, 4.5, C.blue, C.white, 0.75);
      addText_(slide, fmtNum_(r.dau), px(i) - 15, y - 11, 30, 8, { fontSize: 5.5, bold: true, color: C.blue, align: 'CENTER', valign: 'MIDDLE' });
    });
  }

  // X-axis: date (+ penetration when detailed), thinned out for long ranges
  const every = Math.ceil(n / 8);
  const labelIdx = [];
  for (let i = 0; i < n; i += every) labelIdx.push(i);
  if (labelIdx[labelIdx.length - 1] !== n - 1) {
    // Always label the latest date; drop the previous label if the two would collide
    if (n - 1 - labelIdx[labelIdx.length - 1] < every * 0.6) labelIdx.pop();
    labelIdx.push(n - 1);
  }
  labelIdx.forEach(i => {
    const r = rows[i];
    const parts = [{ text: formatShortDate_(r.date), color: C.muted }];
    if (detailed) parts.push({ text: '\n' + fmtPct_(r.penetrationPct, 2), color: C.greenDark, bold: true });
    addRichText_(slide, parts, px(i) - 22, plot.y + plot.h + 3, 44, detailed ? 16 : 8, { fontSize: 5.5, align: 'CENTER' });
  });
}

function drawObservations_(slide, ai, box) {
  const C = LS_CONFIG.COLORS;
  headerPanel_(slide, box, 'EXECUTIVE OBSERVATIONS');

  const innerX = box.x + 10, innerW = box.w - 20;
  let y = box.y + 27;

  // Headline callout
  const hlSize = 7.5;
  const hlLines = Math.min(4, estimateLines_(ai.headline, innerW - 12, hlSize, true));
  const hlH = hlLines * hlSize * LS_CONFIG.LINE_HEIGHT + 8;
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
  const C = LS_CONFIG.COLORS;
  headerPanel_(slide, box, 'ACTION PLAN / PIC');

  const steps = ai.nextSteps;
  if (!steps.length) return;

  const badgeX = box.x + 10;
  const textX = box.x + 28;
  const textW = box.x + box.w - 10 - textX;
  const top = box.y + 28;
  const avail = box.y + box.h - 6 - top;
  const gap = 6;

  const itemHeight = (step, size) => estimateLines_(step.action, textW, size) * size * LS_CONFIG.LINE_HEIGHT + 2 + (size + 4);
  const totalHeight = size => steps.reduce((acc, step) => acc + itemHeight(step, size) + gap, -gap);
  let size = 7.5;
  while (size > 5.5 && totalHeight(size) > avail) size -= 0.5;

  let y = top;
  steps.forEach((step, i) => {
    if (i > 0) hLine_(slide, textX, box.x + box.w - 10, y - gap / 2, C.border, 0.5);

    rect_(slide, SlidesApp.ShapeType.ELLIPSE, badgeX, y, 12, 12, C.greenDark);
    addText_(slide, String(i + 1), badgeX, y, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });

    const actionH = estimateLines_(step.action, textW, size) * size * LS_CONFIG.LINE_HEIGHT;
    addText_(slide, step.action, textX, y, textW, actionH, { fontSize: size, color: C.ink });

    const picText = 'PIC · ' + step.pic;
    const picSize = Math.max(5.5, size - 1);
    const picW = Math.min(textW, estimateTextWidth_(picText, picSize, true) + 12);
    pill_(slide, picText, textX, y + actionH + 2, picW, size + 4, C.greenSoft, C.greenDark, picSize);

    y += itemHeight(step, size) + gap;
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
  const C = LS_CONFIG.COLORS;
  return rect_(slide, SlidesApp.ShapeType.RECTANGLE, box.x, box.y, box.w, box.h, C.panelBg, C.border, 0.75);
}

function sectionTitle_(slide, title, x, y, w) {
  const C = LS_CONFIG.COLORS;
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + 1.5, 3, 8, C.yellow);
  return addText_(slide, title, x + 7, y, w, 11, { fontSize: 8, bold: true, color: C.greenDark, valign: 'MIDDLE' });
}

function headerPanel_(slide, box, title) {
  const C = LS_CONFIG.COLORS;
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
  const ix = LS_CONFIG.TEXT_INSET_X, iy = LS_CONFIG.TEXT_INSET_Y;
  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x - ix, y - iy, w + ix * 2, h + iy * 2);
  shape.getFill().setTransparent();
  shape.getBorder().setTransparent();
  shape.setContentAlignment(opts.valign === 'MIDDLE' ? SlidesApp.ContentAlignment.MIDDLE : SlidesApp.ContentAlignment.TOP);

  const str = String(text == null ? '' : text);
  if (!str) return shape;
  const range = shape.getText();
  range.setText(str);
  range.getTextStyle()
    .setFontFamily(LS_CONFIG.FONT)
    .setFontSize(opts.fontSize || 8)
    .setForegroundColor(opts.color || LS_CONFIG.COLORS.ink)
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

/** Text box made of styled runs: parts = [{ text, color?, bold? }]. */
function addRichText_(slide, parts, x, y, w, h, opts) {
  const shape = addText_(slide, parts.map(p => p.text).join(''), x, y, w, h, opts);
  let pos = 0;
  parts.forEach(p => {
    const end = pos + p.text.length;
    if (p.text.length && (p.bold != null || p.color)) {
      const style = shape.getText().getRange(pos, end).getTextStyle();
      if (p.bold != null) style.setBold(p.bold);
      if (p.color) style.setForegroundColor(p.color);
    }
    pos = end;
  });
  return shape;
}

/* ------------------------------ TEXT FITTING & FORMATTING ------------------------------ */

function estimateTextWidth_(text, size, bold) {
  return String(text || '').length * size * LS_CONFIG.CHAR_WIDTH * (bold ? 1.1 : 1) + 2;
}

function estimateLines_(text, w, size, bold) {
  const maxChars = Math.max(1, Math.floor(w / (size * LS_CONFIG.CHAR_WIDTH * (bold ? 1.1 : 1))));
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
    const height = paragraphs.reduce((acc, p) => acc + estimateLines_(p, w, size, bold) * size * LS_CONFIG.LINE_HEIGHT, 0) +
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

/** Thousands separators; up to `digits` decimals with trailing zeros dropped. */
function fmtNum_(n, digits) {
  if (n == null || !isFinite(n)) return '–';
  const parts = String(Number(Number(n).toFixed(digits || 0))).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function fmtPct_(value, digits) {
  return value == null ? 'n/a' : value.toFixed(digits == null ? 1 : digits) + '%';
}

function formatShortDate_(key) {
  return Utilities.formatDate(parseDate_(key), Session.getScriptTimeZone(), 'd MMM');
}

function formatPeriod_(startKey, endKey) {
  const tz = Session.getScriptTimeZone();
  const start = parseDate_(startKey), end = parseDate_(endKey);
  if (startKey === endKey) return Utilities.formatDate(end, tz, 'd MMM yyyy');
  const startPattern = start.getFullYear() === end.getFullYear() ? 'd MMM' : 'd MMM yyyy';
  return Utilities.formatDate(start, tz, startPattern) + ' – ' + Utilities.formatDate(end, tz, 'd MMM yyyy');
}

/* ------------------------------ PRESENTATION & LOGGING --------------------------- */

function getOrCreatePresentation_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(LS_CONFIG.PRESENTATION_PROPERTY);
  if (existingId) {
    try {
      return SlidesApp.openById(existingId);
    } catch (e) {
      throw new Error('Cannot open presentation. Verify permissions.');
    }
  }
  throw new Error('PRESENTATION_ID property not set in Script Properties.');
}

function logRun_(ss, status, message, url) {
  const sheet = getOrCreateSheet_(ss, LS_CONFIG.SHEET_LOG);
  sheet.appendRow([new Date(), status, message, url || '']);
}
