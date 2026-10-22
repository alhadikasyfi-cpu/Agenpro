/**
 * Lapak Sawit Feature Analysis — Google Apps Script
 * Duplicates the SawitPRO master template (slide 1) and renders a feature analysis snapshot:
 * KPI cards, DAU → feature funnel per app (AgenPRO / TimPRO), daily CTR chart and Gemini insights.
 *
 * Source: the sheet named after the feature (e.g. "Lapak Sawit"), one row per date per app
 * (feature_name = "Lapak Sawit Agenpro", "Lapak Sawit Timpro", ...). Rows are read from row 2 down to
 * the first empty row; columns are matched by header name, so their order does not matter.
 */

const LS_CONFIG = {
  GEMINI_MODEL: 'gemini-3.5-flash-lite',
  PRESENTATION_PROPERTY: 'LAPAK_SAWIT_PRESENTATION_ID',
  API_KEY_PROPERTY: 'GEMINI_API_KEY',
  FEATURE_NAME: 'Lapak Sawit', // also the name of the source sheet
  SHEET_LOG: 'Run Log',
  OUTPUT_LANGUAGE: 'English',
  // Slides generated per run, in order: 'summary' = all apps side by side, 'apps' = one detail slide per app.
  // Use ['apps'] for per-app slides only.
  SLIDES: ['summary', 'apps'],
  FONT: 'Arial',

  // Accepted header names per field, after normalising (lower-case, non-alphanumerics → "_").
  DAILY_COLUMNS: {
    date: ['tanggal', 'date'],
    feature: ['feature_name', 'feature', 'app'],
    dau: ['daily_active_users', 'dau'],
    featureUsers: ['feature_unique_users', 'feature_users', 'unique_users'],
    newUsers: ['new_users', 'new_user'],
    returningUsers: ['returning_users', 'returning_user', 'return_users'],
    penetration: ['penetration_rate_pct', 'penetration_pct'],
    userNames: ['user_names', 'usernames', 'names']
  },
  // Fallback for per-feature user columns such as "feature_lapaksawit_users"
  FEATURE_USERS_PATTERN: /^feature_.+_users$/,
  REQUIRED_COLUMNS: ['date', 'dau', 'featureUsers'],

  // App variants, in display order. `key` is matched against feature_name minus the feature name.
  APPS: [
    {
      key: 'agenpro',
      label: 'AgenPRO',
      color: '#31512B',
      soft: '#DCE8D5',
      description: 'App for agents: view TBS data, register DO, join the Panen Emas loyalty program, manage their members ' +
        '(pengurus, drivers/supir, other members and plantation/kebun data) and register pengurus into the Protim program.'
    },
    {
      key: 'timpro',
      label: 'TimPRO',
      color: '#C2410C',
      soft: '#FDE3CF',
      description: 'Loyalty & operational app for pengurus: join the Protim program and affiliate program, view TBS delivery ' +
        'reports and look for palm fruit to buy/sell (Lapak Sawit).'
    }
  ],
  FALLBACK_APP_COLORS: [{ color: '#2563EB', soft: '#DBEAFE' }, { color: '#7C3AED', soft: '#EDE9FE' }],

  // Product context sent to Gemini
  FEATURE_CONTEXT: [
    'Panen Emas: loyalty program for agents; agents claim tonnage in the app every time they deliver fruit (FFB/TBS) to a partner mill.',
    'Protim: appreciation loyalty program for pengurus; pengurus claim tonnage in the TimPRO app for each fruit delivery, ' +
      'but only after the agent has first claimed it in Panen Emas.',
    'Laporan TBS: TBS (fresh fruit bunch) delivery report, available after every delivery.',
    'Lapak Sawit: marketplace connecting agents and farmers (petani) to run FFB transactions. Available in both AgenPRO and TimPRO.'
  ],

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
    returning: '#E0A800',
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
    funnel: { x: 20, y: 126, w: 420, h: 130 },
    chart: { x: 20, y: 262, w: 420, h: 102 },
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
  const data = getOrCreateSheet_(ss, LS_CONFIG.FEATURE_NAME);
  const log = getOrCreateSheet_(ss, LS_CONFIG.SHEET_LOG);

  setupTab_(data, ['tanggal', 'feature_name', 'daily_active_users', 'feature_lapaksawit_users', 'New_users', 'Returning user',
    'Return user pct', 'penetration_rate_pct', 'user_names'], [
    ['15/10/2026', 'Lapak Sawit Agenpro', 205, 2, 2, 0, 0, 0.98, 'user a, user b'],
    ['15/10/2026', 'Lapak Sawit Timpro', 134, 1, 1, 0, 0, 0.75, 'user c']
  ]);
  setupTab_(log, ['run_at', 'status', 'message', 'presentation_url'], []);
  [data, log].forEach(sh => {
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
    const sheet = readFeatureSheet_(ss);
    const analysisInput = buildAnalysisInput_(sheet);

    // One spec per slide: the summary and/or one detail slide per app
    const specs = [];
    LS_CONFIG.SLIDES.forEach(kind => {
      if (kind === 'summary') specs.push({ app: null });
      if (kind === 'apps') analysisInput.apps.forEach(app => specs.push({ app: app }));
    });
    if (!specs.length) throw new Error('LS_CONFIG.SLIDES is empty.');

    // All Gemini calls run in parallel
    const aiResults = callGeminiAll_(analysisInput, specs.map(s => s.app));

    const presentation = getOrCreatePresentation_();
    const slides = presentation.getSlides();
    if (!slides.length) throw new Error('Presentation missing master template on Slide 1.');

    specs.forEach((spec, i) => {
      // Duplicate the master template and keep this run's slides in order right after it
      const newSlide = slides[0].duplicate();
      newSlide.move(1 + i);

      if (spec.app) drawAppSlide_(newSlide, analysisInput, spec.app, aiResults[i]);
      else drawSummarySlide_(newSlide, analysisInput, aiResults[i]);

      if (aiResults[i].speakerNotes) {
        newSlide.getNotesPage().getSpeakerNotesShape().getText().setText(aiResults[i].speakerNotes);
      }
    });

    presentationUrl = 'https://docs.google.com/presentation/d/' + presentation.getId() + '/edit';
    const issues = analysisInput.dataQualityIssues;
    const message = specs.length + ' slide(s) generated (' + analysisInput.dataPeriod.days + ' days, ' + analysisInput.apps.length + ' app(s)).' +
      (issues.length ? ' Data warnings: ' + issues.join('; ') : '');
    logRun_(ss, issues.length ? 'SUCCESS_WITH_WARNINGS' : 'SUCCESS', message, presentationUrl);

    SpreadsheetApp.getUi().alert(specs.length + ' slide(s) generated successfully!' +
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

function normalizeHeader_(h) {
  return String(h == null ? '' : h).trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

/** Reads the feature sheet from row 2 down to the first fully empty row; columns are located by header name. */
function readFeatureSheet_(ss) {
  const sheetName = LS_CONFIG.FEATURE_NAME;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet "' + sheetName + '".');
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('Sheet "' + sheetName + '" has no data rows.');

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(normalizeHeader_);
  const col = {};
  Object.keys(LS_CONFIG.DAILY_COLUMNS).forEach(field => {
    const names = LS_CONFIG.DAILY_COLUMNS[field];
    col[field] = headers.findIndex(h => names.indexOf(h) >= 0);
  });
  if (col.featureUsers < 0) col.featureUsers = headers.findIndex(h => LS_CONFIG.FEATURE_USERS_PATTERN.test(h));
  LS_CONFIG.REQUIRED_COLUMNS.forEach(field => {
    if (col[field] < 0) {
      throw new Error('Sheet "' + sheetName + '" is missing column "' + LS_CONFIG.DAILY_COLUMNS[field].join('" / "') + '".');
    }
  });

  const optionalNumber = (row, field) => (col[field] >= 0 && row[col[field]] !== '' ? toNumber_(row[col[field]]) : null);
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
      newUsers: optionalNumber(row, 'newUsers'),
      returningUsers: optionalNumber(row, 'returningUsers'),
      penetration: optionalNumber(row, 'penetration'),
      names: col.userNames >= 0 ? parseNames_(row[col.userNames]) : null
    });
  }
  if (!rows.length) throw new Error('Sheet "' + sheetName + '" has no data before the first empty row.');
  return { rows: rows, hasNames: col.userNames >= 0 };
}

/** Accepts real dates, dd/MM/yyyy, yyyyMMdd (number or text) and yyyy-MM-dd. */
function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (value === '' || value == null) return null;
  const s = String(value).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/) || s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function toNumber_(value) {
  const n = Number(String(value == null ? '' : value).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}

function parseNames_(value) {
  return String(value == null ? '' : value).split(',').map(s => s.trim()).filter(Boolean);
}

function nameKey_(name) {
  return String(name).toLowerCase().replace(/\s+/g, ' ').trim();
}

/** "Lapak Sawit Agenpro" → AgenPRO config entry (or a generated one for unknown apps). */
function resolveApp_(featureName, fallbackIndex) {
  const rest = String(featureName || '').toLowerCase().replace(LS_CONFIG.FEATURE_NAME.toLowerCase(), '').replace(/[^a-z0-9]+/g, '');
  const known = LS_CONFIG.APPS.find(a => a.key === rest);
  if (known) return known;
  const fallback = LS_CONFIG.FALLBACK_APP_COLORS[fallbackIndex % LS_CONFIG.FALLBACK_APP_COLORS.length];
  const label = String(featureName || '').replace(new RegExp(LS_CONFIG.FEATURE_NAME, 'i'), '').trim() || LS_CONFIG.FEATURE_NAME;
  return { key: rest || 'all', label: label, color: fallback.color, soft: fallback.soft, description: '' };
}

function buildAnalysisInput_(sheet) {
  const issues = [];

  // Group rows per app (feature_name)
  const groups = {};
  sheet.rows.forEach(r => {
    const name = r.feature || LS_CONFIG.FEATURE_NAME;
    (groups[name] = groups[name] || []).push(r);
  });
  const appOrder = key => {
    const i = LS_CONFIG.APPS.findIndex(a => a.key === key);
    return i < 0 ? LS_CONFIG.APPS.length : i;
  };
  const apps = Object.keys(groups)
    .map((name, i) => analyzeApp_(name, resolveApp_(name, i), groups[name], sheet.hasNames, issues))
    .filter(Boolean)
    .sort((a, b) => appOrder(a.key) - appOrder(b.key) || a.app.localeCompare(b.app));
  if (!apps.length) throw new Error('No valid rows in sheet "' + LS_CONFIG.FEATURE_NAME + '".');

  const allDates = {};
  apps.forEach(a => a.dailySeries.forEach(r => allDates[r.date] = true));
  const dates = Object.keys(allDates).sort();

  // Combined view across apps
  const totalDau = apps.reduce((acc, a) => acc + a.summary.totalDauUserDays, 0);
  const totalFeature = apps.reduce((acc, a) => acc + a.summary.totalFeatureUserDays, 0);
  let combinedUnique = null, combinedRepeat = null;
  if (sheet.hasNames) {
    const days = {};
    apps.forEach(a => a._nameDays && Object.keys(a._nameDays).forEach(k => {
      days[k] = days[k] || {};
      Object.keys(a._nameDays[k]).forEach(d => days[k][d] = true);
    }));
    combinedUnique = Object.keys(days).length;
    combinedRepeat = Object.keys(days).filter(k => Object.keys(days[k]).length >= 2).length;
  }
  apps.forEach(a => delete a._nameDays); // user names stay inside the script; never sent to Gemini

  return {
    feature: LS_CONFIG.FEATURE_NAME,
    funnelDefinition: 'Per app: Step 1 = app Daily Active Users (DAU); Step 2 = users who clicked/opened ' + LS_CONFIG.FEATURE_NAME +
      '. CTR (penetration) = feature users / DAU. Feature users split into new (first time) and returning users.',
    unitsNote: 'Summed daily feature users are user-days. uniqueUsers counts distinct user names over the whole period; ' +
      'repeatUsers = users active on 2 or more distinct days.',
    dataPeriod: { start: dates[0], end: dates[dates.length - 1], days: dates.length },
    combined: {
      avgDau: apps.reduce((acc, a) => acc + a.summary.avgDau, 0),
      avgFeatureUsers: round1_(apps.reduce((acc, a) => acc + a.summary.avgFeatureUsers, 0)),
      ctrPct: totalDau > 0 ? round2_((totalFeature / totalDau) * 100) : null,
      totalDauUserDays: totalDau,
      totalFeatureUserDays: totalFeature,
      uniqueUsers: combinedUnique,
      repeatUsers: combinedRepeat
    },
    apps: apps,
    dataQualityIssues: issues
  };
}

function analyzeApp_(featureName, appDef, rows, hasNames, issues) {
  const tag = appDef.label + ' ';
  const byDate = {};
  rows.forEach(r => {
    if (!r.date) {
      issues.push(tag + 'row ' + r.rowNumber + ': unreadable date (row skipped)');
      return;
    }
    const key = dateKey_(r.date);
    if (byDate[key]) {
      issues.push(tag + 'duplicate date ' + key + ' (row ' + r.rowNumber + ' skipped)');
      return;
    }
    const ctr = r.dau > 0 ? round2_((r.featureUsers / r.dau) * 100) : null;
    if (r.featureUsers > r.dau) issues.push(tag + key + ': feature users exceed DAU');
    if (r.penetration != null && ctr != null && Math.abs(r.penetration - ctr) > 0.05) {
      issues.push(tag + key + ': penetration_rate_pct ' + r.penetration + ' differs from computed ' + ctr);
    }
    if (r.newUsers != null && r.returningUsers != null && r.newUsers + r.returningUsers !== r.featureUsers) {
      issues.push(tag + key + ': new + returning (' + (r.newUsers + r.returningUsers) + ') ≠ feature users (' + r.featureUsers + ')');
    }
    if (r.names && r.names.length && r.names.length !== r.featureUsers) {
      issues.push(tag + key + ': ' + r.names.length + ' user names vs ' + r.featureUsers + ' feature users');
    }
    byDate[key] = { date: key, dau: r.dau, featureUsers: r.featureUsers, newUsers: r.newUsers, returningUsers: r.returningUsers, ctrPct: ctr, names: r.names || [] };
  });

  const full = Object.keys(byDate).sort().map(k => byDate[k]);
  if (!full.length) return null;

  const days = full.length;
  const spanDays = Math.round((parseDate_(full[days - 1].date).getTime() - parseDate_(full[0].date).getTime()) / 86400000) + 1;
  if (spanDays > days) issues.push(tag + (spanDays - days) + ' missing date(s) between ' + full[0].date + ' and ' + full[days - 1].date);

  // Unique and repeat users from the name lists
  let nameDays = null, uniqueUsers = null, repeatUsers = null, maxActiveDays = null;
  if (hasNames) {
    nameDays = {};
    full.forEach(r => r.names.forEach(n => {
      const k = nameKey_(n);
      nameDays[k] = nameDays[k] || {};
      nameDays[k][r.date] = true;
    }));
    const counts = Object.keys(nameDays).map(k => Object.keys(nameDays[k]).length);
    uniqueUsers = counts.length;
    repeatUsers = counts.filter(c => c >= 2).length;
    maxActiveDays = counts.length ? Math.max(...counts) : 0;
  }

  const series = full.map(r => ({ date: r.date, dau: r.dau, featureUsers: r.featureUsers, newUsers: r.newUsers, returningUsers: r.returningUsers, ctrPct: r.ctrPct }));
  const sum = key => series.reduce((acc, r) => acc + (r[key] || 0), 0);
  const totalDau = sum('dau'), totalFeature = sum('featureUsers');
  const hasSplit = series.every(r => r.newUsers != null && r.returningUsers != null);
  const withCtr = series.filter(r => r.ctrPct != null);
  const pick = better => withCtr.reduce((best, r) => (best == null || better(r, best) ? r : best), null);

  return {
    key: appDef.key,
    app: appDef.label,
    featureName: featureName,
    appDescription: appDef.description,
    color: appDef.color,
    soft: appDef.soft,
    summary: {
      days: days,
      avgDau: Math.round(totalDau / days),
      avgFeatureUsers: round1_(totalFeature / days),
      ctrPct: totalDau > 0 ? round2_((totalFeature / totalDau) * 100) : null,
      totalDauUserDays: totalDau,
      totalFeatureUserDays: totalFeature,
      avgNewUsers: hasSplit ? round1_(sum('newUsers') / days) : null,
      avgReturningUsers: hasSplit ? round1_(sum('returningUsers') / days) : null,
      returningSharePct: hasSplit && totalFeature > 0 ? round1_((sum('returningUsers') / totalFeature) * 100) : null,
      uniqueUsers: uniqueUsers,
      repeatUsers: repeatUsers,
      repeatRatePct: uniqueUsers ? round1_((repeatUsers / uniqueUsers) * 100) : null,
      maxActiveDaysBySingleUser: maxActiveDays,
      peakFeatureUsersDay: series.reduce((best, r) => (best == null || r.featureUsers > best.featureUsers ? r : best), null),
      notableDays: notableDays_(series),
      bestDay: pick((a, b) => a.ctrPct > b.ctrPct),
      worstDay: pick((a, b) => a.ctrPct < b.ctrPct),
      latestDay: series[days - 1],
      previousDay: days > 1 ? series[days - 2] : null,
      trend: compareWindows_(series)
    },
    patterns: buildPatterns_(series),
    dailySeries: series,
    _nameDays: nameDays
  };
}

/** Largest day-over-day CTR rise and drop, so Gemini can point at specific dates to investigate. */
function notableDays_(series) {
  const moves = [];
  for (let i = 1; i < series.length; i++) {
    const prev = series[i - 1], cur = series[i];
    if (prev.ctrPct == null || cur.ctrPct == null) continue;
    moves.push({
      date: cur.date,
      previousDate: prev.date,
      ctrChangePp: round2_(cur.ctrPct - prev.ctrPct),
      fromCtrPct: prev.ctrPct,
      toCtrPct: cur.ctrPct,
      featureUsers: cur.featureUsers,
      previousFeatureUsers: prev.featureUsers,
      dau: cur.dau,
      previousDau: prev.dau,
      newUsers: cur.newUsers,
      returningUsers: cur.returningUsers
    });
  }
  const rise = moves.filter(m => m.ctrChangePp > 0).sort((a, b) => b.ctrChangePp - a.ctrChangePp)[0] || null;
  const drop = moves.filter(m => m.ctrChangePp < 0).sort((a, b) => a.ctrChangePp - b.ctrChangePp)[0] || null;
  return { biggestRise: rise, biggestDrop: drop };
}

/** Recent window vs the window before it: 7 vs 7 days when there is enough data, otherwise half vs half. */
function compareWindows_(series) {
  if (series.length < 2) return null;
  const size = series.length >= 14 ? 7 : Math.floor(series.length / 2);
  const recent = series.slice(-size), prior = series.slice(-2 * size, -size);
  const agg = rows => {
    const dau = rows.reduce((a, r) => a + r.dau, 0), users = rows.reduce((a, r) => a + r.featureUsers, 0);
    return { avgDau: dau / rows.length, avgUsers: users / rows.length, ctr: dau > 0 ? (users / dau) * 100 : null };
  };
  const a = agg(recent), b = agg(prior);
  return {
    label: 'last ' + size + ' vs previous ' + size + ' days',
    windowDays: size,
    recentCtrPct: a.ctr == null ? null : round2_(a.ctr),
    priorCtrPct: b.ctr == null ? null : round2_(b.ctr),
    ctrChangePp: a.ctr == null || b.ctr == null ? null : round2_(a.ctr - b.ctr),
    featureUsersChangePct: pctChange_(a.avgUsers, b.avgUsers),
    dauChangePct: pctChange_(a.avgDau, b.avgDau)
  };
}

/** Statistical patterns that are only reported once there is enough data to be meaningful. */
function buildPatterns_(series) {
  const ctrs = series.filter(r => r.ctrPct != null);
  const out = {};
  if (ctrs.length >= 7) {
    const mean = ctrs.reduce((a, r) => a + r.ctrPct, 0) / ctrs.length;
    const sd = Math.sqrt(ctrs.reduce((a, r) => a + Math.pow(r.ctrPct - mean, 2), 0) / ctrs.length);
    out.ctrStdDevPp = round2_(sd);
    out.anomalies = sd > 0 ? ctrs.filter(r => Math.abs(r.ctrPct - mean) > 2 * sd).map(r => ({ date: r.date, ctrPct: r.ctrPct })) : [];
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
    out.ctrByWeekday = names
      .map((n, i) => ({ day: n, ctrPct: buckets[i].dau > 0 ? round2_((buckets[i].users / buckets[i].dau) * 100) : null }))
      .filter(d => d.ctrPct != null);
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

/** One Gemini request per slide (focusApp = null for the summary slide), sent in parallel. */
function callGeminiAll_(data, focusApps) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(LS_CONFIG.API_KEY_PROPERTY);
  if (!apiKey) throw new Error('Gemini API key missing in Script Properties.');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(LS_CONFIG.GEMINI_MODEL) + ':generateContent';
  const requests = focusApps.map(app => ({
    url: url,
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: buildGeminiPrompt_(data, app) }] }],
      generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  }));

  return UrlFetchApp.fetchAll(requests).map((response, i) => {
    if (response.getResponseCode() !== 200) {
      throw new Error('Gemini API error (' + (focusApps[i] ? focusApps[i].app : 'summary') + '): ' + response.getContentText());
    }
    const envelope = JSON.parse(response.getContentText());
    return normalizeAi_(JSON.parse(envelope.candidates[0].content.parts[0].text));
  });
}

function buildGeminiPrompt_(data, focusApp) {
  const appLines = LS_CONFIG.APPS.map(a => '- ' + a.label + ': ' + a.description);
  const payload = focusApp ? {
    feature: data.feature,
    funnelDefinition: data.funnelDefinition,
    unitsNote: data.unitsNote,
    dataPeriod: data.dataPeriod,
    focusApp: focusApp,
    benchmarkOtherApps: data.apps.filter(a => a !== focusApp).map(a => ({
      app: a.app, ctrPct: a.summary.ctrPct, avgFeatureUsers: a.summary.avgFeatureUsers, uniqueUsers: a.summary.uniqueUsers
    })),
    dataQualityIssues: data.dataQualityIssues.filter(i => i.indexOf(focusApp.app + ' ') === 0)
  } : data;

  const scope = focusApp ? [
    'SLIDE SCOPE: this slide covers ONLY ' + focusApp.app + ' (' + (focusApp.appDescription || 'app variant') + ').',
    'Analyze its funnel DAU → click ' + data.feature + ' → returning users. Use benchmarkOtherApps only as a brief comparison.',
    'Recommendations must suit the ' + focusApp.app + ' users (who they are is described in PRODUCT CONTEXT).'
  ] : [
    'SLIDE SCOPE: executive summary comparing all apps side by side (' + data.apps.map(a => a.app).join(' vs ') + ').',
    'Explain which app converts DAU to ' + data.feature + ' better and why that may be, given each app\'s users.'
  ];

  return [
    'You are a senior product analytics consultant at SawitPRO (palm-oil / FFB supply chain platform in Indonesia).',
    'You are analyzing the "' + data.feature + '" feature.',
    '',
    'PRODUCT CONTEXT',
    'SawitPRO has 2 app versions:',
    appLines.join('\n'),
    'Features:',
    LS_CONFIG.FEATURE_CONTEXT.map(l => '- ' + l).join('\n'),
    'Users of AgenPRO are agents; users of TimPRO are pengurus. Interpret differences between the apps in light of these roles.',
    'The other features are listed ONLY so you understand who the users are and what else they do in the app.',
    '',
    scope.join('\n'),
    '',
    'METRICS',
    'Funnel: ' + data.funnelDefinition,
    'Units: ' + data.unitsNote,
    '',
    'Output must be valid JSON matching this schema:',
    '{',
    '  "headline": "Clear single executive takeaway sentence",',
    '  "observations": ["Observation 1", "Observation 2", "Observation 3"],',
    '  "next_steps": [{"action": "Actionable item description", "pic": "Owner/Team"}],',
    '  "speaker_notes": "A detailed, professional presentation script for this slide: walk through the funnel, new vs returning and repeat users, the trend, the specific days that rose or dropped and what may explain them, any data quality warnings, and the reasoning behind each recommendation."',
    '}',
    '',
    'Rules (the slide has limited space, keep on-slide text short):',
    '- Write in ' + LS_CONFIG.OUTPUT_LANGUAGE + '.',
    '- Use plain business language. Never write JSON field names (e.g. uniqueUsers, repeatUsers, ctrPct, user-days); say "unique users", "repeat users", "CTR", "daily users".',
    '- Ground every statement in DATA and quote numbers exactly as given; always say which app a number belongs to.',
    '- Never describe summed feature users as unique users; use uniqueUsers for distinct users.',
    '- If dataPeriod.days < 7, state that trends are preliminary.',
    '- If dataQualityIssues is not empty, mention the most important one in an observation.',
    '',
    'ACTION PLAN RULES (next_steps):',
    '- Every action must be about the ' + data.feature + ' feature itself: getting more users to open it, getting them to come back, or understanding why usage moved.',
    '- Good action types: a ' + data.feature + ' campaign or promo, ads, push notification / WhatsApp blast, in-app banner or reminder for ' + data.feature + ', ' +
      'improving the ' + data.feature + ' flow or content, following up a user segment (new users who never came back, repeat users), ' +
      'or investigating a specific date that spiked or dropped.',
    '- Do NOT recommend integrating, linking or bundling ' + data.feature + ' with other features (Panen Emas, Protim, Laporan TBS, DO, affiliate, etc.), and do not recommend changes to other features.',
    '- At least one action must name a specific date from summary.notableDays (or bestDay / worstDay) and say what to check ' +
      '(e.g. a campaign or broadcast sent that day, FFB price or supply changes, listings available, app release or bug, holiday or harvest day).',
    '- Each action must follow from a number in DATA; mention that number briefly when it fits.',
    '- headline: one sentence, max 20 words.',
    '- observations: exactly 3 items, each max 24 words, no leading bullet characters.',
    '- next_steps: exactly 3 items; "action" max 16 words; "pic" is a short team name (max 4 words).',
    '- speaker_notes: no length limit.',
    '',
    'DATA:',
    JSON.stringify(payload)
  ].join('\n');
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

/** Summary slide: all apps side by side. */
function drawSummarySlide_(slide, data, ai) {
  const L = LS_CONFIG.LAYOUT;
  drawTitle_(slide, data, L.title);
  drawKpiCards_(slide, data, L.kpi);
  drawFunnel_(slide, data, L.funnel);
  drawDailyChart_(slide, data, L.chart);
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

/** Detail slide for a single app. */
function drawAppSlide_(slide, data, app, ai) {
  const L = LS_CONFIG.LAYOUT;
  drawTitle_(slide, data, L.title, app);
  drawAppKpiCards_(slide, data, app, L.kpi);
  drawAppFunnel_(slide, data, app, L.funnel);
  drawAppDailyChart_(slide, app, L.chart);
  drawObservations_(slide, ai, L.observations);
  drawNextSteps_(slide, ai, L.actions);
}

function drawTitle_(slide, data, box, app) {
  const C = LS_CONFIG.COLORS;
  const parts = [{ text: data.feature, color: C.greenDark }];
  if (app) parts.push({ text: '  ·  ', color: C.muted }, { text: app.app, color: app.color });
  parts.push({ text: app ? '   Feature Deep Dive' : '   Feature Analysis Snapshot · All Apps', color: C.muted, bold: false });
  addRichText_(slide, parts, box.x, box.y, 400, box.h, { fontSize: 12, bold: true, valign: 'MIDDLE' });

  const pillText = 'PERIOD  ' + formatPeriod_(data.dataPeriod.start, data.dataPeriod.end) + '  ·  ' + data.dataPeriod.days + ' days';
  const pillW = estimateTextWidth_(pillText, 6.5, true) + 18;
  const pillX = box.x + box.w - pillW;
  pill_(slide, pillText, pillX, box.y + 2, pillW, 14, C.greenSoft, C.greenDark, 6.5);

  const issues = data.dataQualityIssues.filter(i => !app || i.indexOf(app.app + ' ') === 0).length;
  if (issues) {
    const warn = '⚠ ' + issues + ' data warning' + (issues > 1 ? 's' : '') + ' — see Run Log';
    const warnW = estimateTextWidth_(warn, 6, true) + 4;
    addText_(slide, warn, pillX - warnW - 8, box.y + 2, warnW, 14, { fontSize: 6, bold: true, color: C.red, align: 'END', valign: 'MIDDLE' });
  }
}

function drawKpiCards_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const c = data.combined;
  const perApp = fn => data.apps.map(a => a.app + ' ' + fn(a.summary)).join(' · ');

  const cards = [
    { label: 'AVG DAU · ALL APPS', value: fmtNum_(c.avgDau), sub: perApp(s => fmtNum_(s.avgDau)), accent: C.blue },
    { label: data.feature.toUpperCase() + ' USERS / DAY', value: fmtNum_(c.avgFeatureUsers, 1), sub: perApp(s => fmtNum_(s.avgFeatureUsers, 1)), accent: C.greenBar },
    { label: 'CTR · DAU → ' + data.feature.toUpperCase(), value: fmtPct_(c.ctrPct, 2), sub: perApp(s => fmtPct_(s.ctrPct, 2)), accent: C.greenDark },
    c.uniqueUsers != null ? {
      label: 'UNIQUE USERS · ' + data.dataPeriod.days + ' DAYS',
      value: fmtNum_(c.uniqueUsers),
      sub: fmtNum_(c.repeatUsers) + ' repeat (active ≥2 days) · ' + fmtPct_(c.uniqueUsers ? (c.repeatUsers / c.uniqueUsers) * 100 : null, 0),
      accent: C.amber
    } : {
      label: 'FEATURE USER-DAYS',
      value: fmtNum_(c.totalFeatureUserDays),
      sub: 'sum of daily users · ' + data.dataPeriod.days + ' days',
      accent: C.amber
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

  const labelSize = fitFontSize_([card.label], w - 14, 8, 6, 5, 0, true);
  addText_(slide, card.label, x + 10, y + 5, w - 14, 8, { fontSize: labelSize, bold: true, color: C.muted });
  addText_(slide, card.value, x + 10, y + 14, w - 14, 16, { fontSize: 15, bold: true, color: C.ink, valign: 'MIDDLE' });
  const subSize = fitFontSize_([card.sub], w - 14, 7, 5.5, 4.5, 0);
  addText_(slide, card.sub, x + 10, y + 31, w - 14, 7, { fontSize: subSize, color: card.subColor || C.muted });
}

/**
 * DAU → feature funnel, one block per app. Each block: DAU bar (full track) and feature-user bar drawn to the
 * same scale, CTR with trend vs the prior window, and unique / repeat users for the period.
 */
function drawFunnel_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const apps = data.apps;

  panel_(slide, box);
  sectionTitle_(slide, 'FUNNEL  ·  DAU → KLIK ' + data.feature.toUpperCase(), box.x + 10, box.y + 8, 220);
  addText_(slide, 'Average per day · ' + data.dataPeriod.days + ' days', box.x + box.w - 160, box.y + 8, 150, 11, { fontSize: 6, color: C.muted, align: 'END', valign: 'MIDDLE' });

  const col = { app: box.x + 10, appW: 46, step: box.x + 62, stepW: 34, track: box.x + 100, ctr: box.x + box.w - 108, ctrW: 50, uniq: box.x + box.w - 56, uniqW: 46 };
  col.trackW = col.ctr - 10 - col.track;

  const headY = box.y + 24;
  [['APP', col.app, col.appW, 'START'], ['FUNNEL STEP (AVG / DAY)', col.step, col.trackW + col.stepW, 'START'],
    ['CTR', col.ctr, col.ctrW, 'CENTER'], ['UNIQUE USERS', col.uniq, col.uniqW, 'CENTER']]
    .forEach(h => addText_(slide, h[0], h[1], headY, h[2], 7, { fontSize: 5.5, bold: true, color: C.muted, align: h[3] }));
  hLine_(slide, box.x + 10, box.x + box.w - 10, headY + 9, C.border, 0.5);

  const top = box.y + 37, bottom = box.y + box.h - 15;
  const blockH = Math.min(46, (bottom - top) / apps.length);
  const rowH = Math.min(16, (blockH - 6) / 2);

  apps.forEach((a, i) => {
    const s = a.summary;
    const y0 = top + i * blockH;
    const rowA = y0, rowB = y0 + rowH + 2;
    if (i > 0) hLine_(slide, box.x + 10, box.x + box.w - 10, y0 - 2.5, C.grid, 0.5);

    // App chip
    pill_(slide, a.app, col.app, y0 + rowH + 1 - 7, col.appW, 14, a.color, C.white, 6.5);

    // Step 1: DAU (full track)
    addText_(slide, 'DAU', col.step, rowA, col.stepW, rowH, { fontSize: 6, color: C.muted, bold: true, valign: 'MIDDLE' });
    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, rowA + rowH / 2 - 6, col.trackW, 12, a.soft);
    addText_(slide, fmtNum_(s.avgDau), col.track, rowA, col.trackW - 6, rowH, { fontSize: 7, bold: true, color: a.color, align: 'END', valign: 'MIDDLE' });

    // Step 2: feature users, same scale as DAU
    addText_(slide, 'Klik', col.step, rowB, col.stepW, rowH, { fontSize: 6, color: C.ink, bold: true, valign: 'MIDDLE' });
    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, rowB + rowH / 2 - 6, col.trackW, 12, C.track);
    const fillW = s.avgFeatureUsers > 0 ? Math.max(4, col.trackW * (s.ctrPct || 0) / 100) : 0;
    if (fillW) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, rowB + rowH / 2 - 6, fillW, 12, a.color);
    const parts = [{ text: fmtNum_(s.avgFeatureUsers, 1), bold: true, color: C.ink }];
    if (s.avgNewUsers != null) {
      parts.push(
        { text: '   new ', color: C.muted }, { text: fmtNum_(s.avgNewUsers, 1), bold: true, color: C.ink },
        { text: '  ·  returning ', color: C.muted }, { text: fmtNum_(s.avgReturningUsers, 1), bold: true, color: C.ink },
        { text: ' (' + fmtPct_(s.returningSharePct, 0) + ')', color: C.muted }
      );
    }
    const ox = col.track + fillW + 5;
    addRichText_(slide, parts, ox, rowB, col.track + col.trackW - ox, rowH, { fontSize: 6, valign: 'MIDDLE' });

    // CTR + trend
    addText_(slide, fmtPct_(s.ctrPct, 2), col.ctr, rowA, col.ctrW, rowH, { fontSize: 11, bold: true, color: a.color, align: 'CENTER', valign: 'MIDDLE' });
    const t = s.trend;
    if (t && t.ctrChangePp != null) {
      const pp = t.ctrChangePp, up = pp > 0, flat = pp === 0;
      pill_(slide, (flat ? '■ ' : up ? '▲ +' : '▼ −') + fmtNum_(Math.abs(pp), 2) + ' pp', col.ctr + 2, rowB + rowH / 2 - 5.5, col.ctrW - 4, 11,
        flat ? C.cardBg : up ? C.okSoft : C.redSoft, flat ? C.muted : up ? C.okInk : C.red, 5.5);
    }

    // Unique / repeat users
    addText_(slide, s.uniqueUsers == null ? '–' : fmtNum_(s.uniqueUsers), col.uniq, rowA, col.uniqW, rowH, { fontSize: 11, bold: true, color: C.ink, align: 'CENTER', valign: 'MIDDLE' });
    if (s.uniqueUsers != null) {
      addText_(slide, fmtNum_(s.repeatUsers) + ' repeat', col.uniq, rowB, col.uniqW, rowH, { fontSize: 5.5, color: C.muted, align: 'CENTER', valign: 'MIDDLE' });
    }
  });

  const trendLabel = apps[0].summary.trend ? apps[0].summary.trend.label : 'prior period';
  addText_(slide, 'CTR = users who clicked ÷ DAU  ·  pp = CTR change, ' + trendLabel + '  ·  unique = distinct user_names; repeat = active on ≥2 days',
    box.x + 10, box.y + box.h - 11, box.w - 20, 7, { fontSize: 5, color: C.muted });
}

/** Daily CTR per app; point labels when the range is short enough to read them. */
function drawDailyChart_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const apps = data.apps;
  const dates = [];
  apps.forEach(a => a.dailySeries.forEach(r => dates.indexOf(r.date) < 0 && dates.push(r.date)));
  dates.sort();
  const n = dates.length;
  const detailed = n <= 14;

  panel_(slide, box);
  sectionTitle_(slide, 'DAILY CTR BY APP', box.x + 10, box.y + 8, 120);

  // Legend, right-aligned on the title row
  let lx = box.x + box.w - 10;
  apps.slice().reverse().forEach(a => {
    const label = a.app + '  (avg ' + fmtPct_(a.summary.ctrPct, 2) + ')';
    const tw = estimateTextWidth_(label, 6);
    lx -= tw;
    addText_(slide, label, lx, box.y + 8, tw, 11, { fontSize: 6, color: C.ink, valign: 'MIDDLE' });
    lx -= 13;
    hLine_(slide, lx, lx + 10, box.y + 13.5, a.color, 2);
    lx -= 12;
  });

  const plot = { x: box.x + 34, y: box.y + 32, w: box.w - 34 - 16, h: box.h - 32 - 16 };
  const maxCtr = niceMax_(Math.max(1, ...apps.map(a => Math.max(...a.dailySeries.map(r => r.ctrPct || 0)))));
  const band = plot.w / n;
  // Reserve room under the zero line for labels placed below low points
  const baseY = plot.y + plot.h - (detailed ? 9 : 0);
  const valueH = baseY - plot.y;
  const px = i => plot.x + (i + 0.5) * band;
  const py = v => baseY - (v / maxCtr) * valueH;

  [0, 0.5, 1].forEach(fr => {
    const gy = baseY - fr * valueH;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, fr === 0 ? C.border : C.grid, fr === 0 ? 0.75 : 0.5);
    if (fr > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(maxCtr * fr, 1) + '%', box.x + 4, gy - 4, plot.x - box.x - 8, 8, { fontSize: 5.5, color: C.muted, align: 'END', valign: 'MIDDLE' });
  });

  // Lines + markers per app
  const points = apps.map(a => {
    const byDate = {};
    a.dailySeries.forEach(r => byDate[r.date] = r.ctrPct);
    return dates.map(d => (byDate[d] == null ? null : byDate[d]));
  });
  apps.forEach((a, ai) => {
    const vals = points[ai];
    for (let i = 1; i < n; i++) {
      if (vals[i - 1] == null || vals[i] == null) continue;
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, px(i - 1), py(vals[i - 1]), px(i), py(vals[i]));
      line.getLineFill().setSolidFill(a.color);
      line.setWeight(1.75);
    }
    if (!detailed) return;
    vals.forEach((v, i) => {
      if (v == null) return;
      rect_(slide, SlidesApp.ShapeType.ELLIPSE, px(i) - 2.25, py(v) - 2.25, 4.5, 4.5, a.color, C.white, 0.75);
      // Label above the point for the higher series at this date, below for the lower one
      const others = points.filter((_, j) => j !== ai).map(p => p[i]).filter(x => x != null);
      const above = !others.length || v >= Math.max(...others);
      addText_(slide, fmtPct_(v, 1), px(i) - 15, above ? py(v) - 11 : py(v) + 3, 30, 8,
        { fontSize: 5.5, bold: true, color: a.color, align: 'CENTER', valign: 'MIDDLE' });
    });
  });

  // X-axis dates, thinned out for long ranges
  const every = Math.ceil(n / 8);
  const labelIdx = [];
  for (let i = 0; i < n; i += every) labelIdx.push(i);
  if (labelIdx[labelIdx.length - 1] !== n - 1) {
    // Always label the latest date; drop the previous label if the two would collide
    if (n - 1 - labelIdx[labelIdx.length - 1] < every * 0.6) labelIdx.pop();
    labelIdx.push(n - 1);
  }
  labelIdx.forEach(i => {
    addText_(slide, formatShortDate_(dates[i]), px(i) - 20, plot.y + plot.h + 3, 40, 8, { fontSize: 5.5, color: C.muted, align: 'CENTER' });
  });
}

function drawAppKpiCards_(slide, data, app, box) {
  const C = LS_CONFIG.COLORS;
  const s = app.summary;
  const t = s.trend;
  let trendSub = 'feature users ÷ DAU', trendColor = null;
  if (t && t.ctrChangePp != null) {
    const pp = t.ctrChangePp;
    trendSub = (pp > 0 ? '▲ +' : pp < 0 ? '▼ −' : '■ ') + fmtNum_(Math.abs(pp), 2) + ' pp vs previous ' + t.windowDays + ' days';
    trendColor = pp > 0 ? C.okInk : pp < 0 ? C.red : C.muted;
  }
  const cards = [
    { label: app.app.toUpperCase() + ' DAU', value: fmtNum_(s.avgDau), sub: 'avg users/day · ' + s.days + ' days', accent: C.blue },
    {
      label: data.feature.toUpperCase() + ' USERS / DAY',
      value: fmtNum_(s.avgFeatureUsers, 1),
      sub: s.peakFeatureUsersDay ? 'peak ' + fmtNum_(s.peakFeatureUsersDay.featureUsers) + ' on ' + formatShortDate_(s.peakFeatureUsersDay.date) : 'avg per day',
      accent: app.color
    },
    { label: 'CTR · DAU → ' + data.feature.toUpperCase(), value: fmtPct_(s.ctrPct, 2), sub: trendSub, subColor: trendColor, accent: C.greenDark },
    s.uniqueUsers != null ? {
      label: 'UNIQUE USERS · ' + s.days + ' DAYS',
      value: fmtNum_(s.uniqueUsers),
      sub: fmtNum_(s.repeatUsers) + ' repeat (active ≥2 days) · ' + fmtPct_(s.repeatRatePct, 0),
      accent: C.returning
    } : { label: 'FEATURE USER-DAYS', value: fmtNum_(s.totalFeatureUserDays), sub: 'sum of daily users', accent: C.returning }
  ];
  const gap = 8;
  const w = (box.w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => drawKpiCard_(slide, box.x + i * (w + gap), box.y, w, box.h, card));
}

/**
 * Single-app funnel: DAU → clicked the feature → came back (returning), all on one scale,
 * with drop-off notes between steps and a CTR callout on the right.
 */
function drawAppFunnel_(slide, data, app, box) {
  const C = LS_CONFIG.COLORS;
  const s = app.summary;
  const hasSplit = s.avgReturningUsers != null;

  panel_(slide, box);
  sectionTitle_(slide, 'FUNNEL  ·  DAU → KLIK ' + data.feature.toUpperCase() + (hasSplit ? ' → KEMBALI' : ''), box.x + 10, box.y + 8, 260);

  const callout = { x: box.x + box.w - 118, y: box.y + 26, w: 108, h: box.h - 34 };
  const col = { badge: box.x + 10, label: box.x + 28, track: box.x + 110 };
  col.labelW = col.track - col.label - 6;
  col.trackW = callout.x - 12 - col.track;

  const steps = [
    { name: 'Daily Active Users', value: s.avgDau, ratio: 1, color: app.soft, valueColor: app.color, note: 'avg users/day' },
    { name: 'Klik ' + data.feature, value: s.avgFeatureUsers, ratio: (s.ctrPct || 0) / 100, color: app.color, note: fmtPct_(s.ctrPct, 2) + ' of DAU' }
  ];
  if (hasSplit) {
    steps.push({
      name: 'Kembali lagi', value: s.avgReturningUsers, ratio: s.avgDau > 0 ? s.avgReturningUsers / s.avgDau : 0,
      color: C.returning, note: fmtPct_(s.returningSharePct, 0) + ' of clickers were returning'
    });
  }

  const rowH = 16, step = steps.length === 3 ? 27 : 36;
  const rowY = i => box.y + 28 + i * step;
  const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, col.badge + 6, rowY(0) + rowH / 2, col.badge + 6, rowY(steps.length - 1) + rowH / 2);
  connector.getLineFill().setSolidFill(C.border);
  connector.setWeight(1);

  steps.forEach((st, i) => {
    const y = rowY(i), cy = y + rowH / 2;
    rect_(slide, SlidesApp.ShapeType.ELLIPSE, col.badge, cy - 6, 12, 12, i === 0 ? app.color : st.color, C.white, 1);
    addText_(slide, String(i + 1), col.badge, cy - 6, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });
    addText_(slide, st.name, col.label, y, col.labelW, rowH, { fontSize: fitFontSize_([st.name], col.labelW, 9, 7, 6, 0, true), bold: true, color: C.ink, valign: 'MIDDLE' });

    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 6, col.trackW, 12, C.track);
    const fillW = st.value > 0 ? Math.max(4, col.trackW * st.ratio) : 0;
    if (fillW) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 6, fillW, 12, st.color);
    if (fillW >= 40) {
      addText_(slide, fmtNum_(st.value, 1), col.track, y, fillW - 6, rowH, { fontSize: 7, bold: true, color: st.valueColor || C.white, align: 'END', valign: 'MIDDLE' });
    } else {
      const ox = col.track + fillW + 5;
      addRichText_(slide, [
        { text: fmtNum_(st.value, 1), bold: true, color: C.ink },
        { text: '  ·  ' + st.note, color: C.muted }
      ], ox, y, col.track + col.trackW - ox, rowH, { fontSize: 6.5, valign: 'MIDDLE' });
    }
  });

  // Drop-off notes between steps
  const gapNote = (i, parts) => addRichText_(slide, parts, col.track, rowY(i) + rowH + 1, col.trackW, step - rowH - 2, { fontSize: 5.5, valign: 'MIDDLE' });
  if (s.ctrPct != null) {
    gapNote(0, [{ text: '▼ ' + fmtPct_(100 - s.ctrPct, 2), bold: true, color: C.red }, { text: ' of DAU did not open ' + data.feature, color: C.muted }]);
  }
  if (hasSplit && s.avgFeatureUsers > 0) {
    gapNote(1, [
      { text: fmtNum_(s.avgNewUsers, 1) + ' new', bold: true, color: app.color },
      { text: ' (' + fmtPct_(100 - s.returningSharePct, 0) + ') per day are first-time users', color: C.muted }
    ]);
  }

  // Best / lowest day
  if (s.bestDay && s.worstDay) {
    const statsY = box.y + box.h - 15;
    hLine_(slide, box.x + 10, callout.x - 12, statsY - 3, C.border, 0.5);
    addRichText_(slide, [
      { text: 'Best day  ', color: C.muted },
      { text: formatShortDate_(s.bestDay.date) + ' · ' + fmtPct_(s.bestDay.ctrPct, 2), bold: true, color: C.okInk },
      { text: '  (' + fmtNum_(s.bestDay.featureUsers) + ' of ' + fmtNum_(s.bestDay.dau) + ')', color: C.muted },
      { text: '      Lowest  ', color: C.muted },
      { text: formatShortDate_(s.worstDay.date) + ' · ' + fmtPct_(s.worstDay.ctrPct, 2), bold: true, color: C.red },
      { text: '  (' + fmtNum_(s.worstDay.featureUsers) + ' of ' + fmtNum_(s.worstDay.dau) + ')', color: C.muted }
    ], box.x + 10, statsY, callout.x - 22 - box.x, 10, { fontSize: 6, valign: 'MIDDLE' });
  }

  // CTR callout with benchmark against the other apps
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, callout.x, callout.y, callout.w, callout.h, app.soft);
  addText_(slide, 'CLICK-THROUGH RATE', callout.x + 6, callout.y + 7, callout.w - 12, 8, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
  addText_(slide, fmtPct_(s.ctrPct, 2), callout.x + 6, callout.y + 17, callout.w - 12, 22, { fontSize: 19, bold: true, color: app.color, align: 'CENTER', valign: 'MIDDLE' });
  addText_(slide, 'of ' + app.app + ' DAU open ' + data.feature, callout.x + 6, callout.y + 41, callout.w - 12, 8, { fontSize: 5.5, color: C.muted, align: 'CENTER' });

  const t = s.trend;
  if (t && t.ctrChangePp != null) {
    const pp = t.ctrChangePp, up = pp > 0, flat = pp === 0;
    pill_(slide, (flat ? '■ ' : up ? '▲ +' : '▼ −') + fmtNum_(Math.abs(pp), 2) + ' pp', callout.x + 14, callout.y + 53, callout.w - 28, 12,
      flat ? C.cardBg : up ? C.okSoft : C.redSoft, flat ? C.muted : up ? C.okInk : C.red, 6.5);
    addText_(slide, t.label, callout.x + 6, callout.y + 67, callout.w - 12, 7, { fontSize: 5, color: C.muted, align: 'CENTER' });
  }
  const others = data.apps.filter(a => a !== app);
  if (others.length) {
    addText_(slide, 'vs ' + others.map(o => o.app + ' ' + fmtPct_(o.summary.ctrPct, 2)).join(' · '),
      callout.x + 6, callout.y + callout.h - 12, callout.w - 12, 7, { fontSize: 5.5, bold: true, color: C.muted, align: 'CENTER' });
  }
}

/** Single app: feature users as stacked bars (new / returning, right axis), app DAU as a line (left axis). */
function drawAppDailyChart_(slide, app, box) {
  const C = LS_CONFIG.COLORS;
  const rows = app.dailySeries;
  const n = rows.length;
  const detailed = n <= 14;
  const hasSplit = rows.every(r => r.newUsers != null && r.returningUsers != null);

  panel_(slide, box);
  sectionTitle_(slide, 'DAILY TREND · ' + app.app.toUpperCase(), box.x + 10, box.y + 8, 130);

  const legend = [{ label: 'DAU (left)', color: C.blue, kind: 'line' }];
  if (hasSplit) legend.push({ label: 'New', color: app.color, kind: 'bar' }, { label: 'Returning (right)', color: C.returning, kind: 'bar' });
  else legend.push({ label: 'Feature users (right)', color: app.color, kind: 'bar' });
  if (detailed) legend.push({ label: '% = CTR', color: C.greenDark, kind: 'text' });
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
    lx -= 10;
  });

  const plot = { x: box.x + 34, y: box.y + 30, w: box.w - 34 - 34, h: box.h - 30 - (detailed ? 22 : 14) };
  const maxDau = niceMax_(Math.max(1, ...rows.map(r => r.dau)));
  // Bars use a taller scale so they stay in the lower part of the plot, clear of the DAU line
  const maxUsers = niceMax_(Math.max(1, ...rows.map(r => r.featureUsers)) * 1.6);
  const band = plot.w / n;
  const px = i => plot.x + (i + 0.5) * band;
  const yOf = (v, max) => plot.y + plot.h - (v / max) * plot.h;

  [0, 0.5, 1].forEach(fr => {
    const gy = plot.y + plot.h - fr * plot.h;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, fr === 0 ? C.border : C.grid, fr === 0 ? 0.75 : 0.5);
    if (fr > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(maxDau * fr, maxDau < 10 ? 1 : 0), box.x + 6, gy - 4, plot.x - box.x - 10, 8, { fontSize: 5.5, color: C.blue, align: 'END', valign: 'MIDDLE' });
    addText_(slide, fmtNum_(maxUsers * fr, 1), plot.x + plot.w + 4, gy - 4, 24, 8, { fontSize: 5.5, color: app.color, valign: 'MIDDLE' });
  });

  // Stacked bars: new at the bottom, returning on top
  const barW = Math.max(2, Math.min(16, band * 0.55));
  const base = plot.y + plot.h;
  rows.forEach((r, i) => {
    const segments = hasSplit ? [[r.newUsers, app.color], [r.returningUsers, C.returning]] : [[r.featureUsers, app.color]];
    let y = base;
    segments.forEach(seg => {
      const h = (seg[0] / maxUsers) * plot.h;
      if (h >= 0.5) rect_(slide, SlidesApp.ShapeType.RECTANGLE, px(i) - barW / 2, y - h, barW, h, seg[1]);
      y -= h;
    });
    if (detailed) {
      addText_(slide, fmtNum_(r.featureUsers), px(i) - 15, y - 9, 30, 8, { fontSize: 5.5, bold: true, color: C.ink, align: 'CENTER', valign: 'MIDDLE' });
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

  // X-axis: date (+ CTR when detailed), thinned out for long ranges
  const every = Math.ceil(n / 8);
  const labelIdx = [];
  for (let i = 0; i < n; i += every) labelIdx.push(i);
  if (labelIdx[labelIdx.length - 1] !== n - 1) {
    if (n - 1 - labelIdx[labelIdx.length - 1] < every * 0.6) labelIdx.pop();
    labelIdx.push(n - 1);
  }
  labelIdx.forEach(i => {
    const parts = [{ text: formatShortDate_(rows[i].date), color: C.muted }];
    if (detailed) parts.push({ text: '\n' + fmtPct_(rows[i].ctrPct, 2), color: C.greenDark, bold: true });
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
