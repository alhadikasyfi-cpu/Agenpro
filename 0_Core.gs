/**
 * SawitPRO Feature Analysis — core (Google Apps Script)
 *
 * One "Generate" run builds the weekly review slides of every feature listed in FEATURE_REGISTRY_().
 * Each feature lives in its own Feature_*.gs file (data reading, metrics, slide layout, Gemini context);
 * this file holds what they share: config, menu, generate flow, sheet reading, weekly metrics, Gemini and logging.
 * Slide components shared by the features are in 1_SlideComponents.gs.
 */

const CORE = {
  GEMINI_MODEL: 'gemini-3.5-flash-lite',
  PRESENTATION_PROPERTY: 'LAPAK_SAWIT_PRESENTATION_ID',
  API_KEY_PROPERTY: 'GEMINI_API_KEY',
  SHEET_LOG: 'Run Log',
  OUTPUT_LANGUAGE: 'English',
  FONT: 'Arial',

  // Reporting week: 0 = Sunday … 4 = Thursday … 6 = Saturday. Weeks run Thursday–Wednesday.
  // Only complete weeks are compared; days of a week still in progress wait for the next run.
  WEEK_START_DAY: 4,
  MAX_WEEKS: 8,
  // Extra Gemini rounds asking it to rewrite an answer that breaks the feature-only or PM-action rules
  GEMINI_MAX_RETRIES: 2,
  // Action items must be things a Product Manager ships or launches, not analysis work for a data analyst
  ANALYST_ACTION_PATTERNS: [/investigat/i, /analy[sz]/i, /\baudit/i, /look into/i, /find out/i, /understand why/i, /deep[- ]?dive/i,
    /root cause/i, /\bcheck (why|what|whether|if)\b/i, /\b(track|tracking|instrument|monitor)\b/i, /\binterview/i, /\bsurvey/i,
    /\bresearch\b/i, /cari tahu/i, /selidiki/i, /analisis/i, /\bcek\b/i],

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
    blueSoft: '#DBEAFE',
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
    white: '#FFFFFF',
    returning: '#E0A800',
    prevBar: '#CBD5E1',
    marker: '#7C3AED',
    nonPartner: '#94A3B8'
  },

  APPS: {
    agenpro: {
      label: 'AgenPRO',
      color: '#31512B',
      soft: '#DCE8D5',
      // Sent to Gemini: describe the users only, never other feature names (Gemini would build actions around them)
      description: 'App for agents (agen) who buy FFB/TBS from farmers and deliver it to mills, and who manage their own team ' +
        '(pengurus, drivers, members and plantation data).'
    },
    timpro: {
      label: 'TimPRO',
      color: '#C2410C',
      soft: '#FDE3CF',
      description: 'Loyalty and operational app for pengurus, the field coordinators working under an agent who organise FFB/TBS ' +
        'deliveries to mills.'
    }
  },

  // Every SawitPRO feature: `terms` find mentions of it in AI text, so each slide's text stays about its own feature.
  FEATURE_TERMS: {
    'Pasar Sawit': [/pasar\s*sawit/i, /lapak\s*sawit/i],
    'Pricing Intelligence': [/pricing\s*intelligence/i],
    'Affiliate': [/affiliat/i, /afiliasi/i, /referral/i, /toko\s*sawit/i],
    'Protim': [/\bprotim\b/i],
    'Panen Emas': [/panen\s*emas/i],
    'Laporan TBS': [/laporan\s*(tbs|pengiriman)/i, /delivery\s*report/i],
    'FFB Transaction': []
  },

  // Layout grid in points for a 16:9 slide (720 x 405 pt).
  // Content stays between the template header line (~50pt) and the footer bar (~370pt).
  LAYOUT: {
    title: { x: 20, y: 54, w: 680, h: 18 },
    kpi: { x: 20, y: 78, w: 420, h: 42 },
    main: { x: 20, y: 126, w: 420, h: 124 },
    chart: { x: 20, y: 256, w: 420, h: 108 },
    observations: { x: 452, y: 78, w: 248, h: 162 },
    actions: { x: 452, y: 246, w: 248, h: 118 }
  }
};

/** Features in slide order. Each Feature_*.gs file defines one of these functions. */
function FEATURE_REGISTRY_() {
  return [
    pasarSawitFeature_(),
    pricingIntelligenceFeature_(),
    affiliateFeature_(),
    protimFeature_(),
    ffbTransactionFeature_()
  ];
}

function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('Feature Analysis')
    .addItem('Generate all feature slides', 'generateAllFeatureSlides')
    .addSeparator();
  FEATURE_REGISTRY_().forEach(f => menu.addItem('Only: ' + f.name, f.menuFunction));
  menu.addSeparator()
    .addItem('Set Gemini API key', 'setGeminiApiKey')
    .addItem('Open generated presentation', 'openGeneratedPresentation')
    .addToUi();
}

function generateAllFeatureSlides() {
  generateFeatures_(null);
}

/**
 * Builds the slides of the given features (all when ids is null): reads each feature's sheet, asks Gemini for every
 * slide in parallel, then duplicates the template (slide 1) once per slide, keeping this run's slides in order after it.
 * A feature whose sheet fails is skipped and reported; the others still generate.
 */
function generateFeatures_(ids) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Process is running. Please wait.');

  let presentationUrl = '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const features = FEATURE_REGISTRY_().filter(f => !ids || ids.indexOf(f.id) >= 0);
    const specs = [], failed = [], warnings = [];
    features.forEach(f => {
      const issues = [];
      try {
        f.buildSlides(ss, issues).forEach(s => specs.push(s));
      } catch (err) {
        failed.push(f.name + ': ' + (err && err.message ? err.message : err));
      }
      issues.forEach(i => warnings.push(f.name + ' — ' + i));
    });
    if (!specs.length) throw new Error('No slides generated. ' + failed.join(' | '));

    const aiResults = callGeminiAll_(specs);

    const presentation = getOrCreatePresentation_();
    const slides = presentation.getSlides();
    if (!slides.length) throw new Error('Presentation missing master template on Slide 1.');
    specs.forEach((spec, i) => {
      const newSlide = slides[0].duplicate();
      newSlide.move(1 + i);
      spec.draw(newSlide, aiResults[i]);
      if (aiResults[i].speakerNotes) newSlide.getNotesPage().getSpeakerNotesShape().getText().setText(aiResults[i].speakerNotes);
      if (aiResults[i].removedOffTopic) {
        warnings.push(spec.title + ' — AI text removed: other features or analysis-type actions (' + aiResults[i].removedOffTopic + ' item(s))');
      }
    });

    presentationUrl = 'https://docs.google.com/presentation/d/' + presentation.getId() + '/edit';
    const message = specs.length + ' slide(s): ' + specs.map(s => s.title).join(', ') + '.' +
      (failed.length ? ' Skipped: ' + failed.join(' | ') + '.' : '') +
      (warnings.length ? ' Warnings: ' + warnings.join('; ') : '');
    logRun_(ss, failed.length || warnings.length ? 'SUCCESS_WITH_WARNINGS' : 'SUCCESS', message, presentationUrl);

    SpreadsheetApp.getUi().alert(specs.length + ' slide(s) generated.' +
      (failed.length ? '\n\nSkipped:\n' + failed.join('\n') : '') +
      (warnings.length ? '\n\n' + warnings.length + ' warning(s), see Run Log.' : '') +
      '\n\nOpen presentation:\n' + presentationUrl);
  } catch (err) {
    try {
      logRun_(ss, 'ERROR', String(err && err.message ? err.message : err), presentationUrl);
    } catch (_) {}
    SpreadsheetApp.getUi().alert('Error: ' + (err && err.message ? err.message : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ SHEET READING ------------------------------ */

/**
 * Reads one table from a sheet. `fields` maps a field name to accepted headers (normalised: lower-case,
 * non-alphanumerics → "_"); a header may also be matched by a RegExp. Rows are read from row 2 down to the first row
 * where every matched column is empty, so a second table next to it (other columns) does not affect where it ends.
 * Returns { rows: [{ rowNumber, field: value }], headers: { field: original header text } }.
 */
function readTable_(ss, sheetName, fields, required) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet "' + sheetName + '".');
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('Sheet "' + sheetName + '" has no data rows.');

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const raw = values[0].map(h => String(h == null ? '' : h));
  const headers = raw.map(normalizeHeader_);
  const col = {};
  Object.keys(fields).forEach(field => {
    col[field] = headers.findIndex(h => h && fields[field].some(a => (a instanceof RegExp ? a.test(h) : a === h)));
  });
  (required || []).forEach(field => {
    if (col[field] < 0) throw new Error('Sheet "' + sheetName + '" is missing column for "' + field + '".');
  });

  const used = Object.keys(col).filter(f => col[f] >= 0);
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (used.every(f => row[col[f]] === '' || row[col[f]] === null)) break; // table ends at the first empty row
    const out = { rowNumber: r + 1 };
    used.forEach(f => out[f] = row[col[f]]);
    rows.push(out);
  }
  if (!rows.length) throw new Error('Sheet "' + sheetName + '" has no data before the first empty row.');
  const headerText = {};
  used.forEach(f => headerText[f] = raw[col[f]]);
  return { rows: rows, headers: headerText };
}

function numOrNull_(v) {
  return v === '' || v == null ? null : toNumber_(v);
}

/* ------------------------------ DAILY SERIES ------------------------------ */

/**
 * Turns raw table rows into a validated daily series, one row per date, oldest first:
 * { date, dau, steps: { key: n }, newUsers, returningUsers, names, appWau, featureWau, ctrPct }.
 * `stepKeys` are the feature's funnel steps; the first one is the "feature users" step (names, new / returning).
 * Weekly WAU columns are filled only on a week's last day; 0 there means "not filled".
 */
function toDailySeries_(rows, stepKeys, tag, issues) {
  const byDate = {};
  rows.forEach(r => {
    const d = parseDate_(r.date);
    if (!d) {
      issues.push(tag + 'row ' + r.rowNumber + ': unreadable date (row skipped)');
      return;
    }
    const key = dateKey_(d);
    if (byDate[key]) {
      issues.push(tag + 'duplicate date ' + key + ' (row ' + r.rowNumber + ' skipped)');
      return;
    }
    const steps = {};
    stepKeys.forEach(k => steps[k] = toNumber_(r[k]));
    const primary = steps[stepKeys[0]];
    const dau = toNumber_(r.dau);
    const names = r.names != null ? parseNames_(r.names) : null;
    const newUsers = numOrNull_(r.newUsers), returningUsers = numOrNull_(r.returningUsers);
    if (primary > dau) issues.push(tag + key + ': feature users (' + primary + ') exceed DAU (' + dau + ')');
    if (newUsers != null && returningUsers != null && newUsers + returningUsers !== primary) {
      issues.push(tag + key + ': new + returning (' + (newUsers + returningUsers) + ') ≠ feature users (' + primary + ')');
    }
    if (names && names.length && names.length !== primary) {
      issues.push(tag + key + ': ' + names.length + ' user names vs ' + primary + ' feature users');
    }
    const ids = names ? names.filter(n => /^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(n)) : [];
    if (ids.length) issues.push(tag + key + ': ' + ids.length + ' user ID(s) instead of a name in the user list');
    byDate[key] = {
      date: key,
      dau: dau,
      steps: steps,
      newUsers: newUsers,
      returningUsers: returningUsers,
      names: names || [],
      appWau: numOrNull_(r.appWau) || null,
      featureWau: numOrNull_(r.featureWau) || null,
      ctrPct: dau > 0 ? round2_((primary / dau) * 100) : null
    };
  });
  const series = Object.keys(byDate).sort().map(k => byDate[k]);
  if (series.length) {
    const span = Math.round((parseDate_(series[series.length - 1].date) - parseDate_(series[0].date)) / 86400000) + 1;
    if (span > series.length) issues.push(tag + (span - series.length) + ' missing date(s) between ' + series[0].date + ' and ' + series[series.length - 1].date);
  }
  return series;
}

/* ------------------------------ WEEKLY METRICS ------------------------------ */

const WEEK_METRIC_KINDS = { conversionPct: 'pct', ctrPct: 'pct', retentionPct: 'pct', completionPct: 'pct', visitsPerUser: 'ratio' };

/**
 * Complete weeks over a daily series, this week vs last week.
 * Returns null when there is not one full week yet. Feature users per week come from the sheet's feature WAU when
 * filled, otherwise from distinct names; retention and churn need names.
 */
function weeklyAnalysis_(series, stepKeys, tag, issues) {
  if (!series.length) return null;
  const windows = weekWindows_(series[0].date, series[series.length - 1].date);
  if (!windows.length) return null;
  const hasNames = series.some(r => r.names.length);

  const raws = windows.map(w => {
    const rows = series.filter(r => r.date >= w.start && r.date <= w.end);
    const nameDays = {};
    rows.forEach(r => r.names.forEach(n => {
      const k = nameKey_(n);
      nameDays[k] = nameDays[k] || {};
      nameDays[k][r.date] = true;
    }));
    const stepTotals = {};
    stepKeys.forEach(k => stepTotals[k] = rows.reduce((acc, r) => acc + r.steps[k], 0));
    const raw = {
      start: w.start, end: w.end, days: rows.length,
      totalDau: rows.reduce((acc, r) => acc + r.dau, 0),
      stepTotals: stepTotals,
      newFromSheet: rows.length && rows.every(r => r.newUsers != null) ? rows.reduce((acc, r) => acc + r.newUsers, 0) : null,
      appWau: rows.filter(r => r.appWau != null).map(r => r.appWau).pop() || null,
      featureWau: rows.filter(r => r.featureWau != null).map(r => r.featureWau).pop() || null,
      nameDays: nameDays
    };
    const wk = tag + 'week ' + formatRange_(w.start, w.end);
    if (raw.days < 7) issues.push(wk + ': only ' + raw.days + ' of 7 days have data');
    if (raw.appWau == null && raw.days === 7) issues.push(wk + ': app WAU is empty on the last day of the week');
    const named = Object.keys(nameDays).length;
    if (hasNames && raw.featureWau != null && named !== raw.featureWau) {
      issues.push(wk + ': feature WAU ' + raw.featureWau + ' vs ' + named + ' distinct user names');
    }
    return raw;
  });

  const seen = {};
  series.filter(r => r.date < windows[0].start).forEach(r => r.names.forEach(n => seen[nameKey_(n)] = true));
  const weeks = raws.map((raw, i) => {
    const prev = i > 0 ? raws[i - 1] : null;
    const names = hasNames ? Object.keys(raw.nameDays) : null;
    const clickUsers = raw.featureWau != null ? raw.featureWau : names ? names.length : null;
    let newUsers = raw.newFromSheet;
    if (newUsers == null && names) newUsers = names.filter(k => !seen[k]).length;
    const prevNamed = prev && names ? Object.keys(prev.nameDays).length : null;
    const retained = prev && names ? names.filter(k => prev.nameDays[k]).length : null;
    const returning = clickUsers != null && newUsers != null ? Math.max(0, clickUsers - newUsers) : null;
    if (names) names.forEach(k => seen[k] = true);
    const visits = raw.stepTotals[stepKeys[0]];
    const last = raw.stepTotals[stepKeys[stepKeys.length - 1]];
    const week = {
      start: raw.start,
      end: raw.end,
      days: raw.days,
      appWau: raw.appWau,
      clickUsers: clickUsers,
      conversionPct: raw.appWau && clickUsers != null ? round2_((clickUsers / raw.appWau) * 100) : null,
      avgDau: raw.days ? Math.round(raw.totalDau / raw.days) : null,
      totalDau: raw.totalDau,
      visits: visits,
      ctrPct: raw.totalDau > 0 ? round2_((visits / raw.totalDau) * 100) : null,
      completionPct: stepKeys.length > 1 && visits > 0 ? round2_((last / visits) * 100) : null,
      newUsers: newUsers,
      returningUsers: returning,
      prevClickUsers: prevNamed,
      retainedFromPrev: retained,
      retentionPct: prevNamed ? round1_((retained / prevNamed) * 100) : null,
      churnedFromPrev: prevNamed != null ? prevNamed - retained : null,
      repeatInWeek: names ? names.filter(k => Object.keys(raw.nameDays[k]).length >= 2).length : null,
      visitsPerUser: clickUsers ? round1_(visits / clickUsers) : null,
      appStickinessPct: raw.appWau && raw.days ? round1_((raw.totalDau / raw.days / raw.appWau) * 100) : null,
      activeDaysSplit: names ? {
        oneDay: names.filter(k => Object.keys(raw.nameDays[k]).length === 1).length,
        twoDays: names.filter(k => Object.keys(raw.nameDays[k]).length === 2).length,
        threePlusDays: names.filter(k => Object.keys(raw.nameDays[k]).length >= 3).length
      } : null
    };
    stepKeys.forEach(k => week['step_' + k] = raw.stepTotals[k]);
    return week;
  });

  const lastFull = windows[windows.length - 1].end;
  return {
    windows: windows,
    weeks: weeks,
    compare: compareWeeks_(weeks),
    daysInProgress: series.filter(r => r.date > lastFull).length
  };
}

/** Complete weeks, oldest first, from WEEK_START_DAY; the latest one ends on or before lastKey. */
function weekWindows_(firstKey, lastKey) {
  const first = parseDate_(firstKey);
  let last = parseDate_(lastKey);
  const endDay = (CORE.WEEK_START_DAY + 6) % 7;
  while (last.getDay() !== endDay) last = addDays_(last, -1);
  const out = [];
  for (let k = 0; k < CORE.MAX_WEEKS; k++) {
    const end = addDays_(last, -7 * k), start = addDays_(end, -6);
    if (start.getTime() < first.getTime()) break;
    out.unshift({ start: dateKey_(start), end: dateKey_(end) });
  }
  return out;
}

/** This week vs last week for every numeric metric: absolute difference (pp for percentages) and relative change. */
function compareWeeks_(weeks) {
  if (!weeks.length) return null;
  const current = weeks[weeks.length - 1], previous = weeks.length > 1 ? weeks[weeks.length - 2] : null;
  const deltas = {};
  Object.keys(current).forEach(k => {
    const a = current[k], b = previous ? previous[k] : null;
    if (typeof a !== 'number' || typeof b !== 'number') return;
    const pct = WEEK_METRIC_KINDS[k] === 'pct';
    deltas[k] = { diff: round2_(a - b), relPct: pct ? null : pctChange_(a, b) };
  });
  return { current: current, previous: previous, deltas: deltas };
}

function metricKind_(key) {
  return WEEK_METRIC_KINDS[key] === 'pct' ? { kind: 'pct' } : WEEK_METRIC_KINDS[key] === 'ratio' ? { kind: 'count', digits: 1 } : { kind: 'count' };
}

/* ------------------------------ GEMINI ------------------------------ */

/**
 * One Gemini request per slide, sent in parallel. Each spec.ai = { feature, app (CORE.APPS entry or null), description,
 * scope: [lines], definitions: {metric: meaning}, payload, extraRules: [lines], fallbackHeadline }.
 * Answers that mention other features are sent back for a rewrite; anything left after the retries is removed.
 */
function callGeminiAll_(specs) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(CORE.API_KEY_PROPERTY);
  if (!apiKey) throw new Error('Gemini API key missing in Script Properties.');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(CORE.GEMINI_MODEL) + ':generateContent';
  const request = contents => ({
    url: url,
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify({ contents: contents, generationConfig: { temperature: 0.2, responseMimeType: 'application/json' } }),
    muteHttpExceptions: true
  });

  const conversations = specs.map(s => [{ role: 'user', parts: [{ text: buildGeminiPrompt_(s.ai) }] }]);
  const results = [];
  let pending = specs.map((_, i) => i);
  for (let attempt = 0; pending.length && attempt <= CORE.GEMINI_MAX_RETRIES; attempt++) {
    const responses = UrlFetchApp.fetchAll(pending.map(i => request(conversations[i])));
    const retry = [];
    responses.forEach((response, k) => {
      const i = pending[k];
      if (response.getResponseCode() !== 200) {
        throw new Error('Gemini API error (' + specs[i].title + '): ' + response.getContentText());
      }
      const text = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text;
      results[i] = normalizeAi_(JSON.parse(text));
      const feature = specs[i].ai.feature;
      const offending = offTopicTexts_(results[i], feature).concat(analystActions_(results[i]));
      if (offending.length && attempt < CORE.GEMINI_MAX_RETRIES) {
        conversations[i].push(
          { role: 'model', parts: [{ text: text }] },
          { role: 'user', parts: [{ text: 'Your answer breaks the slide rules (' + feature + ' is the only feature you may mention, and ' +
            'actions must be Product Manager actions). These parts break them:\n- ' + offending.join('\n- ') + '\n' +
            'Rewrite the whole JSON. Every observation and action must be about ' + feature + ' only, and every action must be a ' +
            'concrete product change or campaign a Product Manager can ship — no investigation, analysis, tracking or research tasks.' }] }
        );
        retry.push(i);
      }
    });
    pending = retry;
  }
  return results.map((ai, i) => removeOffTopic_(ai, specs[i].ai));
}

function buildGeminiPrompt_(a) {
  const defs = Object.keys(a.definitions || {}).map(k => '- ' + k + ': ' + a.definitions[k]);
  return [
    'You are a senior product analytics consultant at SawitPRO (palm-oil / FFB supply chain platform in Indonesia).',
    'You are writing one review slide about the "' + a.feature + '" feature' + (a.app ? ' in the ' + a.app.label + ' app' : '') + '.',
    '',
    'PRODUCT CONTEXT',
    'The feature: ' + a.feature + ' — ' + a.description,
    a.app ? 'The app: ' + a.app.label + ' — ' + a.app.description : null,
    '',
    'SLIDE SCOPE',
    a.scope.join('\n'),
    '',
    'METRIC DEFINITIONS',
    defs.join('\n'),
    '',
    'Output must be valid JSON matching this schema:',
    '{',
    '  "headline": "Clear single executive takeaway sentence",',
    '  "observations": ["Observation 1", "Observation 2", "Observation 3"],',
    '  "next_steps": [{"action": "Actionable item description", "pic": "Owner/Team"}],',
    '  "speaker_notes": "A detailed presentation script for this slide: walk through every number on it, what changed and why it may have changed, data quality warnings, and the reasoning behind each recommendation."',
    '}',
    '',
    'Rules (the slide has limited space, keep on-slide text short):',
    '- Write in ' + CORE.OUTPUT_LANGUAGE + '.',
    '- Ground every statement in DATA and quote numbers exactly as given; when comparing periods, quote both numbers.',
    '- Use plain business language. Never write JSON field names; write dates like "22 Oct".',
    '- "WAU" always means the APP\'s weekly active users (appWau). Never call the feature\'s users "WAU".',
    '- If only one or two periods can be compared, say trends are early signals.',
    '- If dataQualityIssues is not empty, mention the most important one in an observation.'
  ].concat(a.extraRules || []).concat([
    '',
    'ACTION PLAN RULES (next_steps):',
    '- Actions are for the Product Manager: concrete things they can ship, change or launch for the ' + a.feature + ' feature itself ' +
      'to get more users to use it, get them to come back, or remove friction in its flow.',
    '- Good action types: a ' + a.feature + ' campaign or promo, ads, push notification / WhatsApp blast to a specific user segment, ' +
      'in-app banner, reminder or entry point, onboarding or copy change, UX change to a funnel step, incentive, content or listing change.',
    '- Do NOT recommend analysis or research work for a data analyst: no investigating, analysing, auditing, checking why, ' +
      'tracking or instrumentation, monitoring, interviews or surveys. State the product action directly instead.',
    '- ' + a.feature + ' is the ONLY feature you may name. Do not mention, link to, bundle with or recommend changes to any other ' +
      'SawitPRO feature, program, report or module. This applies to the headline, observations and speaker notes too.',
    '- At least one action must respond to the biggest change or drop-off in DATA with a concrete product change or campaign.',
    '- Each action must follow from a number in DATA; mention that number briefly when it fits.',
    '- headline: one sentence, max 20 words.',
    '- observations: exactly 3 items, each max 24 words, no leading bullet characters.',
    '- next_steps: exactly 3 items; "action" max 16 words; "pic" is a short team name (max 4 words).',
    '- speaker_notes: no length limit.',
    '',
    'DATA:',
    JSON.stringify(a.payload)
  ]).filter(l => l !== null).join('\n');
}

/** Patterns of every feature other than this one (skipping any that would match this feature's own name). */
function offTopicPatterns_(feature) {
  return Object.keys(CORE.FEATURE_TERMS)
    .filter(name => name.toLowerCase() !== feature.toLowerCase())
    .reduce((acc, name) => acc.concat(CORE.FEATURE_TERMS[name]), [])
    .filter(re => !re.test(feature));
}

function isOffTopic_(text, feature) {
  return offTopicPatterns_(feature).some(re => re.test(String(text || '')));
}

/** Slide texts (and speaker-note sentences) that mention another feature. */
function offTopicTexts_(ai, feature) {
  const texts = [ai.headline].concat(ai.observations, ai.nextSteps.map(s => s.action + ' [PIC: ' + s.pic + ']'), splitSentences_(ai.speakerNotes));
  return texts.filter(t => isOffTopic_(t, feature));
}

/** Action items that are analysis / research work rather than a product action. */
function analystActions_(ai) {
  return ai.nextSteps.map(s => s.action).filter(t => CORE.ANALYST_ACTION_PATTERNS.some(re => re.test(t)));
}

/** Last resort after the retries: drop whatever still mentions another feature, and analysis-type action items. */
function removeOffTopic_(ai, a) {
  const before = offTopicTexts_(ai, a.feature).length + analystActions_(ai).length;
  if (!before) return ai;
  return {
    headline: isOffTopic_(ai.headline, a.feature) ? a.fallbackHeadline : ai.headline,
    observations: ai.observations.filter(o => !isOffTopic_(o, a.feature)),
    nextSteps: ai.nextSteps.filter(s => !isOffTopic_(s.action + ' ' + s.pic, a.feature) && !CORE.ANALYST_ACTION_PATTERNS.some(re => re.test(s.action))),
    speakerNotes: splitSentences_(ai.speakerNotes).filter(t => !isOffTopic_(t, a.feature)).join(' '),
    removedOffTopic: before
  };
}

/* ------------------------------ HELPERS ------------------------------ */

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

/** "15–21 Oct", or "29 Sep – 5 Oct" across months. */
function formatRange_(startKey, endKey) {
  const tz = Session.getScriptTimeZone();
  const start = parseDate_(startKey), end = parseDate_(endKey);
  if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
    return Utilities.formatDate(start, tz, 'd') + '–' + Utilities.formatDate(end, tz, 'd MMM');
  }
  return Utilities.formatDate(start, tz, 'd MMM') + ' – ' + Utilities.formatDate(end, tz, 'd MMM');
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

function splitSentences_(text) {
  return String(text || '').split(/(?<=[.!?])\s+/).filter(Boolean);
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

function normalizeHeader_(h) {
  return String(h == null ? '' : h).trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
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

function pctChange_(current, previous) {
  return previous > 0 ? round1_(((current - previous) / previous) * 100) : null;
}

function addDays_(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

function getOrCreatePresentation_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(CORE.PRESENTATION_PROPERTY);
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
  const sheet = getOrCreateSheet_(ss, CORE.SHEET_LOG);
  sheet.appendRow([new Date(), status, message, url || '']);
}

function setGeminiApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt('Gemini API Key', 'Paste your Gemini API key:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const key = result.getResponseText().trim();
  if (key) {
    PropertiesService.getScriptProperties().setProperty(CORE.API_KEY_PROPERTY, key);
    ui.alert('Gemini API key saved.');
  }
}

function openGeneratedPresentation() {
  const id = PropertiesService.getScriptProperties().getProperty(CORE.PRESENTATION_PROPERTY);
  if (!id) {
    SpreadsheetApp.getUi().alert('No presentation generated yet.');
    return;
  }
  const url = 'https://docs.google.com/presentation/d/' + id + '/edit';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<p>Open presentation:</p><p><a target="_blank" href="' + url + '">' + url + '</a></p>').setWidth(460).setHeight(130),
    'Feature Analysis Presentation'
  );
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

