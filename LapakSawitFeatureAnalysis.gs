/**
 * Lapak Sawit Feature Analysis — Google Apps Script
 * Duplicates the SawitPRO master template (slide 1) and renders a feature analysis snapshot:
 * KPI cards, stage-by-stage funnel, DAU trend chart and Gemini-generated insights.
 */

const LS_CONFIG = {
  GEMINI_MODEL: 'gemini-3.5-flash-lite',
  PRESENTATION_PROPERTY: 'LAPAK_SAWIT_PRESENTATION_ID',
  API_KEY_PROPERTY: 'GEMINI_API_KEY',
  SHEET_EVENTS: 'Events',
  SHEET_DAILY: 'Daily Metrics',
  SHEET_LOG: 'Run Log',
  TITLE: 'Lapak Sawit Feature Analysis',
  FONT: 'Arial',

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
  // One shade per funnel stage, darkest at the top of the funnel
  FUNNEL_COLORS: ['#31512B', '#4A7340', '#6E9A5E', '#9CBF8C'],
  FUNNEL_STAGES: [
    { name: 'Home Open', aliases: ['home open', 'home_open', 'home_opened'] },
    { name: 'Daftar RAM', aliases: ['daftar ram', 'daftar_ram', 'ram registration'] },
    { name: 'Jual TBS', aliases: ['jual tbs', 'jual_tbs', 'sell ffb'] },
    { name: 'Permintaan Penjemputan', aliases: ['permintaan penjemputan', 'permintaan_penjemputan', 'pickup request'] }
  ],

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
  const events = getOrCreateSheet_(ss, LS_CONFIG.SHEET_EVENTS);
  const daily = getOrCreateSheet_(ss, LS_CONFIG.SHEET_DAILY);
  const log = getOrCreateSheet_(ss, LS_CONFIG.SHEET_LOG);

  setupTab_(events, ['date', 'event_name', 'event_count'], [
    ['2026-10-01', 'Home Open', 120],
    ['2026-10-01', 'Daftar RAM', 18],
    ['2026-10-01', 'Jual TBS', 3],
    ['2026-10-01', 'Permintaan Penjemputan', 0]
  ]);
  setupTab_(daily, ['date', 'dau', 'returning_users', 'new_users'], [
    ['2026-10-01', 120, 42, 78],
    ['2026-10-02', 135, 48, 87],
    ['2026-10-03', 128, 45, 83]
  ]);
  setupTab_(log, ['run_at', 'status', 'message', 'presentation_url'], []);
  [events, daily, log].forEach(sh => {
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
    const events = readEvents_(ss);
    const daily = readDailyMetrics_(ss);

    const analysisInput = buildAnalysisInput_(events, daily);
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
    logRun_(ss, 'SUCCESS', 'Slide generated successfully.', presentationUrl);

    SpreadsheetApp.getUi().alert('Slide generated successfully!\n\nOpen presentation:\n' + presentationUrl);
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

function readTable_(sheetName, requiredHeaders) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing sheet "' + sheetName + '".');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const indexes = {};
  requiredHeaders.forEach(h => {
    const idx = headers.indexOf(h);
    if (idx < 0) throw new Error('Sheet missing header "' + h + '".');
    indexes[h] = idx;
  });
  return values.slice(1).filter(row => row.some(v => v !== '' && v !== null)).map(row => {
    const out = {};
    requiredHeaders.forEach(h => out[h] = row[indexes[h]]);
    return out;
  });
}

function readEvents_(ss) {
  return readTable_(LS_CONFIG.SHEET_EVENTS, ['date', 'event_name', 'event_count'])
    .map(r => ({ date: parseDate_(r.date), eventName: String(r.event_name || '').trim(), eventCount: toNumber_(r.event_count) }))
    .filter(r => r.date && r.eventName);
}

function readDailyMetrics_(ss) {
  return readTable_(LS_CONFIG.SHEET_DAILY, ['date', 'dau', 'returning_users', 'new_users'])
    .map(r => ({ date: parseDate_(r.date), dau: toNumber_(r.dau), returningUsers: toNumber_(r.returning_users), newUsers: toNumber_(r.new_users) }))
    .filter(r => r.date);
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (!value) return null;
  const s = String(value).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function toNumber_(value) {
  const n = Number(String(value == null ? '' : value).replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}

function buildAnalysisInput_(events, daily) {
  const stages = LS_CONFIG.FUNNEL_STAGES.map(stage => {
    const aliases = stage.aliases.map(normalizeEvent_);
    const matching = events.filter(e => aliases.indexOf(normalizeEvent_(e.eventName)) >= 0);
    const total = matching.reduce((sum, e) => sum + e.eventCount, 0);
    return { name: stage.name, eventCount: total };
  });

  const top = stages[0].eventCount;
  stages.forEach((stage, i) => {
    const prev = i > 0 ? stages[i - 1].eventCount : null;
    stage.stepConversionPct = i === 0 ? 100 : (prev > 0 ? round1_((stage.eventCount / prev) * 100) : null);
    stage.shareOfTopPct = top > 0 ? round1_((stage.eventCount / top) * 100) : null;
    stage.lostFromPrev = i === 0 ? 0 : Math.max(0, prev - stage.eventCount);
  });

  // Largest leak = the step that loses the most events in absolute terms
  let largestLeakIndex = -1;
  stages.forEach((stage, i) => {
    if (i > 0 && stage.lostFromPrev > 0 && (largestLeakIndex < 0 || stage.lostFromPrev > stages[largestLeakIndex].lostFromPrev)) {
      largestLeakIndex = i;
    }
  });

  const sorted = daily.slice().sort((a, b) => a.date.getTime() - b.date.getTime());
  const series = sorted.map(r => ({
    date: dateKey_(r.date),
    dau: r.dau,
    returningUsers: r.returningUsers,
    newUsers: r.newUsers
  }));
  if (!series.length) throw new Error('Sheet "' + LS_CONFIG.SHEET_DAILY + '" has no valid rows.');

  const sum = key => series.reduce((acc, r) => acc + r[key], 0);
  const days = series.length;
  const totalUsers = sum('newUsers') + sum('returningUsers');
  const last = stages[stages.length - 1].eventCount;

  return {
    dataPeriod: { start: series[0].date, end: series[days - 1].date },
    summary: {
      avgDau: Math.round(sum('dau') / days),
      avgNewUsers: Math.round(sum('newUsers') / days),
      avgReturningUsers: Math.round(sum('returningUsers') / days),
      newUserSharePct: totalUsers > 0 ? round1_((sum('newUsers') / totalUsers) * 100) : null,
      totalHomeOpen: top,
      totalTransactions: stages[2].eventCount,
      totalPickupRequests: stages[3].eventCount,
      transactionsShareOfHomePct: stages[2].shareOfTopPct,
      pickupShareOfTransactionsPct: stages[3].stepConversionPct,
      overallConversionPct: top > 0 ? round1_((last / top) * 100) : null,
      largestLeak: largestLeakIndex > 0 ? {
        index: largestLeakIndex,
        from: stages[largestLeakIndex - 1].name,
        to: stages[largestLeakIndex].name,
        lost: stages[largestLeakIndex].lostFromPrev
      } : null
    },
    funnel: stages,
    dailySeries: series
  };
}

function normalizeEvent_(s) {
  return String(s || '').toLowerCase().trim().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
}

function dateKey_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

/* ------------------------------ GEMINI API ------------------------------ */

function callGemini_(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(LS_CONFIG.API_KEY_PROPERTY);
  if (!apiKey) throw new Error('Gemini API key missing in Script Properties.');

  const prompt = [
    'You are a senior product analytics consultant analyzing the Lapak Sawit feature performance for SawitPRO.',
    'Output must be valid JSON matching this schema:',
    '{',
    '  "headline": "Clear single executive takeaway sentence",',
    '  "observations": ["Observation 1", "Observation 2", "Observation 3"],',
    '  "next_steps": [{"action": "Actionable item description", "pic": "Owner/Team"}],',
    '  "speaker_notes": "A very detailed, professional presentation script for the speaker. Explain funnel drop-offs, user acquisition mix, operational risks in palm oil transactions, and step-by-step logic behind proposed recommendations in depth."',
    '}',
    '',
    'Rules (the slide has limited space, keep on-slide text short):',
    '- Ground every statement in DATA and quote numbers exactly as given.',
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
    headline: clampWords_(clean(parsed.headline) || 'Lapak Sawit Performance Executive Summary', 24),
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
    { text: 'Lapak Sawit', color: C.greenDark },
    { text: '   Feature Analysis Snapshot', color: C.muted, bold: false }
  ], box.x, box.y, 420, box.h, { fontSize: 12, bold: true, valign: 'MIDDLE' });

  const pillText = 'PERIOD  ' + formatPeriod_(data.dataPeriod.start, data.dataPeriod.end);
  const pillW = estimateTextWidth_(pillText, 6.5, true) + 18;
  pill_(slide, pillText, box.x + box.w - pillW, box.y + 2, pillW, 14, C.greenSoft, C.greenDark, 6.5);
}

function drawKpiCards_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const s = data.summary;
  const pickupAlert = s.totalPickupRequests === 0 && s.totalTransactions > 0;
  const cards = [
    { label: 'AVG DAILY ACTIVE', value: fmtNum_(s.avgDau), sub: 'users/day · ' + data.dailySeries.length + ' days', accent: C.blue },
    { label: 'HOME OPEN', value: fmtNum_(s.totalHomeOpen), sub: 'total events (funnel entry)', accent: C.greenDark },
    { label: 'JUAL TBS', value: fmtNum_(s.totalTransactions), sub: fmtPct_(s.transactionsShareOfHomePct) + ' of Home Open', accent: C.orange },
    {
      label: 'PICKUP REQUESTS',
      value: fmtNum_(s.totalPickupRequests),
      sub: pickupAlert ? '⚠ 0 of ' + fmtNum_(s.totalTransactions) + ' Jual TBS — check' : fmtPct_(s.pickupShareOfTransactionsPct) + ' of Jual TBS',
      accent: pickupAlert ? C.red : C.greenLine,
      alert: pickupAlert
    }
  ];
  const gap = 8;
  const w = (box.w - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => drawKpiCard_(slide, box.x + i * (w + gap), box.y, w, box.h, card));
}

function drawKpiCard_(slide, x, y, w, h, card) {
  const C = LS_CONFIG.COLORS;
  rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h, card.alert ? C.redSoft : C.cardBg, card.alert ? C.redBorder : C.border);
  rect_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + 7, 3, h - 14, card.accent);

  addText_(slide, card.label, x + 10, y + 5, w - 14, 8, { fontSize: 6, bold: true, color: C.muted });
  addText_(slide, card.value, x + 10, y + 14, w - 14, 16, { fontSize: 15, bold: true, color: card.alert ? C.red : C.ink, valign: 'MIDDLE' });
  addText_(slide, card.sub, x + 10, y + 31, w - 14, 7, { fontSize: 5.5, color: card.alert ? C.red : C.muted, bold: !!card.alert });
}

/**
 * Funnel as ranked rows: step badge, stage name, volume bar on a full-width track,
 * and a colour-coded step-conversion pill. Zero-volume stages and the largest leak are flagged red.
 */
function drawFunnel_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const stages = data.funnel;
  const leak = data.summary.largestLeak;

  panel_(slide, box);
  sectionTitle_(slide, 'FUNNEL CONVERSION', box.x + 10, box.y + 8, 140);

  const headerParts = [
    { text: 'Overall  ', color: C.muted },
    { text: fmtPct_(data.summary.overallConversionPct), color: C.ink, bold: true }
  ];
  if (leak) {
    headerParts.push(
      { text: '     Largest leak  ', color: C.muted },
      { text: leak.from + ' → ' + leak.to + ' (−' + fmtNum_(leak.lost) + ')', color: C.red, bold: true }
    );
  }
  addRichText_(slide, headerParts, box.x + box.w - 290, box.y + 8, 280, 11, { fontSize: 6, align: 'END', valign: 'MIDDLE' });

  const col = {
    badge: box.x + 10,
    label: box.x + 28,
    track: box.x + 122,
    conv: box.x + box.w - 98
  };
  col.labelW = col.track - col.label - 6;
  col.trackW = col.conv - col.track - 10;
  col.convW = box.x + box.w - 10 - col.conv;

  const headY = box.y + 24;
  [['STAGE', col.label, col.labelW, 'START'], ['EVENT VOLUME', col.track, col.trackW, 'START'], ['STEP CONVERSION', col.conv, col.convW, 'CENTER']]
    .forEach(h => addText_(slide, h[0], h[1], headY, h[2], 7, { fontSize: 5.5, bold: true, color: C.muted, align: h[3] }));
  hLine_(slide, box.x + 10, box.x + box.w - 10, headY + 9, C.border, 0.5);

  const rowY0 = box.y + 37, rowH = 18, rowGap = 4;
  const centerY = i => rowY0 + i * (rowH + rowGap) + rowH / 2;
  const maxVal = Math.max(1, ...stages.map(s => s.eventCount));

  // Vertical connector behind the step badges
  const connector = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, col.badge + 6, centerY(0), col.badge + 6, centerY(stages.length - 1));
  connector.getLineFill().setSolidFill(C.border);
  connector.setWeight(1);

  stages.forEach((stage, i) => {
    const y = centerY(i) - rowH / 2;
    const cy = centerY(i);
    const isZero = stage.eventCount === 0;
    const color = isZero ? C.red : LS_CONFIG.FUNNEL_COLORS[i % LS_CONFIG.FUNNEL_COLORS.length];

    // Step badge
    rect_(slide, SlidesApp.ShapeType.ELLIPSE, col.badge, cy - 6, 12, 12, color, C.white, 1);
    addText_(slide, String(i + 1), col.badge, cy - 6, 12, 12, { fontSize: 6, bold: true, color: C.white, align: 'CENTER', valign: 'MIDDLE' });

    // Stage name (shrinks to stay on one line where possible)
    const labelSize = fitFontSize_([stage.name], col.labelW, 9, 7, 6, 0, true);
    addText_(slide, stage.name, col.label, y, col.labelW, rowH, { fontSize: labelSize, bold: true, color: isZero ? C.red : C.ink, valign: 'MIDDLE' });

    // Volume bar on a full-width track
    rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 6, col.trackW, 12, C.track);
    const fillW = isZero ? 0 : Math.max(6, col.trackW * stage.eventCount / maxVal);
    if (fillW) rect_(slide, SlidesApp.ShapeType.ROUND_RECTANGLE, col.track, cy - 6, fillW, 12, color);

    const valueInside = fillW >= 34;
    if (valueInside) {
      addText_(slide, fmtNum_(stage.eventCount), col.track, y, fillW - 6, rowH, { fontSize: 7, bold: true, color: C.white, align: 'END', valign: 'MIDDLE' });
    }
    const outside = [];
    if (isZero) {
      outside.push({ text: '0', bold: true, color: C.red }, { text: '  ·  no events recorded', color: C.red });
    } else {
      if (!valueInside) outside.push({ text: fmtNum_(stage.eventCount), bold: true, color: C.ink });
      if (i > 0) outside.push({ text: (outside.length ? '  ·  ' : '') + fmtPct_(stage.shareOfTopPct) + ' of Home Open', color: C.muted });
    }
    if (outside.length) {
      const ox = col.track + fillW + 5;
      addRichText_(slide, outside, ox, y, col.track + col.trackW - ox, rowH, { fontSize: 6.5, valign: 'MIDDLE' });
    }

    // Step conversion pill
    const pill = conversionPill_(stage, i, leak);
    pill_(slide, pill.text, col.conv, cy - 6, col.convW, 12, pill.fill, pill.ink, 6);
  });
}

function conversionPill_(stage, i, leak) {
  const C = LS_CONFIG.COLORS;
  if (i === 0) return { text: 'Entry point', fill: C.cardBg, ink: C.muted };
  const pct = stage.stepConversionPct;
  if (pct == null) return { text: 'n/a (no prior)', fill: C.cardBg, ink: C.muted };
  const text = fmtPct_(pct) + ' conv  ▼' + fmtPct_(100 - pct, 0);
  if (pct === 0 || (leak && leak.index === i)) return { text: text, fill: C.redSoft, ink: C.red };
  if (pct < 30) return { text: text, fill: C.amberSoft, ink: C.amber };
  return { text: text, fill: C.okSoft, ink: C.okInk };
}

function drawDailyChart_(slide, data, box) {
  const C = LS_CONFIG.COLORS;
  const rows = data.dailySeries;
  const s = data.summary;

  panel_(slide, box);
  sectionTitle_(slide, 'DAILY ACTIVE USERS & USER MIX', box.x + 10, box.y + 8, 170);
  if (s.newUserSharePct != null) {
    addText_(slide, s.newUserSharePct + '% new users', box.x + 183, box.y + 8, 80, 11, { fontSize: 6, color: C.muted, valign: 'MIDDLE' });
  }

  const series = [
    { key: 'dau', label: 'DAU', color: C.blue },
    { key: 'newUsers', label: 'New', color: C.orange },
    { key: 'returningUsers', label: 'Returning', color: C.greenLine }
  ];

  // Legend, right-aligned on the title row
  let lx = box.x + box.w - 10;
  series.slice().reverse().forEach(item => {
    const tw = estimateTextWidth_(item.label, 6);
    lx -= tw;
    addText_(slide, item.label, lx, box.y + 8, tw, 11, { fontSize: 6, color: C.ink, valign: 'MIDDLE' });
    lx -= 13;
    const key = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, lx, box.y + 13.5, lx + 10, box.y + 13.5);
    key.getLineFill().setSolidFill(item.color);
    key.setWeight(2);
    lx -= 10;
  });

  if (!rows.length) return;

  const plot = { x: box.x + 34, y: box.y + 30, w: box.w - 34 - 32, h: box.h - 30 - 20 };
  const maxVal = niceMax_(Math.max(1, ...rows.flatMap(r => [r.dau, r.returningUsers, r.newUsers])));
  const px = i => rows.length === 1 ? plot.x + plot.w / 2 : plot.x + (i * plot.w) / (rows.length - 1);
  const py = v => plot.y + plot.h - (v / maxVal) * plot.h;

  // Gridlines + y-axis labels
  [0, 0.5, 1].forEach(f => {
    const gy = plot.y + plot.h - f * plot.h;
    const line = hLine_(slide, plot.x, plot.x + plot.w, gy, f === 0 ? C.border : C.grid, f === 0 ? 0.75 : 0.5);
    if (f > 0) line.setDashStyle(SlidesApp.DashStyle.DOT);
    addText_(slide, fmtNum_(maxVal * f), box.x + 6, gy - 4, plot.x - box.x - 12, 8, { fontSize: 5.5, color: C.muted, align: 'END', valign: 'MIDDLE' });
  });

  // X-axis date labels (thinned out for long ranges)
  const every = Math.ceil(rows.length / 8);
  rows.forEach((r, i) => {
    if (i % every !== 0 && i !== rows.length - 1) return;
    addText_(slide, formatShortDate_(r.date), px(i) - 20, plot.y + plot.h + 4, 40, 8, { fontSize: 5.5, color: C.muted, align: 'CENTER' });
  });

  // Series lines + point markers
  const showMarkers = rows.length <= 14;
  series.forEach(item => {
    for (let i = 1; i < rows.length; i++) {
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, px(i - 1), py(rows[i - 1][item.key]), px(i), py(rows[i][item.key]));
      line.getLineFill().setSolidFill(item.color);
      line.setWeight(1.75);
    }
    if (showMarkers) {
      rows.forEach((r, i) => rect_(slide, SlidesApp.ShapeType.ELLIPSE, px(i) - 2.25, py(r[item.key]) - 2.25, 4.5, 4.5, item.color, C.white, 0.75));
    }
  });

  // End-of-line value labels, nudged apart so they never overlap
  const lastIdx = rows.length - 1;
  const labels = series
    .map(item => ({ item: item, value: rows[lastIdx][item.key], y: py(rows[lastIdx][item.key]) }))
    .sort((a, b) => a.y - b.y);
  for (let i = 1; i < labels.length; i++) {
    if (labels[i].y - labels[i - 1].y < 8) labels[i].y = labels[i - 1].y + 8;
  }
  labels.forEach(l => {
    addText_(slide, fmtNum_(l.value), px(lastIdx) + 5, l.y - 4, 24, 8, { fontSize: 6, bold: true, color: l.item.color, valign: 'MIDDLE' });
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

function fmtNum_(n) {
  if (n == null || !isFinite(n)) return '–';
  const rounded = Math.abs(n) < 10 && n % 1 !== 0 ? n.toFixed(1) : String(Math.round(n));
  return rounded.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
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
