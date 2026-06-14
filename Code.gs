/**
 * IssueSheet — export Sentry into Google Sheets, on a schedule.
 *
 * Pulls a Sentry org's Issues, Releases, and daily error stats into this
 * spreadsheet via the documented Sentry REST API (/api/0/), which is
 * available on ALL Sentry plans, including the free Developer plan.
 *
 * Endpoints used (all ungated, verified against docs.sentry.io/api/):
 *   GET /api/0/projects/{org}/{project}/issues/        (scope: event:read)
 *   GET /api/0/organizations/{org}/releases/           (scope: project:releases)
 *   GET /api/0/organizations/{org}/stats_v2/           (scope: org:read)
 *
 * Deliberately NOT used: /api/0/organizations/{org}/events/ (Discover) —
 * possibly plan-gated. Everything here works on a free org.
 *
 * Setup: see README.md. Set Script Properties SENTRY_ORG, SENTRY_PROJECT,
 * SENTRY_TOKEN, then reload the sheet and use the "IssueSheet" menu.
 *
 * IssueSheet is an independent tool, not affiliated with or endorsed by
 * Sentry / Functional Software, Inc.
 */

var SENTRY_API_BASE = 'https://sentry.io/api/0/';
var MAX_ROWS = 200; // cap on rows pulled per tab — keeps a sync within Apps Script quotas

// ---------------------------------------------------------------- menu --

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('IssueSheet')
    .addItem('Sync now', 'syncNow')
    .addSeparator()
    .addItem('Clear all tabs', 'resetTabs')
    .addItem('Check settings', 'checkSettings')
    .addToUi();
}

// ---------------------------------------------------------------- sync --

function syncNow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var cfg = getConfig_();

    ss.toast('Pulling issues from Sentry…', 'IssueSheet', 10);
    var nIssues = syncIssues_(cfg, ss);

    ss.toast('Pulling releases…', 'IssueSheet', 10);
    var nReleases = syncReleases_(cfg, ss);

    ss.toast('Pulling 30 days of error stats…', 'IssueSheet', 10);
    var nDays = syncStats_(cfg, ss);

    ss.toast('Building dashboard…', 'IssueSheet', 10);
    buildDashboard_(ss, cfg);

    ss.toast(
      'Done: ' + nIssues + ' issues, ' + nReleases + ' releases, ' +
      nDays + ' days of stats. Dashboard updated.',
      'IssueSheet — synced', 8
    );
  } catch (e) {
    Logger.log('Sync failed: ' + (e && e.stack ? e.stack : e));
    ss.toast(String((e && e.message) || e), 'IssueSheet — sync failed', 10);
  }
}

function syncIssues_(cfg, ss) {
  var issues = fetchAllPages_(
    cfg,
    'projects/' + cfg.org + '/' + cfg.project + '/issues/',
    { query: 'is:unresolved', limit: 100 },
    MAX_ROWS
  );

  var rows = issues.map(function (it) {
    return [
      it.title || '',
      it.level || '',
      Number(it.count || 0),
      Number(it.userCount || 0),
      it.firstSeen ? new Date(it.firstSeen) : '',
      it.lastSeen ? new Date(it.lastSeen) : '',
      it.status || '',
      it.substatus || '', // present on current Sentry; blank-safe if absent
      it.permalink
        ? '=HYPERLINK("' + it.permalink + '","Open in Sentry")'
        : ''
    ];
  });

  var sheet = writeTab_(
    ss, 'Issues',
    ['Title', 'Level', 'Events', 'Users', 'First seen', 'Last seen',
     'Status', 'Substatus', 'Link'],
    rows,
    { 5: 'yyyy-mm-dd hh:mm', 6: 'yyyy-mm-dd hh:mm' }
  );
  sheet.setColumnWidth(1, 340); // keep long titles readable, not enormous
  if (rows.length) {
    var levelColors = rows.map(function (r) { return [levelColor_(r[1])]; });
    sheet.getRange(2, 2, rows.length, 1).setBackgrounds(levelColors);
  }
  return rows.length;
}

function syncReleases_(cfg, ss) {
  var releases = fetchAllPages_(
    cfg,
    'organizations/' + cfg.org + '/releases/',
    {},
    MAX_ROWS
  );

  var rows = releases.map(function (r) {
    var projects = (r.projects || []).map(function (p) { return p.slug; }).join(', ');
    return [
      r.version || '',
      r.dateCreated ? new Date(r.dateCreated) : '',
      Number(r.commitCount || 0),
      Number(r.newGroups || 0),
      Number(r.deployCount || 0),
      projects
    ];
  });

  writeTab_(
    ss, 'Releases',
    ['Version', 'Created', 'Commits', 'New issues', 'Deploys', 'Projects'],
    rows,
    { 2: 'yyyy-mm-dd hh:mm' }
  );
  return rows.length;
}

function syncStats_(cfg, ss) {
  // stats_v2 requires `field` and `groupBy` (verified in docs).
  // 30d window matches free/Team-plan retention; 1d buckets for a daily chart.
  var res = fetchJson_(cfg, 'organizations/' + cfg.org + '/stats_v2/', {
    field: 'sum(quantity)',
    statsPeriod: '30d',
    interval: '1d',
    groupBy: 'outcome',
    category: 'error'
  });

  var intervals = res.data.intervals || [];
  var groups = res.data.groups || [];
  var outcomes = groups.map(function (g) { return (g.by && g.by.outcome) || '?'; });

  var headers = ['Date'].concat(outcomes).concat(['Total errors']);
  var rows = intervals.map(function (ts, i) {
    var row = [new Date(ts)];
    var total = 0;
    groups.forEach(function (g) {
      var series = (g.series && g.series['sum(quantity)']) || [];
      var v = Number(series[i] || 0);
      total += v;
      row.push(v);
    });
    row.push(total);
    return row;
  });

  writeTab_(ss, 'Stats', headers, rows, { 1: 'yyyy-mm-dd' });
  return rows.length;
}

// ----------------------------------------------------------- API layer --

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  var cfg = {
    org: props.getProperty('SENTRY_ORG'),
    project: props.getProperty('SENTRY_PROJECT'),
    token: props.getProperty('SENTRY_TOKEN')
  };
  var missing = [];
  if (!cfg.org) missing.push('SENTRY_ORG');
  if (!cfg.project) missing.push('SENTRY_PROJECT');
  if (!cfg.token) missing.push('SENTRY_TOKEN');
  if (missing.length) {
    throw new Error(
      'Missing Script Properties: ' + missing.join(', ') +
      '. Set them under Project Settings > Script Properties.'
    );
  }
  return cfg;
}

/**
 * One GET against the Sentry API. Returns { data, next } where `next` is the
 * cursor for the next page (parsed from the Link header) or null.
 */
function fetchJson_(cfg, path, params) {
  var qs = [];
  for (var k in params) {
    var v = params[k];
    if (v === null || v === undefined || v === '') continue;
    if (Object.prototype.toString.call(v) === '[object Array]') {
      for (var i = 0; i < v.length; i++) {
        qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(v[i]));
      }
    } else {
      qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
    }
  }
  var url = SENTRY_API_BASE + path + (qs.length ? '?' + qs.join('&') : '');

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code === 401) {
    throw new Error('Sentry rejected the token (401). Re-check SENTRY_TOKEN.');
  }
  if (code === 403) {
    throw new Error(
      'Token is missing a scope (403). It needs: event:read, ' +
      'project:releases, org:read.'
    );
  }
  if (code === 404) {
    throw new Error(
      'Not found (404). Check the SENTRY_ORG / SENTRY_PROJECT slugs ' +
      '(they are the lowercase slugs from your Sentry URLs, not display names).'
    );
  }
  if (code === 429) {
    throw new Error('Sentry rate limit hit (429). Wait a minute and retry.');
  }
  if (code >= 400) {
    throw new Error(
      'Sentry API error ' + code + ' on ' + path + ': ' +
      String(resp.getContentText()).slice(0, 200)
    );
  }

  var headers = resp.getAllHeaders();
  var link = headers['Link'] || headers['link'] || '';
  return {
    data: JSON.parse(resp.getContentText()),
    next: nextCursor_(String(link))
  };
}

/**
 * Follows Link-header cursor pagination until `cap` rows are collected or
 * the API says there are no more results.
 *
 * Sentry Link header format (verified in docs):
 *   <url>; rel="next"; results="true"; cursor="0:100:0", <url>; rel="previous"; ...
 * Cursors are ALWAYS returned; results="false" means the page is empty,
 * so only follow when results="true".
 */
function fetchAllPages_(cfg, path, params, cap) {
  var rows = [];
  var cursor = null;
  do {
    var p = {};
    for (var k in params) p[k] = params[k];
    if (cursor) p.cursor = cursor;
    var page = fetchJson_(cfg, path, p);
    rows = rows.concat(page.data);
    cursor = page.next;
  } while (cursor && rows.length < cap);
  return rows.slice(0, cap);
}

function nextCursor_(linkHeader) {
  if (!linkHeader) return null;
  var parts = linkHeader.split(',');
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];
    if (part.indexOf('rel="next"') === -1) continue;
    if (part.indexOf('results="true"') === -1) return null; // no more data
    var m = part.match(/cursor="([^"]+)"/);
    return m ? m[1] : null;
  }
  return null;
}

// -------------------------------------------------------- sheet writing --

/**
 * (Re)writes a tab: header row (bold, dark, frozen), data rows, optional
 * per-column number formats ({1-based col index: format}), a "Last synced"
 * stamp to the right of the headers, and auto-sized columns.
 */
function writeTab_(ss, name, headers, rows, colFormats) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#2b2b40')
    .setFontColor('#ffffff');
  if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    if (colFormats) {
      for (var col in colFormats) {
        sheet.getRange(2, Number(col), rows.length, 1)
          .setNumberFormat(colFormats[col]);
      }
    }
  }

  var stamp = 'Last synced: ' + Utilities.formatDate(
    new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(1, headers.length + 2)
    .setValue(stamp)
    .setFontStyle('italic')
    .setFontWeight('normal')
    .setFontColor('#888888')
    .setBackground(null);

  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

// ----------------------------------------------------------- dashboard --

var LEVEL_COLORS = {
  fatal: '#f4c7c3',
  error: '#fdecea',
  warning: '#fff2cc',
  info: '#e8f0fe'
};

function levelColor_(level) {
  return LEVEL_COLORS[String(level || '').toLowerCase()] || '#ffffff';
}

/**
 * Builds the Dashboard tab from the already-synced Issues/Releases/Stats
 * tabs: KPI row, top-issues table (severity-colored), and an errors-per-day
 * column chart. Pure presentation — no extra API calls.
 */
function buildDashboard_(ss, cfg) {
  var sheet = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard', 0);
  sheet.clear();
  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  sheet.setHiddenGridlines(true);

  var issues = tabData_(ss, 'Issues');
  var releases = tabData_(ss, 'Releases');
  var stats = tabData_(ss, 'Stats');

  var totalIdx = stats.headers.indexOf('Total errors');
  var sum30 = 0, sum7 = 0;
  stats.rows.forEach(function (r, i) {
    var v = totalIdx >= 0 ? Number(r[totalIdx] || 0) : 0;
    sum30 += v;
    if (i >= stats.rows.length - 7) sum7 += v;
  });

  sheet.getRange(1, 2).setValue('Sentry overview')
    .setFontSize(18).setFontWeight('bold');
  sheet.getRange(2, 2).setValue(
    'project: ' + cfg.project + '  ·  last synced ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd HH:mm'))
    .setFontColor('#888888').setFontStyle('italic');

  var kpis = [
    [issues.rows.length, 'Open issues'],
    [sum30, 'Errors (30d)'],
    [sum7, 'Errors (7d)'],
    [releases.rows.length, 'Releases tracked']
  ];
  kpis.forEach(function (kpi, i) {
    var col = 2 + i * 2;
    sheet.getRange(4, col).setValue(kpi[0])
      .setFontSize(22).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#f1f3f4');
    sheet.getRange(5, col).setValue(kpi[1])
      .setFontSize(9).setFontColor('#666666')
      .setHorizontalAlignment('center').setBackground('#f1f3f4');
  });

  var top = issues.rows.slice().sort(function (a, b) {
    return Number(b[2] || 0) - Number(a[2] || 0);
  }).slice(0, 5);

  sheet.getRange(7, 2).setValue('Top issues by events')
    .setFontWeight('bold').setFontSize(11);
  var headers = ['Title', 'Level', 'Events', 'Users'];
  sheet.getRange(8, 2, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2b2b40').setFontColor('#ffffff');
  if (top.length) {
    var rows = top.map(function (r) { return [r[0], r[1], r[2], r[3]]; });
    sheet.getRange(9, 2, rows.length, headers.length).setValues(rows);
    var colors = top.map(function (r) { return [levelColor_(r[1])]; });
    sheet.getRange(9, 3, rows.length, 1).setBackgrounds(colors);
  }

  sheet.setColumnWidth(1, 24);
  sheet.setColumnWidth(2, 320);

  var statsSheet = ss.getSheetByName('Stats');
  if (statsSheet && stats.rows.length && totalIdx >= 0) {
    var n = stats.rows.length + 1; // + header row
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(statsSheet.getRange(1, 1, n, 1))
      .addRange(statsSheet.getRange(1, totalIdx + 1, n, 1))
      .setNumHeaders(1)
      .setPosition(7, 7, 0, 0)
      .setOption('title', 'Errors per day — last 30 days')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#e8590c'])
      .setOption('width', 620)
      .setOption('height', 300)
      .build();
    sheet.insertChart(chart);
  }
}

/**
 * Header + data rows of a synced tab. Header width stops before the empty
 * spacer column and the "Last synced" stamp that writeTab_ appends.
 */
function tabData_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return { headers: [], rows: [] };
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var width = 0;
  while (width < headerRow.length && String(headerRow[width]) !== '' &&
         String(headerRow[width]).indexOf('Last synced') !== 0) width++;
  if (!width) return { headers: [], rows: [] };
  return {
    headers: headerRow.slice(0, width),
    rows: sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues()
  };
}

// ------------------------------------------------------------- helpers --

/** Clears all data tabs and dashboard charts (e.g. to start fresh). */
function resetTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Dashboard', 'Issues', 'Releases', 'Stats'].forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.clear();
    sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  });
  ss.toast('All tabs cleared.', 'IssueSheet', 5);
}

/** Shows which Script Properties are set (token masked). */
function checkSettings() {
  var props = PropertiesService.getScriptProperties();
  var org = props.getProperty('SENTRY_ORG') || '(not set)';
  var project = props.getProperty('SENTRY_PROJECT') || '(not set)';
  var token = props.getProperty('SENTRY_TOKEN');
  var tokenMsg = token ? 'set (' + token.slice(0, 8) + '…)' : '(not set)';
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'IssueSheet settings',
    'SENTRY_ORG: ' + org + '\n' +
    'SENTRY_PROJECT: ' + project + '\n' +
    'SENTRY_TOKEN: ' + tokenMsg,
    ui.ButtonSet.OK
  );
}
