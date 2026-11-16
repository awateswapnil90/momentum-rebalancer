/**
 * Momentum Rebalancer – Kite sync (Google Apps Script, bound to the Google Sheet)
 *
 * What it does: after you log in to Kite, it reads today's orders, keeps only the
 * filled CNC orders that came from the GitHub page (Kite tag "kc:pub" or "MOMRB"),
 * and updates Holdings, Trade Log (with realized P&L), cash (Settings B5),
 * Settings B6 and History.
 *
 * Setup: paste this file into Extensions > Apps Script, save, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Reload the sheet and use the Kite menu.
 *
 * Your API secret is stored in this script's private properties (your Google
 * account only). It is never written into the sheet.
 */

var API = 'https://api.kite.trade';
var STRATEGY_TAGS = ['kc:pub', 'MOMRB'];
var FINAL = ['COMPLETE', 'CANCELLED', 'REJECTED'];
var HOLD_ROWS = 100;
var TZ = 'Asia/Kolkata';

function props_() { return PropertiesService.getScriptProperties(); }

// ---------------- Sheet menu ----------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kite')
    .addItem('1. Save API key & secret', 'setupKeys')
    .addItem('2. Show links (redirect + sync)', 'showLinks')
    .addSeparator()
    .addItem('Sync fills now', 'showLinks')
    .addSeparator()
    .addItem('Run demo (on a copy of this sheet)', 'runDemo')
    .addItem('Use latest prices for quantities (one-time upgrade)', 'upgradeLivePrices')
    .addItem('Fix old dates to India time (run once)', 'fixDatesToIST')
    .addItem('Update ledger & monthly summary now', 'updateLedgerNow')
    .addItem('Organize tabs & refresh README', 'organizeTabsNow')
    .addToUi();
}

function setupKeys() {
  var ui = SpreadsheetApp.getUi();
  var k = ui.prompt('Kite Connect PERSONAL app', 'API key:', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  var s = ui.prompt('Kite Connect PERSONAL app', 'API secret (kept private in this script):', ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  props_().setProperties({
    KITE_API_KEY: k.getResponseText().trim(),
    KITE_API_SECRET: s.getResponseText().trim(),
    SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId()
  });
  ui.alert('Saved. Next: Kite > Show links.');
}

function showLinks() {
  var p = props_();
  if (!p.getProperty('SHEET_ID')) p.setProperty('SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  var url = ScriptApp.getService().getUrl();
  var key = p.getProperty('KITE_API_KEY');
  var html;
  if (!url) {
    html = '<p>Deploy the script first: <b>Deploy &gt; New deployment &gt; Web app</b> (Execute as: Me, Access: Anyone).</p>';
  } else {
    html = '<p><b>Redirect URL</b> (paste into your Personal app in the Kite developer console):</p>' +
      '<textarea style="width:100%;height:60px" readonly onclick="this.select()">' + url + '</textarea>' +
      '<p><b>Sync link</b> (paste into the GitHub page Setup, or bookmark on your phone):</p>' +
      '<textarea style="width:100%;height:60px" readonly onclick="this.select()">' + url + '</textarea>' +
      (key ? '<p><a href="' + url + '" target="_blank">Open sync now</a></p>' : '<p style="color:#c00">Save your API key & secret first (Kite menu, step 1).</p>');
  }
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<div style="font:14px Arial">' + html + '</div>').setWidth(520).setHeight(330), 'Kite links');
}

// ---------------- Web app (Kite redirects here after login) ----------------
function doGet(e) {
  var p = props_(), key = p.getProperty('KITE_API_KEY'), sec = p.getProperty('KITE_API_SECRET');
  if (!key || !sec) return page_('Not set up', '<p>Open the sheet &gt; Kite &gt; Save API key &amp; secret.</p>');
  var q = (e && e.parameter) || {};
  if (!q.request_token) {
    var login = 'https://kite.zerodha.com/connect/login?v=3&api_key=' + encodeURIComponent(key);
    return page_('Sync fills to sheet',
      '<p>Log in to Kite. You will come back here and the sheet updates with today\'s filled orders from the rebalancer page.</p>' +
      '<p><a class="btn" target="_top" href="' + login + '">Log in to Kite &amp; sync</a></p>');
  }
  if (q.status && q.status !== 'success') return page_('Kite login failed', '<p>' + esc_(q.status) + '</p>');

  // A Kite request_token works only once. Browsers sometimes load this page twice,
  // so remember the result for each token and show it again instead of re-using the token.
  var cache = CacheService.getScriptCache(), ck = 'rt_' + sha256Hex_(q.request_token).slice(0, 40);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var seen = cache.get(ck);
    if (seen) return page_('Sync done', seen);
    var body;
    try {
      body = syncWithToken_(key, sec, q.request_token);
    } catch (err) {
      return page_('Sync failed', '<p>' + esc_(err.message) + '</p>' +
        '<p>App key in use: <b>' + esc_(key.slice(0, 4)) + '…' + esc_(key.slice(-3)) + '</b>. It must be your Kite Connect <b>Personal</b> app ' +
        '(not the Publisher app), with that same app\'s secret. If not, re-save them via Kite &gt; Save API key &amp; secret.</p>' +
        '<p><a class="btn" target="_top" href="' + ScriptApp.getService().getUrl() + '">Try again</a></p>');
    }
    cache.put(ck, body, 21600);
    return page_('Sync done', body);
  } finally {
    lock.releaseLock();
  }
}

function page_(title, body) {
  var css = '<style>body{font:15px/1.45 system-ui,Arial,sans-serif;max-width:640px;margin:20px auto;padding:0 14px;color:#1b1f24}' +
    'table{width:100%;border-collapse:collapse;font-size:13px}td,th{padding:5px 4px;border-bottom:1px solid #e2e5e9;text-align:right}' +
    'td:first-child,th:first-child,td:nth-child(2),th:nth-child(2){text-align:left}.btn{display:inline-block;background:#387ed1;color:#fff;' +
    'padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:600}.BUY{color:#0a7d3b}.SELL{color:#c0262d}</style>';
  return HtmlService.createHtmlOutput(css + '<h2>' + esc_(title) + '</h2>' + body)
    .setTitle('Momentum Rebalancer – ' + title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function esc_(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function inr_(n) { return '₹' + Math.round(n).toLocaleString('en-IN'); }

// ---------------- Kite API ----------------
function kite_(method, path, auth, payload) {
  var opt = { method: method, muteHttpExceptions: true, headers: { 'X-Kite-Version': '3' } };
  if (auth) opt.headers.Authorization = auth;
  if (payload) opt.payload = payload;
  var res = UrlFetchApp.fetch(API + path, opt), j;
  try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('Kite ' + path + ': HTTP ' + res.getResponseCode()); }
  if (j.status !== 'success') throw new Error('Kite ' + path + ': ' + (j.message || res.getResponseCode()));
  return j.data;
}

function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function syncWithToken_(key, sec, token) {
  var sess = kite_('post', '/session/token', null,
    { api_key: key, request_token: token, checksum: sha256Hex_(key + token + sec) });
  var auth = 'token ' + key + ':' + sess.access_token;
  var orders = kite_('get', '/orders', auth);
  var holdings = [];
  try { holdings = kite_('get', '/portfolio/holdings', auth); } catch (e) { /* prices are optional */ }
  return applyOrders_(orders, holdings);
}

// ---------------- Sheet update ----------------
function parseTime_(s) {
  if (!s) return new Date();
  s = String(s);
  if (s.indexOf('T') > -1) return new Date(s);
  return Utilities.parseDate(s.slice(0, 19), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function isStrategy_(o) {
  var tags = (o.tags || []).slice();
  if (o.tag) tags.push(o.tag);
  return tags.some(function (t) { return STRATEGY_TAGS.indexOf(t) > -1; });
}

function applyOrders_(orders, kiteHoldings, ssOverride) {
  {
    var ss = ssOverride || SpreadsheetApp.openById(props_().getProperty('SHEET_ID'));
    if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ); // show all dates in IST
    var log = ss.getSheetByName('Trade Log'), hold = ss.getSheetByName('Holdings'),
        set = ss.getSheetByName('Settings'), hist = ss.getSheetByName('History');

    // already-logged order ids
    var logged = {};
    if (log.getLastRow() > 1) log.getRange(2, 2, log.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0] !== '') logged[String(r[0])] = 1; });

    // current strategy holdings
    var map = {};
    hold.getRange(2, 1, HOLD_ROWS, 5).getValues().forEach(function (r) {
      if (String(r[0]).trim() && Number(r[1]) > 0)
        map[String(r[0]).trim()] = { qty: Number(r[1]), avg: Number(r[2]) || 0, date: r[3], notes: r[4] };
    });

    var ltp = {};
    (kiteHoldings || []).forEach(function (h) { ltp[h.tradingsymbol] = h.last_price; });

    var todo = (orders || []).filter(function (o) {
      return o.filled_quantity > 0 && FINAL.indexOf(o.status) > -1 && o.product === 'CNC' &&
        isStrategy_(o) && !logged[String(o.order_id)];
    }).sort(function (a, b) { return String(a.order_timestamp).localeCompare(String(b.order_timestamp)); });

    var pending = (orders || []).filter(function (o) {
      return isStrategy_(o) && FINAL.indexOf(o.status) < 0;
    }).length;

    var rows = [], cashDelta = 0, realized = 0, nB = 0, nS = 0, warn = [];
    todo.forEach(function (o) {
      var sym = o.tradingsymbol, q = Number(o.filled_quantity), px = Number(o.average_price), val = q * px;
      var dt = parseTime_(o.exchange_timestamp || o.order_timestamp);
      if (o.transaction_type === 'BUY') {
        var h = map[sym] || { qty: 0, avg: 0, date: dt, notes: '' };
        h.avg = (h.qty * h.avg + val) / (h.qty + q);
        h.qty += q;
        map[sym] = h;
        cashDelta -= val; nB++;
        rows.push([dt, String(o.order_id), sym, 'BUY', q, px, val, '', '', o.status]);
      } else {
        var s = map[sym], avgBuy = '', pnl = '';
        if (s) {
          avgBuy = s.avg; pnl = (px - s.avg) * q; realized += pnl;
          s.qty -= q;
          if (s.qty <= 0) delete map[sym];
        } else {
          warn.push(sym + ' was sold but is not in Holdings, so its P&L is blank.');
        }
        cashDelta += val; nS++;
        rows.push([dt, String(o.order_id), sym, 'SELL', q, px, val, avgBuy, pnl, o.status]);
      }
    });

    // write Trade Log
    if (rows.length) log.getRange(log.getLastRow() + 1, 1, rows.length, 10).setValues(rows);

    // rewrite Holdings A:G (H:J keep their formulas)
    var syms = Object.keys(map).sort();
    if (syms.length > HOLD_ROWS) throw new Error('More than ' + HOLD_ROWS + ' holdings.');
    hold.getRange(2, 1, HOLD_ROWS, 7).clearContent();
    if (syms.length) {
      hold.getRange(2, 1, syms.length, 6).setValues(syms.map(function (k) {
        var h = map[k];
        return [k, h.qty, h.avg, h.date || '', h.notes || '', ltp[k] != null ? ltp[k] : ''];
      }));
      hold.getRange(2, 7, syms.length, 1).setFormulas(syms.map(function (k, i) {
        return ['=IFERROR(GOOGLEFINANCE("NSE:"&A' + (i + 2) + '),"")'];
      }));
    }

    // Manual prices on Rebalance (col E) are tied to Holdings rows, which just changed
    if (rows.length) { var rb = ss.getSheetByName('Rebalance'); if (rb) rb.getRange(2, 5, HOLD_ROWS, 1).clearContent(); }

    // Settings: cash, month-1 flag, last sync
    var cash = Number(set.getRange('B5').getValue()) || 0;
    cash += cashDelta;
    set.getRange('B5').setValue(cash);
    if (nB > 0 && String(set.getRange('B6').getValue()).trim().toUpperCase() !== 'YES') set.getRange('B6').setValue('YES');
    var now = new Date();
    set.getRange('B14').setValue(now);

    // History snapshot
    SpreadsheetApp.flush();
    var cost = 0, value = 0;
    syms.forEach(function (k) {
      var h = map[k], p = ltp[k] != null ? ltp[k] : h.avg;
      cost += h.qty * h.avg; value += h.qty * p;
    });
    var realizedCum = 0;
    if (log.getLastRow() > 1) log.getRange(2, 9, log.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { realizedCum += Number(r[0]) || 0; });
    var capital = Number(set.getRange('B2').getValue()) || 0;
    hist.appendRow([now, cost, value, cash, value + cash, realizedCum, value + cash - capital]);
    ensureLedger_(ss);

    // result page
    var t = '<p>' + (rows.length ? 'Recorded <b>' + nB + ' buys</b> and <b>' + nS + ' sells</b>.' : 'No new filled orders from the rebalancer today.') + '</p>';
    if (rows.length) {
      t += '<table><tr><th>Symbol</th><th>Side</th><th>Qty</th><th>Avg</th><th>P&amp;L</th></tr>' + rows.map(function (r) {
        return '<tr><td>' + esc_(r[2]) + '</td><td class="' + r[3] + '">' + r[3] + '</td><td>' + r[4] + '</td><td>' + r[5].toFixed(2) +
          '</td><td>' + (r[8] === '' ? '' : inr_(r[8])) + '</td></tr>';
      }).join('') + '</table>';
    }
    if (nS) t += '<p>Realized P&amp;L this sync: <b>' + inr_(realized) + '</b> (before charges)</p>';
    t += '<p>Holdings: <b>' + syms.length + '</b> stocks · Cash: <b>' + inr_(cash) + '</b> · Portfolio ≈ <b>' + inr_(value + cash) + '</b></p>';
    if (pending) t += '<p style="color:#c0262d">' + pending + ' rebalancer order(s) are still open in Kite. Sync again after they complete.</p>';
    warn.forEach(function (w) { t += '<p style="color:#c0262d">' + esc_(w) + '</p>'; });
    t += '<p><a href="https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit" target="_top">Open the sheet</a></p>';
    return t;
  }
}


// ---------------- Demo (runs on a COPY, never on this sheet) ----------------
function runDemo() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Run demo', 'This makes a COPY of this sheet and fills it with made-up trades ' +
    '(a month-1 buy, then a month-2 rebalance). Your real sheet is not changed. Continue?', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var src = SpreadsheetApp.getActiveSpreadsheet();
  var copy = src.copy('Momentum Rebalancer – DEMO (safe to delete)');
  SpreadsheetApp.flush();

  var fmt = function (d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm:ss'); };
  var id = 900000, day1 = new Date(), day2 = new Date(day1.getTime() + 21 * 86400000);
  var mk = function (sym, side, q, px, d) {
    id++;
    return { order_id: 'DEMO' + id, tradingsymbol: sym, transaction_type: side, filled_quantity: q,
      average_price: Math.round(px * 100) / 100, status: 'COMPLETE', tags: ['kc:pub'], product: 'CNC',
      order_timestamp: fmt(d), exchange_timestamp: fmt(d) };
  };

  // Month 1: the Initial tab's plan, filled 0.2% above the scanner price
  var plan = copy.getSheetByName('Initial').getRange(2, 2, 20, 7).getValues()
    .filter(function (r) { return r[0] && r[6] > 0; });
  if (!plan.length) { ui.alert('The Initial tab has no orders, so there is nothing to demo.'); return; }
  var price = {};
  plan.forEach(function (r) { price[r[0]] = Number(r[1]); });
  var o1 = plan.map(function (r) { return mk(r[0], 'BUY', Number(r[6]), Number(r[1]) * 1.002, day1); });

  // Month 2: sell the first two (one up 10%, one down 8%), buy rank 21 and 22 with the proceeds
  var s1 = plan[0], s2 = plan[1 % plan.length];
  var sells = [mk(s1[0], 'SELL', Number(s1[6]), Number(s1[1]) * 1.10, day2)];
  if (plan.length > 1) sells.push(mk(s2[0], 'SELL', Number(s2[6]), Number(s2[1]) * 0.92, day2));
  var proceeds = sells.reduce(function (a, o) { return a + o.filled_quantity * o.average_price; }, 0);
  var ranked = copy.getSheetByName('Ranked').getRange(22, 2, 10, 2).getValues()
    .filter(function (r) { return r[0] && !price[r[0]]; }).slice(0, sells.length);
  var buys = ranked.map(function (r) {
    price[r[0]] = Number(r[1]);
    return mk(r[0], 'BUY', Math.floor(proceeds / sells.length / Number(r[1])), Number(r[1]) * 1.003, day2);
  }).filter(function (o) { return o.filled_quantity > 0; });

  // Pretend prices have moved +6% since
  var ltp = Object.keys(price).map(function (k) { return { tradingsymbol: k, last_price: Math.round(price[k] * 1.06 * 100) / 100 }; });

  applyOrders_(o1, ltp, copy);
  applyOrders_(sells.concat(buys), ltp, copy);

  var url = copy.getUrl();
  ui.showModalDialog(HtmlService.createHtmlOutput(
    '<div style="font:14px Arial"><p>Demo copy created with ' + o1.length + ' month-1 buys, ' + sells.length + ' sells and ' + buys.length +
    ' replacement buys (all made up).</p><p>Look at its <b>Dashboard</b>, <b>Holdings</b>, <b>Trade Log</b> and <b>History</b> tabs.</p>' +
    '<p><a href="' + url + '" target="_blank">Open the demo copy</a></p><p>Delete the copy when you are done.</p></div>')
    .setWidth(420).setHeight(220), 'Demo ready');
}


// ---------------- Upgrade: size orders on the latest price, not the scanner LCP ----------------
function upgradeLivePrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var set = ss.getSheetByName('Settings'), calc = ss.getSheetByName('Calc'),
      rk = ss.getSheetByName('Ranked'), rb = ss.getSheetByName('Rebalance');
  if (!set || !calc || !rk || !rb) { ui.alert('This sheet is missing Settings/Calc/Ranked/Rebalance tabs.'); return; }
  var N = 300, H = 100, i, f;
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);

  set.getRange('A15:C15').setValues([['Price for quantities (LIVE/LCP)', 'LIVE',
    'LIVE = latest NSE price from Google Finance (about 15-20 min delayed; last close when the market is shut). LCP = scanner price']]);

  calc.getRange('E1:F1').setValues([['Price used', 'Source']]);
  f = [];
  for (i = 2; i <= N + 1; i++) {
    var g = 'GOOGLEFINANCE("NSE:"&A' + i + ')';
    f.push(['=IF(A' + i + '="","",IF(UPPER(TRIM(Settings!$B$15))="LIVE",IFERROR(IF(' + g + '>0,' + g + ',B' + i + '),B' + i + '),B' + i + '))',
            '=IF(A' + i + '="","",IF(E' + i + '=B' + i + ',"LCP","LIVE"))']);
  }
  calc.getRange(2, 5, N, 2).setFormulas(f);

  rk.getRange('C1').setValue('Price');
  f = [];
  for (i = 2; i <= N + 1; i++) f.push(['=IFERROR(INDEX(Calc!$E$2:$E$' + (N + 1) + ',MATCH(A' + i + ',Calc!$D$2:$D$' + (N + 1) + ',0)),"")']);
  rk.getRange(2, 3, N, 1).setFormulas(f);

  rb.getRange('D1').setValue('New price');
  f = [];
  for (i = 2; i <= H + 1; i++) f.push(['=IF(A' + i + '="","",IFERROR(INDEX(Calc!$E$2:$E$' + (N + 1) + ',MATCH(A' + i + ',Calc!$A$2:$A$' + (N + 1) + ',0)),""))']);
  rb.getRange(2, 4, H, 1).setFormulas(f);

  var ini = ss.getSheetByName('Initial'); if (ini) ini.getRange('C1').setValue('Price');
  var bu = ss.getSheetByName('Buys'); if (bu) bu.getRange('D1').setValue('Price');

  SpreadsheetApp.flush();
  ui.alert('Done. Quantities now use the latest NSE price (Settings B15 = LIVE). ' +
    'Calc tab, column F shows LIVE or LCP per stock; LCP means Google Finance had no price for that symbol, so the scanner price is used. ' +
    'Set B15 to LCP to go back.');
}


// ---------------- One-off: convert dates written while the sheet was on US Pacific time ----------------
function fixDatesToIST() {
  var OLD = 'America/Los_Angeles', ui = SpreadsheetApp.getUi(), p = props_();
  if (p.getProperty('DATES_FIXED_IST')) { ui.alert('Already done earlier. Nothing changed.'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);
  var F = 'yyyy-MM-dd HH:mm:ss', n = 0;
  var conv = function (v) {
    if (!(v instanceof Date)) return v;
    n++;
    return Utilities.parseDate(Utilities.formatDate(v, TZ, F), OLD, F);
  };
  var fixCol = function (sheetName, col) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var rg = sh.getRange(2, col, sh.getLastRow() - 1, 1);
    rg.setValues(rg.getValues().map(function (r) { return [conv(r[0])]; }));
  };
  fixCol('Trade Log', 1);
  fixCol('Holdings', 4);
  fixCol('History', 1);
  var b14 = ss.getSheetByName('Settings').getRange('B14');
  b14.setValue(conv(b14.getValue()));
  ss.getSheetByName('Trade Log').getRange('A2:A').setNumberFormat('dd-mmm-yyyy hh:mm');
  ss.getSheetByName('History').getRange('A2:A').setNumberFormat('dd-mmm-yyyy hh:mm');
  p.setProperty('DATES_FIXED_IST', new Date().toISOString());
  ui.alert('Converted ' + n + ' dates to India time (Trade Log, Holdings, History, Settings B14).');
}


// ---------------- Ledger: running cash / running P&L on Trade Log + Monthly Summary tab ----------------
function updateLedgerNow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureLedger_(ss);
  organizeTabs_(ss);
  SpreadsheetApp.getUi().alert('Ledger updated (Trade Log columns K-M + Monthly Summary). Tabs re-ordered and colour-coded; helper tabs hidden. See README for the tab guide.');
}

function organizeTabsNow() {
  organizeTabs_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Tabs re-ordered and colour-coded; helper tabs hidden. README has the full tab guide.');
}

function ensureLedger_(ss) {
  var log = ss.getSheetByName('Trade Log');
  if (!log) return;
  log.getRange('K1:M1').setValues([['Cycle', 'Cash balance', 'Realized P&L to date']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F4E78');
  var last = log.getLastRow();
  if (last >= 2) {
    var f = [];
    for (var r = 2; r <= last; r++) {
      f.push(['=IF(B' + r + '="","",TEXT(A' + r + ',"mmm-yyyy"))',
              '=IF(B' + r + '="","",Settings!$B$2+SUMIFS($G$2:G' + r + ',$D$2:D' + r + ',"SELL")-SUMIFS($G$2:G' + r + ',$D$2:D' + r + ',"BUY"))',
              '=IF(B' + r + '="","",SUM($I$2:I' + r + '))']);
    }
    log.getRange(2, 11, f.length, 3).setFormulas(f);
    log.getRange(2, 12, f.length, 2).setNumberFormat('#,##0;[Red]-#,##0');
    log.getRange(2, 6, f.length, 4).setNumberFormat('#,##0.00');
  }
  rebuildSummary_(ss);
}

function rebuildSummary_(ss) {
  var log = ss.getSheetByName('Trade Log'), hist = ss.getSheetByName('History'), set = ss.getSheetByName('Settings');
  var sum = ss.getSheetByName('Monthly Summary') || ss.insertSheet('Monthly Summary', 2);
  var capital = Number(set.getRange('B2').getValue()) || 0;
  var head = ['Month', 'Cycle', 'Buys', 'Buy value', 'Sells', 'Sell value', 'Stocks sold', 'Stocks bought',
              'Realized P&L', 'Cash at month end', 'Portfolio total at month end (last sync)'];
  var rows = [], months = {}, order = [], cash = capital;
  var n = log.getLastRow() - 1;
  var data = n > 0 ? log.getRange(2, 1, n, 9).getValues() : [];
  data.forEach(function (r) {
    if (!(r[0] instanceof Date) || r[1] === '') return;
    var k = Utilities.formatDate(r[0], TZ, 'yyyy-MM');
    var m = months[k];
    if (!m) { m = months[k] = { label: Utilities.formatDate(r[0], TZ, 'MMM-yyyy'), b: 0, bv: 0, s: 0, sv: 0, sold: [], bought: [], pnl: 0, cash: 0 }; order.push(k); }
    var val = Number(r[6]) || 0;
    if (r[3] === 'BUY') { m.b++; m.bv += val; m.bought.push(r[2]); cash -= val; }
    else { m.s++; m.sv += val; m.sold.push(r[2]); m.pnl += Number(r[8]) || 0; cash += val; }
    m.cash = cash;
  });
  var tot = {};
  if (hist && hist.getLastRow() > 1) hist.getRange(2, 1, hist.getLastRow() - 1, 5).getValues().forEach(function (r) {
    if (r[0] instanceof Date) tot[Utilities.formatDate(r[0], TZ, 'yyyy-MM')] = r[4];
  });
  order.sort().forEach(function (k, i) {
    var m = months[k];
    rows.push([m.label, i === 0 ? 'Initial buy' : 'Rebalance', m.b, m.bv, m.s, m.sv,
      m.sold.join(', '), m.bought.join(', '), m.pnl, m.cash, tot[k] != null ? tot[k] : '']);
  });
  sum.clear();
  sum.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F4E78').setWrap(true);
  if (rows.length) {
    sum.getRange(2, 1, rows.length, head.length).setValues(rows);
    var t = rows.length + 2;
    sum.getRange(t, 1, 1, 9).setValues([['TOTAL', '', '=SUM(C2:C' + (t - 1) + ')', '=SUM(D2:D' + (t - 1) + ')',
      '=SUM(E2:E' + (t - 1) + ')', '=SUM(F2:F' + (t - 1) + ')', '', '', '=SUM(I2:I' + (t - 1) + ')']]).setFontWeight('bold');
    sum.getRange(2, 4, t - 1, 1).setNumberFormat('#,##0');
    sum.getRange(2, 6, t - 1, 1).setNumberFormat('#,##0');
    sum.getRange(2, 9, t - 1, 3).setNumberFormat('#,##0;[Red]-#,##0');
  }
  sum.setFrozenRows(1);
  sum.setColumnWidths(1, 2, 100); sum.setColumnWidths(3, 4, 90); sum.setColumnWidths(7, 2, 260); sum.setColumnWidths(9, 3, 130);
  sum.getRange(rows.length + 4, 1).setValue('Built from Trade Log on every sync (or Kite > Update ledger & monthly summary now). P&L is before brokerage and taxes.');
}


// ---------------- Tab layout + README ----------------
var TABS = [
  // name, colour, hidden, group, purpose
  ['Dashboard', '#0b8043', false, 'VIEW', 'Portfolio totals: invested, current value, unrealized and realized P&L, cash, total return, last sync.'],
  ['Holdings', '#0b8043', false, 'VIEW', 'Stocks you own now: qty, average buy price, buy date, notes, live price, value and P&L per stock. Kept up to date by Sync.'],
  ['Trade Log', '#1a73e8', false, 'LEDGER', 'Every buy and sell in date order (never overwritten): price, value, realized P&L on sells, Cycle (month), running cash balance, running realized P&L.'],
  ['Monthly Summary', '#1a73e8', false, 'LEDGER', 'One row per month: initial buy or rebalance, stocks sold, stocks bought, realized P&L, cash and portfolio total at month end.'],
  ['History', '#1a73e8', false, 'LEDGER', 'Snapshot after every sync: invested, value, cash, portfolio total, realized and total P&L.'],
  ['Scanner', '#f9ab00', false, 'MONTHLY INPUT', 'Paste/import the new scanner CSV here after the 5th (Ctrl+A, Delete, click A1, File > Import > Replace data at selected cell).'],
  ['Rebalance', '#f9ab00', false, 'MONTHLY CHECK', 'Each holding\'s new rank and KEEP/SELL. Type a price in column E only when the Warning column asks.'],
  ['Settings', '#9aa0a6', false, 'SETUP', 'Capital, number of stocks, exit rank, cash (updated by Sync), month-1 flag, price mode (LIVE/LCP), last sync.'],
  ['README', '#9aa0a6', false, 'SETUP', 'This guide.'],
  ['Feed', '#9aa0a6', false, 'SYSTEM', 'Orders the GitHub page reads (published to the web as CSV). Kept visible because hidden tabs may stop being published. Do not edit.'],
  ['Buys', '#dadce0', true, 'HELPER', 'Sizes the new buys for empty slots (cash shared equally).'],
  ['Initial', '#dadce0', true, 'HELPER', 'Month-1 plan: top N by Return %, equal amounts, leftover cash spread.'],
  ['Ranked', '#dadce0', true, 'HELPER', 'Scanner sorted by Return %, with HELD flags and next-buy order.'],
  ['Calc', '#dadce0', true, 'HELPER', 'Reads the Scanner columns by header name, picks the price (LIVE or LCP) and ranks.'],
  ['Orders', '#dadce0', true, 'HELPER', 'Old copy-paste order JSON (backup only; the page uses Feed).']
];

function organizeTabs_(ss) {
  var pos = 1;
  TABS.forEach(function (t) {
    var sh = ss.getSheetByName(t[0]);
    if (!sh) return;
    if (sh.isSheetHidden()) sh.showSheet();
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos++);
    sh.setTabColor(t[1]);
  });
  writeReadme_(ss);
  TABS.forEach(function (t) {
    var sh = ss.getSheetByName(t[0]);
    if (sh && t[2]) sh.hideSheet();
  });
  var d = ss.getSheetByName('Dashboard');
  if (d) ss.setActiveSheet(d);
}

function writeReadme_(ss) {
  var sh = ss.getSheetByName('README') || ss.insertSheet('README');
  sh.clear();
  var R = [];
  var add = function (a, b, c, d) { R.push([a || '', b || '', c || '', d || '']); };
  add('Momentum Portfolio Rebalancer – guide');
  add('');
  add('TABS', 'Shown?', 'Group', 'What it is for');
  TABS.forEach(function (t) { add(t[0], t[2] ? 'Hidden' : 'Visible', t[3], t[4]); });
  add('');
  add('Colours: green = view your portfolio, blue = ledger/records, orange = monthly input/check, grey = setup/system, light grey = hidden helpers.');
  add('To see a hidden tab: View > Hidden sheets (or the tab list button at bottom-left) and pick it. Kite > Organize tabs & refresh README hides them again.');
  add('');
  add('EVERY MONTH (after the 5th)');
  add('1. Scanner tab: Ctrl+A, Delete, click A1, File > Import > upload new CSV > Replace data at selected cell.');
  add('2. Rebalance tab: check KEEP/SELL; if the Warning column asks, type that stock\'s current price in column E.');
  add('3. GitHub page: Reload from sheet > Review sells > OK > place in Kite (use LIMIT, buffer 0.5).');
  add('4. After sells fill: Sync fills to sheet. Then Reload > Review buys > OK > place in Kite > after fills, Sync again (same day).');
  add('5. Check Dashboard, Trade Log and Monthly Summary.');
  add('');
  add('KITE MENU');
  add('Save API key & secret', '', '', 'Kite Connect Personal app keys (stored privately in the script).');
  add('Show links', '', '', 'Redirect/Sync link (/exec) for the Kite app and the GitHub page.');
  add('Update ledger & monthly summary now', '', '', 'Rebuilds ledger columns and Monthly Summary, then re-organizes tabs.');
  add('Organize tabs & refresh README', '', '', 'Re-orders, colours and hides tabs; rewrites this guide.');
  add('Use latest prices for quantities', '', '', 'One-time: size orders on the latest NSE price (Settings B15 LIVE/LCP).');
  add('Run demo', '', '', 'Makes a COPY of this sheet with made-up trades to preview everything.');
  add('');
  add('NOTES');
  add('Only orders sent from the GitHub page (Kite tag kc:pub / MOMRB) are recorded; other Kite trades are ignored.');
  add('P&L is before brokerage and taxes. Live prices come from Google Finance (about 15-20 min delayed).');
  add('Not investment advice. Every order is reviewed and placed by you in Kite.');
  sh.getRange(1, 1, R.length, 4).setValues(R);
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  sh.getRange(3, 1, 1, 4).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F4E78');
  TABS.forEach(function (t, i) { sh.getRange(4 + i, 1).setBackground(t[1]).setFontColor(t[2] ? '#202124' : '#ffffff').setFontWeight('bold'); });
  R.forEach(function (r, i) {
    if (/^(EVERY MONTH|KITE MENU|NOTES)/.test(r[0])) sh.getRange(i + 1, 1).setFontWeight('bold');
  });
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 640);
  sh.getRange(1, 4, R.length, 1).setWrap(true);
}
