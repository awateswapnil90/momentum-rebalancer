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
  try {
    return page_('Sync done', syncWithToken_(key, sec, q.request_token));
  } catch (err) {
    return page_('Sync failed', '<p>' + esc_(err.message) + '</p><p>Try the sync link again.</p>');
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

function applyOrders_(orders, kiteHoldings) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(props_().getProperty('SHEET_ID'));
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
  } finally {
    lock.releaseLock();
  }
}
