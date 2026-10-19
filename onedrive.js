/* OneDrive Fetch and Update for Sasha's Third Eye (Internal Reporting).
 *
 * The production, QA and roster workbooks live in the OneDrive of Purviewautomate@purview-healthcare.com.
 * Each person signs in to Microsoft with their own Purview work account (a pop-up), and every read
 * and write is done as, and recorded against, the person who clicked.
 *
 * Link:   a file's OneDrive share link is pasted once (per client for production, once each for the
 *         QA tracker and the roster, per scope). The links are kept in "Third Eye links.json" in the
 *         "Sashas Third Eye" folder of that OneDrive, so everyone uses the same files.
 * Fetch:  downloads the linked files and loads them into the upload boxes, exactly as if they had
 *         been picked. Every sheet with production headers is read; columns and tabs whose names
 *         start with "Brahmastra" or "Third Eye" are left out.
 * Update: writes Third Eye's results back into each production file, in place, through the Excel
 *         workbook API: its own "Third Eye ..." columns at the right of each sheet and its own
 *         "Third Eye summary" tab. The team's columns and Brahmastra's are never written.
 *
 * Nothing runs on its own: Fetch and Update are buttons. Microsoft only sends people back to the
 * registered address, so the bar shows on the team site only (or with a test mock).
 * Test hook: window.__THIRDEYE_MS_MOCK = { account(), signIn(), token(), forget(), fetch(url, opts) }.
 */
(function () {
  'use strict';
  var MS = {
    clientId: '66b15331-dd65-4904-9ec1-12f60e8f837c',   // the "Purview Brahmastra" app registration
    tenantId: '94be55e5-2194-40e4-9d98-1dcc7e91c268',   // purview-healthcare.com
    redirectUri: 'https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html',
    scopes: ['User.Read', 'Files.ReadWrite.All'],
    owner: 'Purviewautomate@purview-healthcare.com',      // whose OneDrive holds the shared folder
    home: 'Sashas Third Eye',
    linksFile: 'Third Eye links.json',
  };
  if (window.PV_ONEDRIVE_CONFIG) Object.assign(MS, window.PV_ONEDRIVE_CONFIG);
  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var TE_COLS = ['Third Eye Audited', 'Third Eye QA Score', 'Third Eye QA Errors', 'Third Eye Note Check', 'Third Eye Updated By', 'Third Eye Updated On'];
  var SUMMARY_TAB = 'Third Eye summary';
  var SUMMARY_COLS = ['Date Worked', 'Client', 'Team Lead', 'Biller', 'Claims Worked', 'Target', 'Attainment %', 'Claims Audited', 'Sampling %', 'QA Score %', 'QA Errors', 'Updated By', 'Updated On'];
  var OWN = /^\s*(brahmastra|third\s*eye)\b/i;          // other apps' columns and tabs, and our own
  var DATE_RE = /date\s*worked|worked\s*date|date\s*of\s*work|production\s*date|work\s*date|worked\s*on|^dow$/i;
  var WHO_RE = /worked\s*by|biller|associate\s*name|agent|employee\s*name|user\s*name/i;
  var CLAIM_RE = /claim\s*id|claim\s*#|claim\s*no|claim number|charge\s*id|\bicn\b|\bdcn\b/i;
  var CLIENT_RE = /^(client|project)( name)?$/i;
  var BLOCK = 1500;
  var HEAD_FILL = '#16156A';

  function mock() { return window.__THIRDEYE_MS_MOCK || null; }
  function supported() {
    if (mock()) return true;
    try { return location.origin === new URL(MS.redirectUri).origin; } catch (e) { return false; }
  }
  function baseUrl() { return location.href.replace(/[?#].*$/, '').replace(/[^/]*$/, ''); }
  function scopeKey() { return typeof SCOPE !== 'undefined' ? SCOPE : 'ar'; }
  function scopeName() { return typeof SCOPE_NAME !== 'undefined' ? SCOPE_NAME : 'AR'; }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  /* ---------- Microsoft sign-in (MSAL in the browser) ---------- */
  function loadScript(src) { return new Promise(function (res, rej) { var s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = function () { rej(err('network', 'The Microsoft sign-in code did not load. Check your connection and try again.')); }; document.head.appendChild(s); }); }
  var msalP = null;
  function msalApp() {
    if (!msalP) {
      var p = (window.msal ? Promise.resolve() : loadScript(baseUrl() + 'vendor/msal-browser-5.22.0.min.js')).then(function () {
        var app = new window.msal.PublicClientApplication({
          auth: { clientId: MS.clientId, authority: 'https://login.microsoftonline.com/' + MS.tenantId, redirectUri: MS.redirectUri },
          cache: { cacheLocation: 'localStorage' },
          system: { popupBridgeTimeout: 300000 },   // five minutes, time enough for an Authenticator code
        });
        return app.initialize().then(function () { return app; });
      });
      p.catch(function () { if (msalP === p) msalP = null; });
      msalP = p;
    }
    return msalP;
  }
  function err(kind, message, status) { var e = new Error(message); e.kind = kind; e.status = status || 0; return e; }
  function graphError(status) {
    if (status === 401) return err('auth', 'Microsoft asked you to sign in again. Click Sign in and try once more.', status);
    if (status === 403) return err('denied', 'You do not have access to this file in OneDrive yet. Ask Shreya to share it with you (Can edit).', status);
    if (status === 404) return err('missing', 'OneDrive could not find it. The file may have been moved, renamed or deleted. Link it again with Change link.', status);
    if (status === 409 || status === 423) return err('locked', 'The file is busy or open in the Excel desktop app. Close it there, wait a minute and try again.', status);
    if (status === 429 || status >= 500) return err('busy', 'Microsoft is busy right now. Wait a minute and try again.', status);
    return err('other', 'OneDrive turned this down (' + status + '). Try again, and if it keeps happening send it through Feedback.', status);
  }
  function msalMessage(e) {
    if (e && e.kind) return e.message;
    var code = String(e && (e.errorCode || e.code) || '');
    var text = String(e && (e.errorMessage || e.message) || '');
    if (code === 'user_cancelled') return 'Microsoft sign in was closed before it finished.';
    if (code === 'popup_window_error' || code === 'empty_window_error') return 'Your browser blocked the Microsoft sign in window. Allow pop-ups for this site, then click again.';
    if (code === 'timed_out') return 'Microsoft sign in took too long. Click again to retry.';
    if (code === 'interaction_in_progress') return 'A Microsoft sign in window is already open. Finish it there, or close it and click again.';
    if (/AADSTS65001|AADSTS90094|AADSTS90008|consent|approval/i.test(text)) return 'Microsoft needs admin approval for this app first. In Entra, open App registrations, Purview Brahmastra, API permissions, and click Grant admin consent.';
    if (/AADSTS50020|AADSTS50034|AADSTS50177|AADSTS700016|not exist in tenant/i.test(text)) return 'That Microsoft account is not part of purview-healthcare.com. Sign in with your Purview work account, or ask Shreya to invite it as a guest.';
    if (/AADSTS50011|AADSTS50194|reply url|redirect/i.test(text)) return 'Microsoft did not accept this page’s address. It has to be added to Purview Brahmastra as a single-page application redirect URI.';
    if (/AADSTS50076|AADSTS50079|multi-?factor|\bmfa\b/i.test(text)) return 'Microsoft needs you to finish multi-factor sign in. Click again and complete it in the window.';
    return 'Microsoft sign in did not finish' + (code ? ' (' + code + ')' : '') + '. Try again.';
  }
  var msx = {
    account: async function () {
      var m = mock(); if (m) return m.account ? m.account() : null;
      var app = await msalApp();
      return app.getActiveAccount() || app.getAllAccounts()[0] || null;
    },
    signIn: async function () {
      var m = mock(); if (m) return m.signIn();
      var app = await msalApp();
      var r = await app.loginPopup({ scopes: MS.scopes, prompt: 'select_account', overrideInteractionInProgress: true });
      app.setActiveAccount(r.account);
      return r.account;
    },
    token: async function () {
      var m = mock(); if (m) return m.token();
      var app = await msalApp();
      var acct = app.getActiveAccount() || app.getAllAccounts()[0];
      if (!acct) throw err('signin', 'Sign in to Microsoft first.');
      try { return (await app.acquireTokenSilent({ scopes: MS.scopes, account: acct })).accessToken; }
      catch (e) {
        var r = await app.acquireTokenPopup({ scopes: MS.scopes, account: acct, overrideInteractionInProgress: true });
        app.setActiveAccount(r.account);
        return r.accessToken;
      }
    },
    forget: async function () {
      var m = mock(); if (m) { if (m.forget) await m.forget(); return; }
      var app = await msalApp();
      var acct = app.getActiveAccount() || app.getAllAccounts()[0];
      app.setActiveAccount(null);
      if (acct) { try { await app.clearCache({ account: acct }); } catch (e) { /* forgotten on this device either way */ } }
    },
    // retries 429, 503 and 504 a few times, as Microsoft asks
    request: async function (url, opts, auth) {
      opts = opts || {};
      var m = mock(), doFetch = (m && m.fetch) || function (u, o) { return fetch(u, o); };
      for (var attempt = 0; ; attempt++) {
        var headers = Object.assign({}, opts.headers || {});
        if (auth !== false) headers.Authorization = 'Bearer ' + (await msx.token());
        var r;
        try { r = await doFetch(url, { method: opts.method || 'GET', headers: headers, body: opts.body }); }
        catch (e) { throw err('network', 'Could not reach Microsoft. Check your connection and try again.'); }
        if ((r.status === 429 || r.status === 503 || r.status === 504) && attempt < 4) {
          var ra = +(r.headers && r.headers.get && r.headers.get('Retry-After')) || 0;
          await wait(Math.min(30, ra || Math.pow(2, attempt + 1)) * 1000);
          continue;
        }
        if (!r.ok) throw graphError(r.status);
        return r;
      }
    },
    graph: async function (path, opts) {
      opts = opts || {};
      if (opts.json !== undefined) { opts.body = JSON.stringify(opts.json); opts.headers = Object.assign({ 'Content-Type': 'application/json' }, opts.headers || {}); }
      var r = await msx.request(GRAPH + path, opts, true);
      if (r.status === 204) return null;
      var t = await r.text();
      return t ? JSON.parse(t) : null;
    },
    // OneDrive's short-lived download address is already signed, so it is fetched with
    // no Authorization header, which is Microsoft's advice for apps running in a browser
    download: async function (url) {
      var r = await msx.request(url, {}, false);
      return new Uint8Array(await r.arrayBuffer());
    },
  };
  var who = { name: '', email: '' };
  function rememberAccount(a) { who = { name: (a && (a.name || a.username)) || '', email: (a && a.username) || '' }; }

  /* ---------- links: one shared file in the Sashas Third Eye folder ---------- */
  function shareIdOf(link) {
    var b64 = btoa(unescape(encodeURIComponent(String(link).trim())));
    return 'u!' + b64.replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  }
  function itemPath(driveId, id) { return '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(id); }
  var HOME_KEY = 'te-onedrive-home';
  function cachedHome() { try { return JSON.parse(localStorage.getItem(HOME_KEY) || 'null'); } catch (e) { return null; } }
  function cacheHome(h) { try { localStorage.setItem(HOME_KEY, JSON.stringify(h)); } catch (e) { /* found again next time */ } }
  async function findHome() {
    var h = cachedHome();
    if (h && h.driveId && h.folderId) return h;
    var path = '/users/' + encodeURIComponent(MS.owner) + '/drive/root:/' + encodeURIComponent(MS.home);
    try {
      var it = await msx.graph(path + '?$select=id,name,folder,parentReference');
      h = { driveId: it.parentReference.driveId, folderId: it.id };
    } catch (e) {
      // the shared account makes the folder the first time; anyone else is asked for its link
      if (e.status === 404 && who.email.toLowerCase() === MS.owner.toLowerCase()) {
        var made = await msx.graph('/me/drive/root/children', { method: 'POST', json: { name: MS.home, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' } });
        h = { driveId: made.parentReference.driveId, folderId: made.id };
      } else if (e.status === 404 || e.status === 403) {
        throw err('home', 'Third Eye could not open the "' + MS.home + '" folder in ' + MS.owner + '’s OneDrive. Ask Shreya to share that folder with you (Can edit), or paste its link below.');
      } else throw e;
    }
    cacheHome(h);
    return h;
  }
  async function homeFromLink(link) {
    var it = await msx.graph('/shares/' + shareIdOf(link) + '/driveItem');
    if (!it.folder) throw err('input', 'That link opens a file. Paste the link of the "' + MS.home + '" folder itself.');
    var h = { driveId: it.parentReference.driveId, folderId: it.id };
    cacheHome(h);
    return h;
  }
  var EMPTY_LINKS = function () { return { version: 1, scopes: {}, updates: {} }; };
  async function readLinks(home) {
    var meta;
    try { meta = await msx.graph(itemPath(home.driveId, home.folderId) + ':/' + encodeURIComponent(MS.linksFile)); }
    catch (e) { if (e.status === 404) return EMPTY_LINKS(); throw e; }
    var url = meta && meta['@microsoft.graph.downloadUrl'];
    if (!url) return EMPTY_LINKS();
    var bytes = await msx.download(url);
    try { var j = JSON.parse(new TextDecoder().decode(bytes)); j.scopes = j.scopes || {}; j.updates = j.updates || {}; return j; }
    catch (e) { return EMPTY_LINKS(); }
  }
  // re-read just before writing, so two people linking at once do not undo each other
  async function changeLinks(fn) {
    var home = await findHome();
    var cur = await readLinks(home);
    fn(cur);
    await msx.graph(itemPath(home.driveId, home.folderId) + ':/' + encodeURIComponent(MS.linksFile) + ':/content',
      { method: 'PUT', body: JSON.stringify(cur, null, 1), headers: { 'Content-Type': 'application/json' } });
    S.links = cur;
    return cur;
  }
  function scopeLinks(l) { l = l || S.links || EMPTY_LINKS(); var k = scopeKey(); if (!l.scopes[k]) l.scopes[k] = { production: {}, qa: null, roster: null }; if (!l.scopes[k].production) l.scopes[k].production = {}; return l.scopes[k]; }
  async function resolveLink(link, wantExcel) {
    var s = String(link || '').trim();
    if (!/^https:\/\//i.test(s)) throw err('input', 'Paste the whole link, starting with https://');
    var it = await msx.graph('/shares/' + shareIdOf(s) + '/driveItem');
    if (it.folder) throw err('folder', 'That link opens a folder. Open the folder, copy the link of the Excel file itself, and paste that.');
    if (wantExcel && !/\.xlsx$/i.test(it.name || '')) throw err('type', '"' + it.name + '" is not an .xlsx file. Only .xlsx files can be read and written here.');
    if (!wantExcel && !/\.(xlsx|xlsm|xls|csv)$/i.test(it.name || '')) throw err('type', '"' + it.name + '" is not an Excel or CSV file.');
    var pr = it.parentReference || {};
    var folder = (decodeURIComponent(String(pr.path || '').split('root:')[1] || '').replace(/^\//, '').split('/').pop()) || '';
    return { driveId: pr.driveId, itemId: it.id, name: it.name, webUrl: it.webUrl || '', folder: folder, linkedBy: who.name, linkedOn: new Date().toISOString() };
  }

  /* ---------- reading workbooks ---------- */
  async function downloadItem(rec) {
    var it = await msx.graph(itemPath(rec.driveId, rec.itemId));
    var url = it && it['@microsoft.graph.downloadUrl'];
    if (!url) throw err('other', 'OneDrive did not offer "' + rec.name + '" for download. Try again in a minute.');
    return { bytes: await msx.download(url), modified: it.lastModifiedDateTime || '' };
  }
  // a formula or format copied down a whole column makes Excel report a range
  // running to row 1,048,576; read only as far as the last real cell
  function trimRanges(wb) {
    wb.SheetNames.forEach(function (n) {
      var ws = wb.Sheets[n]; if (!ws || !ws['!ref']) return;
      var maxR = 0, maxC = 0;
      Object.keys(ws).forEach(function (k) {
        if (k[0] === '!') return;
        var c = ws[k]; if (!c || c.v === undefined || c.v === null || c.v === '') return;
        var a = XLSX.utils.decode_cell(k); if (a.r > maxR) maxR = a.r; if (a.c > maxC) maxC = a.c;
      });
      var rg = XLSX.utils.decode_range(ws['!ref']); rg.e.r = Math.min(rg.e.r, maxR); rg.e.c = Math.min(rg.e.c, maxC);
      ws['!ref'] = XLSX.utils.encode_range(rg);
    });
  }
  function cellText(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function dateKeyOf(v) { var d = (typeof normDate === 'function') ? normDate(v) : null; return d ? d.key : ''; }
  function nameKeyOf(v) { return typeof normName === 'function' ? normName(v) : cellText(v).toUpperCase(); }
  function claimKey(v) { return cellText(v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
  // every sheet whose header row (within the first 15 rows) has a date-worked and a worked-by column
  function productionSheets(wb) {
    var out = [];
    wb.SheetNames.forEach(function (name) {
      if (OWN.test(name) || /how to fill|^lists?$|validation/i.test(name)) return;
      var aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
      for (var h = 0; h < Math.min(15, aoa.length); h++) {
        var row = (aoa[h] || []).map(cellText);
        var dc = row.findIndex(function (x) { return x && !OWN.test(x) && DATE_RE.test(x); });
        var wc = row.findIndex(function (x) { return x && !OWN.test(x) && WHO_RE.test(x); });
        if (dc < 0 || wc < 0) continue;
        out.push({ name: name, aoa: aoa, h: h, headers: row, dateCol: dc, whoCol: wc,
          claimCol: row.findIndex(function (x) { return x && !OWN.test(x) && CLAIM_RE.test(x); }),
          clientCol: row.findIndex(function (x) { return CLIENT_RE.test(x); }) });
        break;
      }
    });
    return out;
  }
  // header names as the dashboard reads them: blanks dropped, repeats numbered
  function keptHeaders(headers) {
    var seen = {}, out = [];
    headers.forEach(function (h, j) {
      if (!h || OWN.test(h)) return;
      seen[h] = (seen[h] || 0) + 1;
      out.push({ j: j, name: seen[h] === 1 ? h : h + '#' + seen[h] });
    });
    return out;
  }
  function monthBefore(fromKey) {
    var m = String(fromKey).match(/^(\d{4})-(\d{2})/); if (!m) return null;
    var d = new Date(+m[1], +m[2] - 2, 1), e = new Date(+m[1], +m[2] - 1, 0);
    function k(x) { return x.getFullYear() + '-' + String(x.getMonth() + 1).padStart(2, '0') + '-' + String(x.getDate()).padStart(2, '0'); }
    return { from: k(d), to: k(e) };
  }
  function workbookFile(rows, headers, name, validation) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows, { header: headers }), 'Production');
    if (validation) XLSX.utils.book_append_sheet(wb, validation, 'Validation');
    var out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    return new File([out], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  // hand a file to an upload box and wait for the dashboard to take it in
  async function feed(inputId, file, doneWhen) {
    var input = document.getElementById(inputId);
    if (!input) return false;
    var dt = new DataTransfer(); dt.items.add(file);
    input.files = dt.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
    for (var i = 0; i < 300; i++) { await wait(100); if (doneWhen()) return true; }
    return false;
  }

  /* ---------- state ---------- */
  var S = { links: null, signedIn: false, busy: false, fetched: null, editing: null, msg: '', msgKind: '', report: null };
  var FROM_KEY = 'te-onedrive-from', TO_KEY = 'te-onedrive-to', PREV_KEY = 'te-onedrive-prev';
  function todayKey() { var d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
  function pref(k, def) { try { return localStorage.getItem(k) || def; } catch (e) { return def; } }
  function setPref(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* a convenience only */ } }

  async function onFetch() {
    var from = val('odFrom'), to = val('odTo');
    if (!from || !to || from > to) return say('Pick the dates worked: a From date on or before the To date.', 'err');
    setPref(FROM_KEY, from); setPref(TO_KEY, to);
    var withPrev = !!(document.getElementById('odPrev') || {}).checked; setPref(PREV_KEY, withPrev ? '1' : '0');
    var sl = scopeLinks();
    var clients = Object.keys(sl.production).filter(function (c) { return !S.skip || !S.skip[c]; });
    if (!clients.length && !sl.qa && !sl.roster) return say('Link at least one file first.', 'err');
    await busy(async function () {
      var prev = withPrev ? monthBefore(from) : null;
      var cur = [], old = [], headers = [], hseen = {}, validation = null, files = {}, counts = {}, problems = [];
      function addHeader(n) { if (!hseen[n]) { hseen[n] = 1; headers.push(n); } }
      for (var i = 0; i < clients.length; i++) {
        var client = clients[i], rec = sl.production[client];
        say('Reading ' + client + ' (' + (i + 1) + ' of ' + clients.length + ')…');
        try {
          var dl = await downloadItem(rec);
          var wb = XLSX.read(dl.bytes, { type: 'array', cellDates: true });
          trimRanges(wb);
          if (!validation) { var vt = wb.SheetNames.find(function (n) { return /validation/i.test(n); }); if (vt) validation = wb.Sheets[vt]; }
          var sheets = productionSheets(wb);
          if (!sheets.length) { problems.push(client + ': no sheet with Date Worked and Worked By columns'); continue; }
          var fx = { client: client, rec: rec, sheets: [], rows: [], clients: {} };
          var n = 0;
          sheets.forEach(function (sh) {
            var kh = keptHeaders(sh.headers); kh.forEach(function (x) { addHeader(x.name); });
            fx.sheets.push({ name: sh.name, headerRow: sh.h + 1, headers: sh.headers, dateCol: sh.dateCol, whoCol: sh.whoCol, claimCol: sh.claimCol });
            for (var r = sh.h + 1; r < sh.aoa.length; r++) {
              var row = sh.aoa[r]; if (!row || row.every(function (c) { return c === null || c === ''; })) continue;
              var dk = dateKeyOf(row[sh.dateCol]);
              var inCur = dk && dk >= from && dk <= to, inPrev = prev && dk && dk >= prev.from && dk <= prev.to;
              if (!inCur && !inPrev) continue;
              var o = {}; kh.forEach(function (x) { o[x.name] = row[x.j]; });
              if (inCur) {
                cur.push(o); n++;
                var cl = sh.clientCol >= 0 ? cellText(row[sh.clientCol]) : client;
                fx.clients[cl || client] = 1;
                fx.rows.push({ sheet: sh.name, excelRow: r + 1, dateKey: dk, nameKey: nameKeyOf(row[sh.whoCol]),
                  claim: sh.claimCol >= 0 ? claimKey(row[sh.claimCol]) : '', project: cl || client });
              } else old.push(o);
            }
          });
          counts[client] = n;
          files[rec.itemId] = fx;
          rec.lastFetched = { at: new Date().toISOString(), rows: n };
        } catch (e) { problems.push(client + ': ' + msalMessage(e).replace(/\.$/, '')); }
      }
      var stamp = from + ' to ' + to;
      if (sl.roster) {
        say('Reading the team roster…');
        try { var rd = await downloadItem(sl.roster); var before = ROSTER;
          await feed('rosterInput', new File([rd.bytes], sl.roster.name), function () { return ROSTER !== before; });
        } catch (e) { problems.push('Team roster: ' + msalMessage(e).replace(/\.$/, '')); }
      }
      if (sl.qa) {
        say('Reading the QA tracker…');
        try { var qd = await downloadItem(sl.qa); var qb = QA;
          await feed('qaInput', new File([qd.bytes], sl.qa.name), function () { return QA !== qb; });
        } catch (e) { problems.push('QA tracker: ' + msalMessage(e).replace(/\.$/, '')); }
      }
      if (cur.length) {
        say('Loading ' + cur.length.toLocaleString() + ' production rows…');
        var pb = PROD;
        await feed('prodInput', workbookFile(cur, headers, 'OneDrive production ' + stamp + '.xlsx', validation), function () { return PROD !== pb; });
      }
      if (old.length) {
        say('Loading last month’s production for comparison…');
        var bb = (typeof BASE !== 'undefined' && BASE) ? BASE.prod : null;
        await feed('baseProdInput', workbookFile(old, headers, 'OneDrive production ' + prev.from + ' to ' + prev.to + '.xlsx', null), function () { return typeof BASE !== 'undefined' && BASE && BASE.prod !== bb; });
      }
      S.fetched = { from: from, to: to, at: new Date().toISOString(), files: files, counts: counts };
      var got = Object.keys(counts).map(function (c) { return c + ' ' + counts[c].toLocaleString(); }).join(', ');
      say((cur.length ? 'Fetched ' + cur.length.toLocaleString() + ' production rows (' + got + ')' : 'No production rows for those dates') +
        (sl.qa ? ', the QA tracker' : '') + (sl.roster ? ' and the roster' : '') + '.' + (problems.length ? ' Not read: ' + problems.join('; ') + '.' : ''), problems.length ? 'warn' : 'ok');
    });
  }

  /* ---------- what Third Eye writes back ---------- */
  function fmtPctTxt(v) { return v === null || v === undefined || isNaN(v) ? '' : (Math.round(v * 10) / 10) + '%'; }
  function safeText(v) { var s = v === null || v === undefined ? '' : String(v); return /^[=+\-@]/.test(s) ? ' ' + s : s; }
  function stampNow() { var d = new Date(); return d.getDate() + '-' + d.toLocaleString('en-US', { month: 'short' }) + '-' + d.getFullYear() + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function sameClient(a, b) { a = cellText(a).toLowerCase(); b = cellText(b).toLowerCase(); return !a || !b || a === b; }
  // per-claim results for the rows of one file
  function claimResults(fx, from, to) {
    var qaBy = new Map();
    (typeof QA !== 'undefined' ? QA : []).forEach(function (q) {
      if (!q.dateWorked || q.dateWorked < from || q.dateWorked > to) return;
      [q.claimId, q.uni].forEach(function (c) {
        var k = claimKey(c); if (!k) return;
        var key = q.dateWorked + '|' + q.nameKey + '|' + k;
        if (!qaBy.has(key)) qaBy.set(key, []); qaBy.get(key).push(q);
      });
    });
    var det = (typeof PRODDET !== 'undefined' ? PRODDET : []).filter(function (r) { return fx.clients[r.project] && r.dateKey >= from && r.dateKey <= to; });
    var noteOf = new Map(), canCheck = scopeKey() === 'ar' && det.some(function (r) { return r.note; });
    if (canCheck) {
      try { noteForensicsRun(det); } catch (e) { canCheck = false; }
      try { _nfKey = null; } catch (e) { /* the on-screen view recomputes its own flags */ }
    }
    det.forEach(function (r) {
      var k = r.dateKey + '|' + r.nameKey + '|' + claimKey(r.claimId || r.uid);
      if (noteOf.has(k) || !canCheck) return;
      if (!r.note) { noteOf.set(k, 'No note'); return; }
      var found = [];
      try {
        docErrorsOf(r).forEach(function (x) { if (found.indexOf(x) < 0) found.push(x); });
        var d = denialFromSAAG(r);
        if (d.scen) wfGapsOf(r, d.scen).open.forEach(function (x) { found.push('Workflow: ' + x); });
        saagErrorsOf(r).forEach(function (x) { found.push('SAAG: ' + x); });
      } catch (e) { /* a check that cannot run leaves the claim unflagged */ }
      noteOf.set(k, found.length ? found.join('; ') : 'OK');
    });
    var by = who.name, on = stampNow();
    return fx.rows.map(function (row) {
      var k = row.dateKey + '|' + row.nameKey + '|' + row.claim;
      var qs = row.claim ? (qaBy.get(k) || []).filter(function (q) { return sameClient(q.project, row.project); }) : [];
      var score = qs.length ? qs.reduce(function (s, q) { return s + (+q.score || 0); }, 0) / qs.length : null;
      var errs = []; qs.forEach(function (q) { (q.errors || []).forEach(function (e) { var c = e.cat || e.pillar; if (c && errs.indexOf(c) < 0) errs.push(c); }); });
      return { row: row, values: [qs.length ? 'Yes' : 'No', fmtPctTxt(score), errs.join('; '), canCheck ? (noteOf.get(k) || '') : 'Not checked in this scope', by, on] };
    });
  }
  // one row per biller per date, the same figures the dashboard shows with this client picked
  function summaryRows(fx, from, to) {
    var prod = PROD.filter(function (p) { return fx.clients[p.project] && p.dateKey >= from && p.dateKey <= to; });
    var dayTot = new Map();
    PROD.forEach(function (p) { var k = p.nameKey + '||' + p.dateKey; dayTot.set(k, (dayTot.get(k) || 0) + (p.count || 0)); });
    var m = new Map();
    prod.forEach(function (p) {
      var k = p.dateKey + '||' + (p.project || '') + '||' + p.nameKey;
      if (!m.has(k)) m.set(k, { dateKey: p.dateKey, project: p.project, nameKey: p.nameKey, name: p.name, count: 0, target: 0 });
      var r = m.get(k), tot = dayTot.get(p.nameKey + '||' + p.dateKey) || 0;
      r.count += p.count; r.target += tot > 0 ? targetForRow(p) * p.count / tot : targetForRow(p);
    });
    var out = [];
    m.forEach(function (r) {
      var target = Math.round(r.target);
      var aud = QA.filter(function (q) { return q.nameKey === r.nameKey && q.dateWorked === r.dateKey && sameClient(q.project, r.project); });
      out.push({ dateKey: r.dateKey, values: [r.dateKey, r.project, teamLeadOf(r.nameKey) || '', r.name, r.count, target,
        target > 0 ? r.count / target : null, aud.length, r.count > 0 ? aud.length / r.count : null,
        aud.length ? aud.reduce(function (s, q) { return s + (+q.score || 0); }, 0) / aud.length / 100 : null,
        aud.reduce(function (s, q) { return s + (q.errCount || 0); }, 0), who.name, stampNow()] });
    });
    out.sort(function (a, b) { var x = a.values, y = b.values; return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : String(x[1] + x[2] + x[3]).localeCompare(String(y[1] + y[2] + y[3])); });
    return out;
  }
  function hash(s) { var h = 2166136261; for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); }

  /* ---------- Excel workbook API ---------- */
  function colLetter(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function serialOf(dk) { var m = dk.match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000) + 25569 : dk; }
  function Book(rec) {
    var base = itemPath(rec.driveId, rec.itemId) + '/workbook', sid = null;
    function call(path, opts) { opts = opts || {}; opts.headers = Object.assign({ 'workbook-session-id': sid }, opts.headers || {}); return msx.graph(base + path, opts); }
    function ws(id) { return '/worksheets/' + encodeURIComponent(id); }
    function rng(id, a) { return ws(id) + "/range(address='" + a + "')"; }
    return {
      open: async function () { sid = (await msx.graph(base + '/createSession', { method: 'POST', json: { persistChanges: true } })).id; },
      close: async function () { if (sid) { try { await call('/closeSession', { method: 'POST', json: {} }); } catch (e) { /* the session times out by itself */ } sid = null; } },
      sheets: async function () { return ((await call('/worksheets?$select=id,name')) || {}).value || []; },
      addSheet: async function (name) { return call('/worksheets/add', { method: 'POST', json: { name: name } }); },
      used: async function (id) { return call(ws(id) + '/usedRange(valuesOnly=true)?$select=address,columnCount,columnIndex,rowCount,rowIndex,values'); },
      get: async function (id, a) { return ((await call(rng(id, a) + '?$select=values')) || {}).values || []; },
      put: async function (id, a, values, formats) {
        if (formats) await call(rng(id, a), { method: 'PATCH', json: { numberFormat: formats } });
        await call(rng(id, a), { method: 'PATCH', json: { values: values } });
      },
      clear: async function (id, a) { await call(rng(id, a) + '/clear', { method: 'POST', json: { applyTo: 'Contents' } }); },
      style: async function (id, a, width) {
        await call(rng(id, a) + '/format/fill', { method: 'PATCH', json: { color: HEAD_FILL } });
        await call(rng(id, a) + '/format/font', { method: 'PATCH', json: { color: '#FFFFFF', bold: true } });
        if (width) await call(rng(id, a) + '/format', { method: 'PATCH', json: { columnWidth: width } });
      },
    };
  }
  function fill2d(rows, cols, v) { var o = []; for (var i = 0; i < rows; i++) { var r = []; for (var j = 0; j < cols; j++) r.push(v); o.push(r); } return o; }
  function keyMatches(vals, sh, row, minCol) {
    if (dateKeyOf(vals[sh.dateCol - minCol]) !== row.dateKey) return false;
    if (nameKeyOf(vals[sh.whoCol - minCol]) !== row.nameKey) return false;
    if (sh.claimCol >= 0 && claimKey(vals[sh.claimCol - minCol]) !== row.claim) return false;
    return true;
  }
  function blocksOf(results) {
    var rows = results.map(function (r) { return r.row.excelRow; }).sort(function (a, b) { return a - b; });
    var blocks = []; rows.forEach(function (r) { var b = blocks[blocks.length - 1]; if (b && r - b[0] < BLOCK) b[1] = r; else blocks.push([r, r]); });
    return blocks;
  }
  // re-read the key columns of every row about to be written, and stop if rows moved since Fetch
  async function verifySheet(book, sh, sheetId, results) {
    var byRow = new Map(); results.forEach(function (r) { byRow.set(r.row.excelRow, r); });
    var keyCols = [sh.dateCol, sh.whoCol].concat(sh.claimCol >= 0 ? [sh.claimCol] : []);
    var kMin = Math.min.apply(null, keyCols), kMax = Math.max.apply(null, keyCols);
    var blocks = blocksOf(results);
    for (var b = 0; b < blocks.length; b++) {
      var keys = await book.get(sheetId, colLetter(kMin) + blocks[b][0] + ':' + colLetter(kMax) + blocks[b][1]);
      for (var r = blocks[b][0]; r <= blocks[b][1]; r++) {
        var res = byRow.get(r); if (!res) continue;
        if (!keyMatches(keys[r - blocks[b][0]] || [], sh, res.row, kMin)) {
          throw err('moved', 'Rows in the "' + sh.name + '" sheet have moved since Fetch (someone sorted, added or deleted rows). Nothing was written to this file. Click Fetch, then Update again.');
        }
      }
    }
  }
  async function writeSheet(book, sh, sheetId, results, report) {
    // Third Eye's columns: found by name, made just after the last used header the first time
    var used = await book.used(sheetId);
    var lastCol = used ? used.columnIndex + used.columnCount - 1 : sh.headers.length - 1;
    var head = (await book.get(sheetId, 'A' + sh.headerRow + ':' + colLetter(Math.max(lastCol, 0)) + sh.headerRow))[0] || [];
    var pos = {}, lastHeader = -1;
    head.forEach(function (h, j) { var t = cellText(h); if (t) lastHeader = j; var i = TE_COLS.findIndex(function (c) { return c.toLowerCase() === t.toLowerCase(); }); if (i >= 0 && pos[i] === undefined) pos[i] = j; });
    var next = lastHeader + 1;
    for (var i = 0; i < TE_COLS.length; i++) {
      if (pos[i] !== undefined) continue;
      pos[i] = next++;
      var a = colLetter(pos[i]) + sh.headerRow;
      await book.put(sheetId, a, [[TE_COLS[i]]], [['@']]);
      await book.style(sheetId, a, 120);
    }
    // contiguous runs of our columns are written together
    var order = TE_COLS.map(function (c, i) { return { i: i, col: pos[i] }; }).sort(function (a, b) { return a.col - b.col; });
    var runs = []; order.forEach(function (o) { var last = runs[runs.length - 1]; if (last && last[last.length - 1].col === o.col - 1) last.push(o); else runs.push([o]); });
    var byRow = new Map(); results.forEach(function (r) { byRow.set(r.row.excelRow, r); });
    var blocks = blocksOf(results);
    for (var b = 0; b < blocks.length; b++) {
      var r0 = blocks[b][0], r1 = blocks[b][1], n = r1 - r0 + 1;
      for (var k = 0; k < runs.length; k++) {
        var run = runs[k], a2 = colLetter(run[0].col) + r0 + ':' + colLetter(run[run.length - 1].col) + r1;
        // rows in the block that are outside the dates picked keep what they had
        var cur = await book.get(sheetId, a2);
        var vals = [];
        for (var x = 0; x < n; x++) {
          var res = byRow.get(r0 + x), old = cur[x] || [];
          vals.push(run.map(function (o, c) { return res ? safeText(res.values[o.i]) : (old[c] === null || old[c] === undefined ? '' : old[c]); }));
        }
        await book.put(sheetId, a2, vals, fill2d(n, run.length, '@'));
      }
    }
    report.rows += results.length;
  }
  async function writeSummary(book, sheets, rows, fx, from, to) {
    var tab = sheets.find(function (s) { return s.name.toLowerCase() === SUMMARY_TAB.toLowerCase(); });
    var kept = [];
    if (!tab) tab = await book.addSheet(SUMMARY_TAB);
    else {
      var used = await book.used(tab.id);
      var vals = (used && used.values) || [];
      var hdr = (vals[0] || []).map(cellText);
      var di = hdr.indexOf('Date Worked'), ci = hdr.indexOf('Client');
      // rows for other dates or other clients stay as they were
      for (var i = 1; i < vals.length; i++) {
        var v = vals[i]; if (!v || v.every(function (c) { return c === '' || c === null; })) continue;
        var dk = dateKeyOf(v[di]), cl = cellText(v[ci]);
        if (dk && dk >= from && dk <= to && fx.clients[cl]) continue;
        kept.push({ dateKey: dk, values: SUMMARY_COLS.map(function (c) { var j = hdr.indexOf(c); return j >= 0 ? v[j] : ''; }) });
      }
      if (used && used.address) await book.clear(tab.id, used.address.split('!').pop());
    }
    var all = kept.map(function (r) { r.values[0] = r.dateKey || r.values[0]; return r; }).concat(rows);
    all.sort(function (a, b) { var x = a.values, y = b.values; return String(x[0]) < String(y[0]) ? -1 : String(x[0]) > String(y[0]) ? 1 : String(x[1] + x[2] + x[3]).localeCompare(String(y[1] + y[2] + y[3])); });
    var last = colLetter(SUMMARY_COLS.length - 1);
    await book.put(tab.id, 'A1:' + last + '1', [SUMMARY_COLS], [SUMMARY_COLS.map(function () { return '@'; })]);
    await book.style(tab.id, 'A1:' + last + '1', 100);
    var FMT = ['dd-mmm-yyyy', '@', '@', '@', '0', '0', '0.0%', '0', '0.0%', '0.0%', '0', '@', '@'];
    for (var s = 0; s < all.length; s += BLOCK) {
      var part = all.slice(s, s + BLOCK), r0 = s + 2, r1 = s + 1 + part.length;
      var vals2 = part.map(function (r) { return r.values.map(function (v, j) {
        if (j === 0) return /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? serialOf(String(v)) : v;
        if (FMT[j] === '@') return safeText(v);
        return v === null || v === undefined || v === '' || isNaN(v) ? '' : +v;
      }); });
      await book.put(tab.id, 'A' + r0 + ':' + last + r1, vals2, part.map(function () { return FMT; }));
    }
  }

  async function onUpdate() {
    var f = S.fetched;
    if (!f || !Object.keys(f.files).length) return say('Click Fetch first, so Third Eye knows which rows to write to.', 'err');
    if (!PROD.length) return say('There is no production loaded. Click Fetch first.', 'err');
    await busy(async function () {
      var done = [], unchanged = [], problems = [];
      var updates = {};
      var ids = Object.keys(f.files);
      for (var i = 0; i < ids.length; i++) {
        var fx = f.files[ids[i]], rec = fx.rec;
        if (!/\.xlsx$/i.test(rec.name || '')) { problems.push(fx.client + ': only .xlsx files can be written'); continue; }
        say('Working out the results for ' + fx.client + '…');
        var results = claimResults(fx, f.from, f.to), summary = summaryRows(fx, f.from, f.to);
        // only dates whose results changed since the last Update are rewritten
        var last = ((S.links && S.links.updates) || {})[rec.itemId] || {};
        var sigs = {};
        results.forEach(function (r) { var d = r.row.dateKey; sigs[d] = (sigs[d] || '') + r.row.sheet + r.row.excelRow + r.values.slice(0, 4).join('\u0001'); });
        summary.forEach(function (r) { sigs[r.dateKey] = (sigs[r.dateKey] || '') + r.values.slice(0, 11).join('\u0001'); });
        var changed = Object.keys(sigs).filter(function (d) { return !last[d] || last[d].sig !== hash(sigs[d]); });
        if (!changed.length) { unchanged.push(fx.client); continue; }
        var ch = {}; changed.forEach(function (d) { ch[d] = 1; });
        var rs = results.filter(function (r) { return ch[r.row.dateKey]; });
        var book = Book(rec), report = { rows: 0 };
        say('Writing to ' + rec.name + '…');
        try {
          await book.open();
          var sheets = await book.sheets(), plan = [];
          for (var s = 0; s < fx.sheets.length; s++) {
            var sh = fx.sheets[s], mine = rs.filter(function (r) { return r.row.sheet === sh.name; });
            if (!mine.length) continue;
            var ws = sheets.find(function (x) { return x.name === sh.name; });
            if (!ws) throw err('moved', 'The "' + sh.name + '" sheet is no longer in the file. Nothing was written to this file. Click Fetch, then Update again.');
            plan.push({ sh: sh, id: ws.id, rows: mine });
          }
          // every sheet is checked before anything is written to the file
          for (s = 0; s < plan.length; s++) await verifySheet(book, plan[s].sh, plan[s].id, plan[s].rows);
          for (s = 0; s < plan.length; s++) await writeSheet(book, plan[s].sh, plan[s].id, plan[s].rows, report);
          var sumRows = summary.filter(function (r) { return ch[r.dateKey]; });
          var lo = changed.slice().sort()[0], hi = changed.slice().sort().pop();
          // the summary is rewritten for the changed dates only; other dates stay
          await writeSummary(book, sheets, sumRows.concat(summary.filter(function (r) { return !ch[r.dateKey] && r.dateKey >= lo && r.dateKey <= hi; })), fx, lo, hi);
          updates[rec.itemId] = {}; changed.forEach(function (d) { updates[rec.itemId][d] = { sig: hash(sigs[d]), at: new Date().toISOString(), by: who.name }; });
          rec.lastUpdated = { at: new Date().toISOString(), by: who.name };
          done.push(fx.client + ' (' + report.rows.toLocaleString() + ' rows, ' + changed.length + ' date' + (changed.length === 1 ? '' : 's') + ')');
        } catch (e) { problems.push(fx.client + ': ' + msalMessage(e).replace(/\.$/, '')); }
        finally { await book.close(); }
      }
      if (Object.keys(updates).length) {
        try {
          await changeLinks(function (l) {
            Object.keys(updates).forEach(function (id) { l.updates[id] = Object.assign(l.updates[id] || {}, updates[id]); });
            var sl = scopeLinks(l);
            Object.keys(sl.production).forEach(function (c) { var r = sl.production[c]; if (updates[r.itemId]) r.lastUpdated = { at: new Date().toISOString(), by: who.name }; });
          });
        } catch (e) { problems.push('Could not record the update in ' + MS.linksFile + ': ' + msalMessage(e).replace(/\.$/, '')); }
      }
      var unmatched = PROD_UNMATCHED(f);
      S.report = unmatched;
      var parts = [];
      if (done.length) parts.push('Updated ' + done.join(', ') + '.');
      if (unchanged.length) parts.push('Already up to date: ' + unchanged.join(', ') + '.');
      if (problems.length) parts.push('Not updated: ' + problems.join('; ') + '.');
      if (done.length) parts.push('OneDrive version history is the undo.');
      say(parts.join(' '), problems.length ? 'err' : (unmatched.length ? 'warn' : 'ok'));
    });
  }
  // claims the dashboard holds that no row of a linked file carries (for example a production
  // row edited after Fetch); they are listed rather than written somewhere wrong
  function PROD_UNMATCHED(f) {
    var have = new Set();
    Object.keys(f.files).forEach(function (id) { f.files[id].rows.forEach(function (r) { have.add(r.dateKey + '|' + r.nameKey + '|' + r.claim); }); });
    var out = [];
    (typeof PRODDET !== 'undefined' ? PRODDET : []).forEach(function (r) {
      if (r.dateKey < f.from || r.dateKey > f.to) return;
      var linked = Object.keys(f.files).some(function (id) { return f.files[id].clients[r.project]; });
      if (!linked) return;
      if (!have.has(r.dateKey + '|' + r.nameKey + '|' + claimKey(r.claimId || r.uid))) out.push(r);
    });
    return out;
  }

  /* ---------- the bar ---------- */
  function val(id) { var el = document.getElementById(id); return el ? el.value : ''; }
  function say(text, kind) { S.msg = text || ''; S.msgKind = kind || ''; var m = document.getElementById('odMsg'); if (m) { m.textContent = S.msg; m.className = 'od-msg ' + S.msgKind; } }
  async function busy(fn) {
    if (S.busy) return; S.busy = true; draw();
    try { await fn(); } catch (e) { say(msalMessage(e), 'err'); }
    finally { S.busy = false; draw(); }
  }
  function when(iso) { if (!iso) return ''; var d = new Date(iso); return d.getDate() + ' ' + d.toLocaleString('en-US', { month: 'short' }) + ', ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function rowHTML(kind, client, rec) {
    var key = kind + '|' + (client || '');
    var label = kind === 'production' ? 'Production' : kind === 'qa' ? 'QA tracker' : 'Team roster';
    if (S.editing === key) {
      return '<tr class="od-edit"><td>' + label + '</td><td colspan="6"><div class="od-form">' +
        (kind === 'production' ? '<input id="odClient" placeholder="Client, as named in the production file" value="' + esc(client || '') + '">' : '') +
        '<input id="odLink" class="od-link" placeholder="Paste the file’s OneDrive link (right-click the file, Copy link)">' +
        '<button type="button" class="od-btn" data-od="save" data-key="' + esc(key) + '">Save link</button>' +
        '<button type="button" class="od-btn ghost" data-od="cancel">Cancel</button>' +
        (rec ? '<button type="button" class="od-btn ghost danger" data-od="remove" data-key="' + esc(key) + '">Remove</button>' : '') + '</div></td></tr>';
    }
    var fetched = S.fetched && rec && S.fetched.counts && kind === 'production' && S.fetched.counts[client] !== undefined;
    var status = !rec ? '<span class="od-pill warn">Uses Upload</span>'
      : kind !== 'production' ? '<span class="od-pill ok">Read only</span>'
      : fetched ? '<span class="od-pill ok">Fetched</span>' : '<span class="od-pill">Linked</span>';
    return '<tr><td>' + label + '</td><td>' + esc(kind === 'production' ? client : 'All ' + scopeName()) + '</td>' +
      '<td>' + (rec ? (rec.webUrl ? '<a href="' + esc(rec.webUrl) + '" target="_blank" rel="noopener">' + esc((rec.folder ? rec.folder + ' › ' : '') + rec.name) + '</a>' : esc(rec.name)) : '<span class="od-dim">Not linked</span>') + '</td>' +
      '<td><button type="button" class="od-btn sm ghost" data-od="edit" data-key="' + esc(key) + '"' + (S.busy ? ' disabled' : '') + '>' + (rec ? 'Change link' : 'Link file') + '</button></td>' +
      '<td>' + (fetched ? when(S.fetched.at) + ' · ' + S.fetched.counts[client].toLocaleString() + ' rows' : '<span class="od-dim">–</span>') + '</td>' +
      '<td>' + (kind !== 'production' ? '<span class="od-dim">never written</span>' : rec && rec.lastUpdated ? when(rec.lastUpdated.at) + ' by ' + esc(rec.lastUpdated.by || '') : '<span class="od-dim">–</span>') + '</td>' +
      '<td>' + status + '</td></tr>';
  }
  function draw() {
    var host = document.getElementById('odBar'); if (!host) return;
    if (!S.signedIn) {
      host.innerHTML = '<div class="od-card"><div class="od-top"><h3>OneDrive files</h3><span class="od-dim">Fetch the production, QA and roster files straight from OneDrive, and write Third Eye’s results back into the production file.</span><span class="od-grow"></span>' +
        '<button type="button" class="od-btn" data-od="signin"' + (S.busy ? ' disabled' : '') + '>Sign in to Microsoft</button></div><div id="odMsg" class="od-msg ' + S.msgKind + '">' + esc(S.msg) + '</div>' + homeAsk() + '</div>';
      return;
    }
    var sl = scopeLinks();
    var clients = Object.keys(sl.production).sort();
    var dis = S.busy ? ' disabled' : '';
    host.innerHTML = '<div class="od-card"><div class="od-top"><h3>OneDrive files</h3><span class="od-pill ok">Signed in as ' + esc(who.email || who.name) + '</span>' +
      '<button type="button" class="od-link-btn" data-od="signout"' + dis + '>Sign out</button><span class="od-grow"></span>' +
      '<label class="od-dates">Dates worked <input type="date" id="odFrom" value="' + esc(pref(FROM_KEY, todayKey().slice(0, 8) + '01')) + '"> to <input type="date" id="odTo" value="' + esc(pref(TO_KEY, todayKey())) + '"></label>' +
      '<label class="od-check"><input type="checkbox" id="odPrev"' + (pref(PREV_KEY, '1') === '1' ? ' checked' : '') + '> Previous month too</label>' +
      '<button type="button" class="od-btn" data-od="fetch"' + dis + '>Fetch</button>' +
      '<button type="button" class="od-btn dark" data-od="update"' + dis + (S.fetched ? '' : ' title="Fetch first"') + '>Update</button></div>' +
      '<div class="od-table-wrap"><table class="od-table"><thead><tr><th>File</th><th>Client</th><th>Linked file</th><th></th><th>Last fetched</th><th>Last updated in Excel</th><th>Status</th></tr></thead><tbody>' +
      clients.map(function (c) { return rowHTML('production', c, sl.production[c]); }).join('') +
      (S.editing === 'production|' ? rowHTML('production', '', null) : '') +
      rowHTML('qa', '', sl.qa) + rowHTML('roster', '', sl.roster) +
      '</tbody></table></div>' +
      '<div class="od-foot"><button type="button" class="od-link-btn" data-od="add"' + dis + '>+ Link a client’s production file</button></div>' +
      '<div id="odMsg" class="od-msg ' + S.msgKind + '">' + esc(S.msg) + '</div>' + unmatchedHTML() +
      '<div class="od-note">Fetch reads every month sheet with production headers and leaves out every “Brahmastra …” and “Third Eye …” column and tab. Update writes only Third Eye’s own columns and its “' + SUMMARY_TAB + '” tab, for the dates fetched. Clients that aren’t linked keep using the upload boxes below.</div></div>';
  }
  function homeAsk() {
    if (!S.askHome) return '';
    return '<div class="od-form" style="margin-top:8px"><input id="odHome" class="od-link" placeholder="Paste the link of the “' + esc(MS.home) + '” folder"><button type="button" class="od-btn" data-od="home">Use this folder</button></div>';
  }
  function unmatchedHTML() {
    var u = S.report; if (!u || !u.length) return '';
    return '<details class="od-unmatched"><summary>' + u.length + ' claim' + (u.length === 1 ? '' : 's') + ' could not be matched to a row in OneDrive and were skipped</summary><table class="od-table"><thead><tr><th>Client</th><th>Date Worked</th><th>Worked By</th><th>Claim</th></tr></thead><tbody>' +
      u.slice(0, 200).map(function (r) { return '<tr><td>' + esc(r.project) + '</td><td>' + esc(r.dateKey) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.claimId || r.uid) + '</td></tr>'; }).join('') +
      '</tbody></table>' + (u.length > 200 ? '<div class="od-dim">and ' + (u.length - 200) + ' more</div>' : '') + '</details>';
  }
  async function afterSignIn() {
    var a = await msx.account(); if (!a) return;
    rememberAccount(a);
    S.signedIn = true;
    say('Reading the file links…');
    draw();
    try { S.links = await readLinks(await findHome()); S.askHome = false; say(''); }
    catch (e) { if (e.kind === 'home') { S.askHome = true; S.signedIn = false; } say(msalMessage(e), 'err'); }
    draw();
  }
  async function onClick(e) {
    var b = e.target.closest('[data-od]'); if (!b || b.disabled) return;
    var act = b.dataset.od, key = b.dataset.key || '';
    if (act === 'signin') return busy(async function () { await msx.signIn(); await afterSignIn(); });
    if (act === 'signout') { await msx.forget(); S.signedIn = false; S.links = null; S.fetched = null; S.report = null; say(''); return draw(); }
    if (act === 'home') return busy(async function () { await homeFromLink(val('odHome')); S.askHome = false; await afterSignIn(); });
    if (act === 'fetch') return onFetch();
    if (act === 'update') return onUpdate();
    if (act === 'add') { S.editing = 'production|'; draw(); var c = document.getElementById('odClient'); if (c) c.focus(); return; }
    if (act === 'edit') { S.editing = key; draw(); var l = document.getElementById('odLink'); if (l) l.focus(); return; }
    if (act === 'cancel') { S.editing = null; return draw(); }
    if (act === 'remove' || act === 'save') {
      var kind = key.split('|')[0], oldClient = key.split('|').slice(1).join('|');
      var client = kind === 'production' ? val('odClient').trim() : '';
      var link = val('odLink');
      return busy(async function () {
        if (act === 'save') {
          if (kind === 'production' && !client) throw err('input', 'Type the client’s name, as it appears in the production file’s Client column.');
          say('Checking the link…');
          var rec = await resolveLink(link, kind === 'production');
          await changeLinks(function (l) {
            var sl = scopeLinks(l);
            if (kind === 'production') { if (oldClient && oldClient !== client) delete sl.production[oldClient]; sl.production[client] = rec; }
            else sl[kind] = rec;
          });
          say('Linked ' + rec.name + '. Everyone signed in to Third Eye now uses this file.', 'ok');
        } else {
          await changeLinks(function (l) { var sl = scopeLinks(l); if (kind === 'production') delete sl.production[oldClient]; else sl[kind] = null; });
          say('Link removed.', 'ok');
        }
        S.editing = null;
      });
    }
  }

  /* ---------- styles (the page's own colours, so dark mode follows) ---------- */
  var css = document.createElement('style');
  css.textContent = [
    '#odBar .od-card{background:var(--panel);border:1px solid var(--border);border-radius:var(--radius,10px);padding:12px 15px;margin-bottom:12px;box-shadow:var(--shadow);font-size:13px;}',
    '#odBar .od-top{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}',
    '#odBar h3{margin:0;font-size:15px;color:var(--text);}',
    '#odBar .od-grow{flex:1;}',
    '#odBar .od-dim{color:var(--text-faint);}',
    '#odBar .od-dates,#odBar .od-check{display:inline-flex;align-items:center;gap:6px;color:var(--text-dim);}',
    '#odBar input[type=date],#odBar .od-form input{font:inherit;font-size:13px;padding:6px 9px;border:1px solid var(--border);border-radius:8px;background:var(--panel);color:var(--text);}',
    '#odBar .od-btn{font:inherit;font-weight:600;font-size:13px;padding:7px 14px;border:0;border-radius:8px;cursor:pointer;color:var(--panel);background:var(--accent);}',
    '#odBar .od-btn.dark{background:var(--text);color:var(--panel);}',
    '#odBar .od-btn.ghost{background:transparent;color:var(--accent);border:1px solid var(--accent);}',
    '#odBar .od-btn.danger{color:var(--bad);border-color:var(--bad);}',
    '#odBar .od-btn.sm{padding:3px 10px;font-size:12px;}',
    '#odBar .od-btn[disabled]{opacity:.55;cursor:default;}',
    '#odBar .od-link-btn{font:inherit;font-size:12.5px;border:0;background:none;color:var(--accent);cursor:pointer;padding:2px 4px;}',
    '#odBar .od-pill{display:inline-block;border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600;background:var(--accent-dim);color:var(--accent);}',
    '#odBar .od-pill.ok{background:var(--good-dim);color:var(--good);}',
    '#odBar .od-pill.warn{background:var(--warn-dim);color:var(--warn);}',
    '#odBar .od-table-wrap{overflow-x:auto;margin-top:10px;}',
    '#odBar .od-table{border-collapse:collapse;width:100%;}',
    '#odBar .od-table th,#odBar .od-table td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--border-soft);white-space:nowrap;}',
    '#odBar .od-table th{color:var(--text-dim);font-weight:600;font-size:12px;}',
    '#odBar .od-table a{color:var(--accent);}',
    '#odBar .od-form{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
    '#odBar .od-form input{min-width:220px;}',
    '#odBar .od-form .od-link{flex:1;min-width:280px;}',
    '#odBar .od-foot{margin-top:6px;}',
    '#odBar .od-msg{min-height:18px;margin-top:6px;color:var(--text-dim);}',
    '#odBar .od-msg.ok{color:var(--good);}',
    '#odBar .od-msg.warn{color:var(--warn);}',
    '#odBar .od-msg.err{color:var(--bad);}',
    '#odBar .od-note{margin-top:6px;font-size:12px;color:var(--text-faint);}',
    '#odBar .od-unmatched{margin-top:6px;}',
    '#odBar .od-unmatched summary{cursor:pointer;color:var(--warn);}',
    '@media print{#odBar{display:none !important;}}',
  ].join('\n');
  (document.head || document.documentElement).appendChild(css);

  function start() {
    if (!supported()) return;
    var row = document.getElementById('uploadRow'); if (!row) return;
    var host = document.getElementById('odBar');
    if (!host) { host = document.createElement('div'); host.id = 'odBar'; row.parentNode.insertBefore(host, row); }
    host.addEventListener('click', onClick);
    draw();
    // someone already signed in on this device carries straight on
    msx.account().then(function (a) { if (a) busy(afterSignIn); }).catch(function () { /* the Sign in button stays */ });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  window.pvOneDrive = { state: function () { return S; } };
})();
