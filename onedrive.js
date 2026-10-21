/* Upload from OneDrive, for Sasha's Third Eye (Internal and Client Reporting).
 *
 * Every Excel upload box gets an "Upload from OneDrive" button. It opens a small file browser on the
 * OneDrive of Purviewautomate@purview-healthcare.com only, where supervisors keep their folders.
 * The picked file is loaded into that box exactly as if it had been chosen from the computer.
 *
 * The Microsoft sign-in only accepts the Purviewautomate account: any other account that signs in
 * is forgotten again at once, and the person's own OneDrive is never asked for. Uploads only read
 * files. On Internal Reporting, Save to OneDrive keeps the targets, RCA and a period summary for the dates in
 * view in "Third Eye saved work.xlsx", in a folder the supervisor picks; picking the same dates later loads
 * them back, wherever that file was saved.
 *
 * Microsoft only sends people back to the registered address, so the buttons show on the team site
 * only (or with a test mock).
 * Test hook: window.__THIRDEYE_MS_MOCK = { account(), signIn(), token(), forget(), fetch(url, opts) }.
 */
(function () {
  'use strict';
  var MS = {
    clientId: '2e3408ab-40a3-4e0d-8985-8a1474e867e1',   // the "Purview Sasha's Third Eye" app registration
    tenantId: '94be55e5-2194-40e4-9d98-1dcc7e91c268',   // purview-healthcare.com
    redirectUri: 'https://purview-healthcare.github.io/Sashas-third-eye/auth-redirect.html',
    scopes: ['User.Read', 'Files.ReadWrite.All'],       // as consented; Third Eye itself only reads
    owner: 'Purviewautomate@purview-healthcare.com',      // the only OneDrive the browser opens
  };
  if (window.PV_ONEDRIVE_CONFIG) Object.assign(MS, window.PV_ONEDRIVE_CONFIG);
  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var SAVED_NAME = 'Third Eye saved work.xlsx';          // a supervisor's saved targets, RCA and period summaries
  var DATA_TAB = 'Third Eye data';                        // hidden tab Third Eye reads back; the other tabs are for people
  var KNOWN_KEY = 'te-saved-work-files';                  // saved work files this browser wrote, as search can lag
  var NAME_KEY = 'pvfb-name';                             // the name the Feedback panel already asks for
  var S = { folders: {}, lastFolder: null, sig: '', checked: '', per: [], busy: false };

  function mock() { return window.__THIRDEYE_MS_MOCK || null; }
  function supported() {
    if (mock()) return true;
    try { return location.origin === new URL(MS.redirectUri).origin; } catch (e) { return false; }
  }
  function baseUrl() { return location.href.replace(/[?#].*$/, '').replace(/[^/]*$/, ''); }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function isOwner(a) { return !!a && String(a.username || '').toLowerCase() === MS.owner.toLowerCase(); }

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
    if (status === 401) return err('auth', 'Microsoft asked you to sign in again. Click Upload from OneDrive once more.', status);
    if (status === 403) return err('denied', 'Purview Automate’s OneDrive turned this down. Check that you signed in as ' + MS.owner + '.', status);
    if (status === 404) return err('missing', 'OneDrive could not find it. It may have been moved, renamed or deleted. Go back a folder and look again.', status);
    if (status === 409 || status === 423) return err('locked', 'The file is open in the Excel desktop app or busy. Close it there, wait a minute and try again.', status);
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
    if (/AADSTS65001|AADSTS90094|AADSTS90008|consent|approval/i.test(text)) return 'Microsoft needs admin approval for this app first. In Entra, open App registrations, Purview Sasha’s Third Eye, API permissions, and click Grant admin consent.';
    if (/AADSTS50020|AADSTS50034|AADSTS50177|AADSTS700016|not exist in tenant/i.test(text)) return 'That Microsoft account is not part of purview-healthcare.com. Sign in as ' + MS.owner + '.';
    if (/AADSTS50011|AADSTS50194|reply url|redirect/i.test(text)) return 'Microsoft did not accept this page’s address. It has to be added to Purview Sasha’s Third Eye as a single-page application redirect URI.';
    if (/AADSTS50076|AADSTS50079|multi-?factor|\bmfa\b/i.test(text)) return 'Microsoft needs you to finish multi-factor sign in. Click again and complete it in the window.';
    if (/AADSTS50126|invalid username or password/i.test(text)) return 'Microsoft did not accept that password. Click again and re-enter it.';
    return 'Microsoft sign in did not finish' + (code ? ' (' + code + ')' : '') + '. Try again.';
  }
  var wrongAccount = function (a) {
    return err('wrong', 'You signed in as ' + ((a && a.username) || 'another account') + '. Upload from OneDrive only works with ' + MS.owner + '. Click again and sign in with that account.');
  };
  var msx = {
    // only ever the Purviewautomate account; any other account cached on this device is left alone
    account: async function () {
      var m = mock(); if (m) { var ma = m.account ? await m.account() : null; return isOwner(ma) ? ma : null; }
      var app = await msalApp();
      return app.getAllAccounts().filter(isOwner)[0] || null;
    },
    signIn: async function () {
      var m = mock(), a;
      if (m) a = await m.signIn();
      else {
        var app = await msalApp();
        a = (await app.loginPopup({ scopes: MS.scopes, loginHint: MS.owner, prompt: 'login', overrideInteractionInProgress: true })).account;
      }
      if (!isOwner(a)) { await msx.forget(a); throw wrongAccount(a); }
      return a;
    },
    token: async function () {
      var m = mock(); if (m) return m.token();
      var app = await msalApp();
      var acct = await msx.account();
      if (!acct) acct = await msx.signIn();
      try { return (await app.acquireTokenSilent({ scopes: MS.scopes, account: acct })).accessToken; }
      catch (e) {
        var r = await app.acquireTokenPopup({ scopes: MS.scopes, account: acct, loginHint: MS.owner, overrideInteractionInProgress: true });
        if (!isOwner(r.account)) { await msx.forget(r.account); throw wrongAccount(r.account); }
        return r.accessToken;
      }
    },
    forget: async function (a) {
      var m = mock(); if (m) { if (m.forget) await m.forget(a); return; }
      var app = await msalApp();
      var acct = a || app.getAllAccounts().filter(isOwner)[0];
      if (acct) { try { await app.clearCache({ account: acct }); } catch (e) { /* forgotten on this device either way */ } }
    },
    // retries 429, 503 and 504 a few times, as Microsoft asks
    request: async function (url, auth, opts) {
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
    graph: async function (pathOrUrl) {
      var r = await msx.request(/^https:/.test(pathOrUrl) ? pathOrUrl : GRAPH + pathOrUrl, true);
      var t = await r.text();
      return t ? JSON.parse(t) : null;
    },
    // OneDrive's short-lived download address is already signed, so it is fetched with
    // no Authorization header, which is Microsoft's advice for apps running in a browser
    download: async function (url) {
      var r = await msx.request(url, false);
      return await r.arrayBuffer();
    },
  };

  /* ---------- the folder browser ---------- */
  var B = null;   // { input, crumbs:[{id,name}], items, filter, picked, busy, msg, msgKind, exts }
  function exts(input) {
    var a = String(input.getAttribute('accept') || '').toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s.charAt(0) === '.'; });
    return a.length ? a : ['.xlsx', '.xlsm', '.xls', '.csv'];
  }
  function fits(name) { if (B.mode === 'save') return name === SAVED_NAME; var n = String(name).toLowerCase(); return B.exts.some(function (x) { return n.slice(-x.length) === x; }); }
  function when(iso) { if (!iso) return ''; var d = new Date(iso); return d.getDate() + ' ' + d.toLocaleString('en-US', { month: 'short' }) + ' ' + d.getFullYear() + ', ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function size(n) { if (!n && n !== 0) return ''; return n < 1024 * 1024 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (Math.round(n / 104857.6) / 10) + ' MB'; }
  function title(input) {
    var host = input.closest('.upload-zone, .slot, .drop');
    var t = host && host.querySelector('.u-title, .ttl');
    var s = t ? t.childNodes[0] && t.childNodes[0].nodeType === 3 ? t.childNodes[0].textContent : t.textContent : '';
    s = String(s || '').trim();
    return s;
  }

  async function listFolder(id) {
    var path = id ? '/me/drive/items/' + encodeURIComponent(id) + '/children' : '/me/drive/root/children';
    var url = path + '?$select=id,name,folder,file,size,lastModifiedDateTime,lastModifiedBy,shared&$top=999', out = [];
    while (url) { var page = await msx.graph(url); out = out.concat(page.value || []); url = page['@odata.nextLink'] || ''; }
    return out.sort(function (a, b) { return (b.folder ? 1 : 0) - (a.folder ? 1 : 0) || a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }); });
  }
  async function openFolder(crumbs) {
    B.busy = true; B.picked = null; B.filter = ''; say('Opening ' + crumbs[crumbs.length - 1].name + '…'); draw();
    try {
      B.items = await listFolder(crumbs[crumbs.length - 1].id);
      B.crumbs = crumbs; say('');
    } catch (e) {
      // a remembered start folder that has gone falls back to My files
      if (e.status === 404 && crumbs.length > 1 && !B.crumbs.length) { B.busy = false; return openFolder([crumbs[0]]); }
      say(msalMessage(e), 'err');
    }
    B.busy = false; draw();
  }
  async function pick() {
    if (B.mode === 'save') return saveHere();
    var it = B.picked && B.items.filter(function (x) { return x.id === B.picked; })[0];
    if (!it) return;
    B.busy = true; say('Downloading ' + it.name + '…'); draw();
    try {
      var meta = await msx.graph('/me/drive/items/' + encodeURIComponent(it.id) + '?$select=id,name,@microsoft.graph.downloadUrl');
      var data = await msx.download(meta['@microsoft.graph.downloadUrl']);
      var file = new File([data], it.name, { type: it.file && it.file.mimeType || '', lastModified: Date.parse(it.lastModifiedDateTime) || Date.now() });
      var input = B.input;
      S.folders[input.id || 'other'] = B.crumbs.slice();   // where this box's file came from
      S.lastFolder = B.crumbs.slice();
      close();
      var dt = new DataTransfer(); dt.items.add(file);
      input.files = dt.files;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    } catch (e) { B.busy = false; say(msalMessage(e), 'err'); draw(); }
  }
  function say(text, kind) { if (B) { B.msg = text || ''; B.msgKind = kind || ''; } }

  var FOLDER_ICON = '<svg class="od-ico" width="24" height="20" viewBox="0 0 24 20" aria-hidden="true"><path d="M1 3a2 2 0 012-2h6l2 2.5h10a2 2 0 012 2V17a2 2 0 01-2 2H3a2 2 0 01-2-2z" fill="#FFB900"/><path d="M1 6h22v11a2 2 0 01-2 2H3a2 2 0 01-2-2z" fill="#FFD75E"/></svg>';
  var FILE_ICON = '<svg class="od-ico" width="20" height="22" viewBox="0 0 20 22" aria-hidden="true"><path d="M4 1h9l5 5v14a1 1 0 01-1 1H4a1 1 0 01-1-1V2a1 1 0 011-1z" fill="#fff" stroke="#C8C8C8"/><rect x="0" y="8" width="11" height="10" rx="1.5" fill="#107C41"/><path d="M3 10.5l5 5M8 10.5l-5 5" stroke="#fff" stroke-width="1.6"/></svg>';
  function draw() {
    var host = document.getElementById('odPick'); if (!host || !B) return;
    var q = B.filter.toLowerCase();
    var rows = (B.items || []).filter(function (x) { return (x.folder || fits(x.name)) && (!q || x.name.toLowerCase().indexOf(q) >= 0); });
    var hidden = (B.items || []).filter(function (x) { return !x.folder && !fits(x.name); }).length;
    var body;
    if (!B.signedIn) {
      body = '<div class="od-empty"><p>Sign in with <b>' + esc(MS.owner) + '</b> to open its OneDrive. Other accounts are not accepted here.</p>' +
        '<button type="button" class="od-btn" data-od="signin"' + (B.busy ? ' disabled' : '') + '>Sign in to Purview Automate’s OneDrive</button></div>';
    } else {
      body = '<div class="od-crumbs">' + B.crumbs.map(function (c, i) {
        return i === B.crumbs.length - 1 ? '<span class="od-here">' + esc(c.name) + '</span>' : '<button type="button" class="od-crumb" data-od="crumb" data-i="' + i + '">' + esc(c.name) + '</button><span class="od-sep">›</span>';
      }).join('') + '<input type="search" class="od-filter" id="odFilter" placeholder="Search this folder" value="' + esc(B.filter) + '"></div>' +
        '<div class="od-list"><table><thead><tr><th class="od-name">Name</th><th>Modified</th><th>Modified By</th><th>File size</th><th>Sharing</th></tr></thead>' +
        (rows.length ? '<tbody>' + rows.map(function (x) {
          var cls = x.id === B.picked ? ' class="sel"' : '';
          var by = x.lastModifiedBy && x.lastModifiedBy.user && x.lastModifiedBy.user.displayName || '';
          var n = x.folder && x.folder.childCount;
          return '<tr data-od="' + (x.folder ? 'open' : 'select') + '" data-id="' + esc(x.id) + '"' + cls + '><td class="od-name"><div class="od-nm">' + (x.folder ? FOLDER_ICON : FILE_ICON) + '<span>' + esc(x.name) + '</span></div></td>' +
            '<td>' + esc(when(x.lastModifiedDateTime)) + '</td><td>' + esc(by) + '</td>' +
            '<td>' + (x.folder ? (n != null ? n + ' item' + (n === 1 ? '' : 's') : '') : size(x.size)) + '</td>' +
            '<td>' + (x.shared ? 'Shared' : 'Private') + '</td></tr>';
        }).join('') + '</tbody></table>' : '</table><div class="od-empty">' + (B.busy ? 'Loading\u2026' : q ? 'Nothing in this folder matches.' : B.mode === 'save' ? 'No saved work in this folder yet. Click Save here to create ' + SAVED_NAME + ' in it.' : 'This folder has no files that can be uploaded here.') + '</div>') + '</div>';
    }
    host.innerHTML = '<div class="od-box" role="dialog" aria-modal="true" aria-label="OneDrive">' +
      '<div class="od-hd"><span class="od-brand">OneDrive</span><span class="od-who">' + esc(MS.owner) + '</span>' +
      '<button type="button" class="od-x" data-od="close" aria-label="Close">×</button></div>' +
      (B.mode === 'save' ? '<div class="od-for">Saving <b>' + esc(B.title) + '</b>: the targets, RCA &amp; mitigation plans and period summary. Open the folder to keep it in, then click <b>Save here</b>.</div>'
        : B.title ? '<div class="od-for">Uploading to <b>' + esc(B.title) + '</b>. Open the folder, click the file, then click Upload.</div>' : '') + body +
      '<div class="od-msg ' + B.msgKind + '">' + esc(B.msg) + '</div>' +
      '<div class="od-ft"><span class="od-dim">' + (B.signedIn && B.mode === 'save' ? '<label class="od-name-in">Your name <input id="odWho" value="' + esc(B.who) + '" placeholder="Shown as Saved by"></label> ' +
        '<button type="button" class="od-crumb" data-od="signout">Sign out of OneDrive</button>' : B.signedIn ? 'Only ' + esc(B.exts.join(', ')) + ' files are listed' + (hidden ? ' (' + hidden + ' other file' + (hidden === 1 ? '' : 's') + ' hidden)' : '') + '. ' +
        '<button type="button" class="od-crumb" data-od="signout">Sign out of OneDrive</button>' : '') + '</span>' +
      '<span><button type="button" class="od-btn ghost" data-od="close">Cancel</button>' +
      (B.signedIn ? '<button type="button" class="od-btn" data-od="pick"' + ((B.picked || B.mode === 'save') && !B.busy ? '' : ' disabled') + '>' + (B.mode === 'save' ? 'Save here' : 'Upload') + '</button>' : '') + '</span></div></div>';
    var f = document.getElementById('odFilter');
    if (f && B.filterFocus) { f.focus(); f.setSelectionRange(f.value.length, f.value.length); }
  }
  function close() {
    var host = document.getElementById('odPick'); if (host) host.remove();
    document.removeEventListener('keydown', onKey, true);
    B = null;
  }
  function onKey(e) { if (e.key === 'Escape' && B) { e.stopPropagation(); close(); } }

  async function afterSignIn() {
    B.signedIn = true;
    // uploads always start at My files, as in OneDrive; a save starts where the production file came from
    await openFolder(B.start && B.start.length ? B.start : [{ id: '', name: 'My files' }]);
  }
  async function open(input, opt) {
    close();
    opt = opt || {};
    B = { input: input, mode: opt.mode || 'upload', start: opt.start || null, who: opt.who || '', title: opt.title || title(input), exts: input ? exts(input) : ['.xlsx'], crumbs: [], items: [], filter: '', picked: null, busy: true, signedIn: false, msg: '', msgKind: '' };
    var host = document.createElement('div'); host.id = 'odPick';
    document.body.appendChild(host);
    host.addEventListener('click', onClick);
    host.addEventListener('dblclick', onDbl);
    host.addEventListener('input', function (e) { if (e.target.id === 'odWho') { B.who = e.target.value; return; } if (e.target.id === 'odFilter') { B.filter = e.target.value; B.filterFocus = true; draw(); B.filterFocus = false; } });
    document.addEventListener('keydown', onKey, true);
    draw();
    try { if (await msx.account()) await afterSignIn(); else { B.busy = false; draw(); } }
    catch (e) { if (B) { B.busy = false; say(msalMessage(e), 'err'); draw(); } }
  }
  async function onClick(e) {
    if (e.target.id === 'odPick') return close();   // a click outside the box
    var b = e.target.closest('[data-od]'); if (!b || b.disabled || !B) return;
    var act = b.dataset.od;
    if (act === 'close') return close();
    if (B.busy && act !== 'select') return;
    if (act === 'signin') {
      B.busy = true; say('Waiting for the Microsoft sign in window…'); draw();
      try { await msx.signIn(); await afterSignIn(); }
      catch (x) { if (B) { B.busy = false; say(msalMessage(x), 'err'); draw(); } }
      return;
    }
    if (act === 'signout') { await msx.forget(); B.signedIn = false; B.items = []; B.picked = null; say('Signed out of Purview Automate’s OneDrive on this computer.', 'ok'); return draw(); }
    if (act === 'crumb') return openFolder(B.crumbs.slice(0, +b.dataset.i + 1));
    if (act === 'open') { var f = B.items.filter(function (x) { return x.id === b.dataset.id; })[0]; if (f) return openFolder(B.crumbs.concat([{ id: f.id, name: f.name }])); return; }
    if (act === 'select') {
      // marked in place, so the row stays the same element for a double click
      B.picked = b.dataset.id;
      document.querySelectorAll('#odPick tr[data-od]').forEach(function (r) { r.classList.toggle('sel', r === b); });
      var up = document.querySelector('#odPick [data-od="pick"]'); if (up) up.disabled = B.busy;
      return;
    }
    if (act === 'pick') return pick();
  }
  function onDbl(e) { var r = e.target.closest('[data-od="select"]'); if (r && B && !B.busy && B.mode !== 'save') { B.picked = r.dataset.id; pick(); } }

  /* ---------- a button on every Excel upload box ---------- */
  var ICON = '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M10.5 18H6a4 4 0 01-.6-7.96A6 6 0 0117 8.5a4.5 4.5 0 01.5 8.97V18z"/></svg>';
  function wants(input) {
    if (input.dataset.odDone) return false;
    var acc = String(input.getAttribute('accept') || '').toLowerCase();
    return /\.xlsx|\.csv/.test(acc);   // Excel and CSV boxes; not logos or exported .json views
  }
  function addButtons(root) {
    (root.querySelectorAll ? root : document).querySelectorAll('input[type=file]').forEach(function (input) {
      if (!wants(input)) return;
      var host = input.closest('.upload-zone, .slot, .drop, label.btn');
      if (!host) return;
      input.dataset.odDone = '1';
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'od-up no-print'; b.innerHTML = ICON + 'Upload from OneDrive';
      b.addEventListener('click', function (e) { e.preventDefault(); e.stopPropagation(); open(input); });
      if (host.matches('label.btn')) host.insertAdjacentElement('afterend', b);       // the small SAAG button
      else if (host.matches('.upload-zone')) { var row = document.createElement('div'); row.className = 'od-row'; row.appendChild(b); host.appendChild(row); }   // its own line, so the counts keep their room
      else { var pk = host.querySelector('.pick, button'); if (pk) pk.insertAdjacentElement('afterend', b); else host.appendChild(b); }
    });
  }

  /* ---------- saved work: targets, RCA and a period summary per date range ---------- */
  // Internal Reporting only. A save is one record per scope, date range and client, kept in the
  // "Third Eye data" tab as JSON; the other tabs are rebuilt from it for people to read. Each
  // client's record carries only the targets and RCA that touch that client, so loading client A
  // never changes what was set for client B.
  function page() {
    try { return typeof TARGETS !== 'undefined' && typeof RCA !== 'undefined' && typeof PROD !== 'undefined' && typeof render === 'function'; }
    catch (e) { return false; }
  }
  function scopeCode() { try { return typeof SCOPE !== 'undefined' ? SCOPE : 'ar'; } catch (e) { return 'ar'; } }
  function scopeLabel() { try { return typeof SCOPE_NAME !== 'undefined' ? SCOPE_NAME : 'AR'; } catch (e) { return 'AR'; } }
  // the production dates in view (the audit dates while the two are linked), else the data's own span
  function period() {
    var link = true, f = '', t = '';
    try { link = DATE_LINK; f = link ? DATE_FROM : PROD_FROM; t = link ? DATE_TO : PROD_TO; } catch (e) { /* not this page */ }
    if (!f || !t) {
      var ds = PROD.map(function (p) { return p.dateKey; }).filter(Boolean).sort();
      f = f || ds[0] || ''; t = t || ds[ds.length - 1] || '';
    }
    return { from: f, to: t };
  }
  function inRange(dk, pr) { return !!dk && dk >= pr.from && dk <= pr.to; }
  function clientsIn(pr) {
    var c = {};
    PROD.forEach(function (p) { if (p.project && inRange(p.dateKey, pr)) c[p.project] = 1; });
    (typeof QA !== 'undefined' ? QA : []).forEach(function (q) { if (q.project && inRange(q.dateWorked, pr)) c[q.project] = 1; });
    return Object.keys(c).sort();
  }
  function dayLabel(dk) {
    var m = String(dk || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) return dk || '';
    return +m[3] + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+m[2] - 1] + ' ' + m[1];
  }
  function rangeLabel(r) { return dayLabel(r.from) + ' to ' + dayLabel(r.to); }
  function overlap(a, b) { return a.some(function (x) { return b.indexOf(x) >= 0; }); }
  function clone(o) { return JSON.parse(JSON.stringify(o || {})); }
  function myName() { try { return localStorage.getItem(NAME_KEY) || ''; } catch (e) { return ''; } }
  function knownFiles() { try { return JSON.parse(localStorage.getItem(KNOWN_KEY) || '[]'); } catch (e) { return []; } }
  function rememberFile(id) { try { var k = knownFiles().filter(function (x) { return x !== id; }); k.unshift(id); localStorage.setItem(KNOWN_KEY, JSON.stringify(k.slice(0, 30))); } catch (e) { /* search finds it later */ } }
  function hasRca(e) { return !!(e && (e.why || e.rca || e.mitigation || e.owner || e.due || e.status)); }

  // one row per client and biller over the range, the same share-of-day targets the report uses
  function periodSummary(pr, clients) {
    var dayTot = new Map();
    PROD.forEach(function (p) { var k = p.nameKey + '||' + p.dateKey; dayTot.set(k, (dayTot.get(k) || 0) + (p.count || 0)); });
    var m = new Map();
    PROD.forEach(function (p) {
      if (!inRange(p.dateKey, pr) || clients.indexOf(p.project) < 0) return;
      var k = p.project + '||' + p.nameKey;
      if (!m.has(k)) m.set(k, { project: p.project, nameKey: p.nameKey, name: p.name || p.nameKey, count: 0, target: 0 });
      var r = m.get(k), tot = dayTot.get(p.nameKey + '||' + p.dateKey) || 0, t = 0;
      try { t = targetForRow(p); } catch (e) { t = 0; }
      r.count += p.count || 0; r.target += tot > 0 ? t * (p.count || 0) / tot : t;
    });
    var qa = (typeof QA !== 'undefined' ? QA : []).filter(function (q) { return inRange(q.dateWorked, pr); });
    var out = [];
    m.forEach(function (r) {
      var aud = qa.filter(function (q) { return q.nameKey === r.nameKey && (!q.project || q.project === r.project); });
      var cats = {}; aud.forEach(function (q) { (q.errors || []).forEach(function (e) { var c = e.cat || e.pillar; if (c) cats[c] = (cats[c] || 0) + 1; }); });
      var top = Object.keys(cats).sort(function (a, b) { return cats[b] - cats[a]; })[0] || '';
      var tl = ''; try { tl = teamLeadOf(r.nameKey) || ''; } catch (e) { /* no roster */ }
      out.push({ client: r.project, teamLead: tl, biller: r.name, claims: r.count, target: Math.round(r.target), audited: aud.length,
        qa: aud.length ? aud.reduce(function (s, q) { return s + (+q.score || 0); }, 0) / aud.length : null,
        errors: aud.reduce(function (s, q) { return s + (q.errCount || 0); }, 0), topError: top });
    });
    return out.sort(function (a, b) { return (a.client + a.teamLead + a.biller).localeCompare(b.client + b.teamLead + b.biller); });
  }
  // the targets that reach one client's rows: its own client and practice figures, and the billers,
  // team leads and weeks that worked it in the range (a biller's day target is shared across clients)
  function targetsFor(c, pr) {
    var t = TARGETS || {}, out = { __default__: t.__default__, byProject: {}, byPractice: {}, byName: {}, byWeek: {}, notes: {} };
    var keys = { practices: [], billers: [], teamLeads: [], weeks: [] }, add = function (a, v) { if (v && a.indexOf(v) < 0) a.push(v); };
    PROD.forEach(function (p) {
      if (p.project !== c || !inRange(p.dateKey, pr)) return;
      add(keys.practices, p.practice || ''); add(keys.billers, p.nameKey);
      try { add(keys.teamLeads, teamLeadOf(p.nameKey) || ''); add(keys.weeks, weekKeyOf(p.dateKey)); } catch (e) { /* not this page */ }
    });
    if ((t.byProject || {})[c] > 0) out.byProject[c] = t.byProject[c];
    keys.practices.forEach(function (k) { if ((t.byPractice || {})[k] > 0) out.byPractice[k] = t.byPractice[k]; });
    keys.billers.forEach(function (k) { if ((t.byName || {})[k] > 0) out.byName[k] = t.byName[k]; });
    keys.weeks.forEach(function (wk) {
      var w = (t.byWeek || {})[wk]; if (!w) return;
      var o = { tl: {}, name: {} };
      keys.teamLeads.forEach(function (k) { if ((w.tl || {})[k] > 0) o.tl[k] = w.tl[k]; });
      keys.billers.forEach(function (k) { if ((w.name || {})[k] > 0) o.name[k] = w.name[k]; });
      if (Object.keys(o.tl).length || Object.keys(o.name).length) out.byWeek[wk] = o;
    });
    var noteFor = { 'default|': 1, 'project|': [c], 'practice|': keys.practices, 'biller|': keys.billers };
    Object.keys(t.notes || {}).forEach(function (k) {
      var i = k.indexOf('|') + 1, lvl = k.slice(0, i), who = k.slice(i), m = noteFor[lvl];
      if (m === 1 || (m && m.indexOf(who) >= 0)) out.notes[k] = clone(t.notes[k]);
    });
    return { targets: out, keys: keys };
  }
  function currentRecords(who) {
    var pr = period(), names = {}, at = new Date().toISOString();
    PROD.forEach(function (p) { if (p.nameKey && p.name) names[p.nameKey] = p.name; });   // so the Targets tab shows people's names
    return clientsIn(pr).map(function (c) {
      var rca = {}, tf = targetsFor(c, pr), nm = {};
      Object.keys(RCA || {}).forEach(function (k) { if (k.indexOf(c + '||') === 0 && hasRca(RCA[k])) rca[k] = clone(RCA[k]); });
      tf.keys.billers.forEach(function (k) { if (names[k]) nm[k] = names[k]; });
      return { scope: scopeCode(), from: pr.from, to: pr.to, clients: [c], by: who, at: at,
        targets: tf.targets, keys: tf.keys, rca: rca, summary: periodSummary(pr, [c]), names: nm };
    });
  }

  /* reading and writing the workbook (SheetJS is already on the page) */
  function readRecords(buf) {
    var wb = XLSX.read(buf, { type: 'array' });
    var ws = wb.Sheets[DATA_TAB]; if (!ws) return { wb: wb, recs: [] };
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }), recs = [];
    rows.slice(1).forEach(function (r) {
      try { var j = JSON.parse(r.slice(6).join('')); recs.push(Object.assign(j, { scope: String(r[0]), from: String(r[1]), to: String(r[2]), clients: String(r[3]).split('\n').filter(Boolean), by: String(r[4]), at: String(r[5]) })); }
      catch (e) { /* a damaged row is skipped, the rest still load */ }
    });
    return { wb: wb, recs: recs };
  }
  function cell(v) { var x = v === null || v === undefined ? '' : v; return typeof x === 'string' && /^[=+\-@]/.test(x) ? ' ' + x : x; }
  function stamp(iso) { return when(iso); }
  function sheetOf(head, rows, widths) {
    var ws = XLSX.utils.aoa_to_sheet([head].concat(rows.map(function (r) { return r.map(cell); })));
    ws['!cols'] = widths.map(function (w) { return { wch: w }; });
    if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: head.length - 1 } }) };
    return ws;
  }
  function buildBook(old, recs) {
    var wb = XLSX.utils.book_new();
    recs.sort(function (a, b) { return a.scope.localeCompare(b.scope) || b.from.localeCompare(a.from) || b.to.localeCompare(a.to); });
    var head3 = function (r) { return [scopeName(r.scope), dayLabel(r.from), dayLabel(r.to)]; };
    var rca = [], tg = [], sm = [];
    recs.forEach(function (r) {
      Object.keys(r.rca || {}).forEach(function (k) {
        var p = k.split('||'), e = r.rca[k];
        rca.push(head3(r).concat([p[0], p[1] === '__all__' ? 'Whole client' : p[1], p.slice(2).join('||'),
          e.why === 'Other' ? (e.whyOther || 'Other') : (e.why || ''), e.rca || '', e.mitigation || '', e.owner || '', e.due || '',
          e.status === 'Other' ? (e.statusOther || 'Other') : (e.status || 'Not started'), r.by, stamp(r.at)]));
      });
      var t = r.targets || {}, add = function (lvl, who, wk, v) { if (typeof v === 'number' && v > 0) tg.push(head3(r).concat([r.clients.join(', '), lvl, who, wk, v, r.by, stamp(r.at)])); };
      add('Default', 'All billers', '', t.__default__);
      Object.keys(t.byProject || {}).forEach(function (k) { add('Client', k, '', t.byProject[k]); });
      Object.keys(t.byPractice || {}).forEach(function (k) { add('Practice', k, '', t.byPractice[k]); });
      Object.keys(t.byName || {}).forEach(function (k) { add('Biller', (r.names || {})[k] || k, '', t.byName[k]); });
      Object.keys(t.byWeek || {}).forEach(function (wk) {
        var w = t.byWeek[wk] || {};
        Object.keys(w.tl || {}).forEach(function (k) { add('Team, week of', k, dayLabel(wk), w.tl[k]); });
        Object.keys(w.name || {}).forEach(function (k) { add('Biller, week of', (r.names || {})[k] || k, dayLabel(wk), w.name[k]); });
      });
      (r.summary || []).forEach(function (x) {
        sm.push(head3(r).concat([x.client, x.teamLead, x.biller, x.claims, x.target, x.target > 0 ? Math.round(x.claims * 1000 / x.target) / 10 : '',
          x.audited, x.claims > 0 ? Math.round(x.audited * 1000 / x.claims) / 10 : '', x.qa === null ? '' : Math.round(x.qa * 10) / 10, x.errors, x.topError, r.by, stamp(r.at)]));
      });
    });
    XLSX.utils.book_append_sheet(wb, sheetOf(['Scope', 'From', 'To', 'Client', 'Practice', 'Error category', 'Why', 'Root cause', 'Mitigation plan', 'Owner', 'Due date', 'Status', 'Saved by', 'Saved on'], rca, [10, 11, 11, 22, 22, 24, 18, 40, 40, 16, 11, 13, 16, 17]), 'RCA & Mitigation');
    XLSX.utils.book_append_sheet(wb, sheetOf(['Scope', 'From', 'To', 'Client', 'Applies to', 'Name', 'Week of', 'Target / Day', 'Saved by', 'Saved on'], tg, [10, 11, 11, 22, 16, 28, 11, 12, 16, 17]), 'Targets');
    XLSX.utils.book_append_sheet(wb, sheetOf(['Scope', 'From', 'To', 'Client', 'Team Lead', 'Biller', 'Claims Worked', 'Target', 'Attainment %', 'Claims Audited', 'Sampling %', 'QA Score %', 'QA Errors', 'Top error', 'Saved by', 'Saved on'], sm, [10, 11, 11, 22, 18, 22, 13, 9, 13, 14, 11, 11, 10, 24, 16, 17]), 'Period summary');
    // the record itself, split into cells Excel can hold (32,767 characters each)
    var data = [['Scope', 'From', 'To', 'Clients', 'Saved by', 'Saved on', 'Saved work (read by Third Eye, please do not edit)']];
    recs.forEach(function (r) {
      var j = JSON.stringify({ targets: r.targets, keys: r.keys, rca: r.rca, summary: r.summary, names: r.names }), parts = [];
      for (var i = 0; i < j.length; i += 30000) parts.push(j.slice(i, i + 30000));
      data.push([r.scope, r.from, r.to, r.clients.join('\n'), r.by, r.at].concat(parts));
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), DATA_TAB);
    // any tab a supervisor added themselves is carried over untouched
    if (old) old.SheetNames.forEach(function (n) { if (wb.SheetNames.indexOf(n) < 0) XLSX.utils.book_append_sheet(wb, old.Sheets[n], n); });
    wb.Workbook = { Sheets: wb.SheetNames.map(function (n) { return { name: n, Hidden: n === DATA_TAB ? 1 : 0 }; }) };
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }
  function scopeName(code) { return { ar: 'AR', pp: 'Payment Posting', sub: 'Submission' }[code] || code; }

  async function downloadItem(id) {
    var meta = await msx.graph('/me/drive/items/' + encodeURIComponent(id) + '?$select=id,name,parentReference,lastModifiedDateTime,@microsoft.graph.downloadUrl');
    return { meta: meta, buf: await msx.download(meta['@microsoft.graph.downloadUrl']) };
  }
  function folderLabel(meta) {
    var p = meta && meta.parentReference && meta.parentReference.path || '';
    var rest = decodeURIComponent(p.replace(/^.*?root:?/, '')).split('/').filter(Boolean);
    return ['My files'].concat(rest).join(' › ');
  }
  async function saveHere() {
    var who = String(B.who || '').trim();
    if (!who) { say('Type your name first, so others can see who saved this.', 'err'); return draw(); }
    try { localStorage.setItem(NAME_KEY, who); } catch (e) { /* asked again next time */ }
    B.busy = true; say('Saving to ' + B.crumbs[B.crumbs.length - 1].name + '…'); draw();
    try {
      var mine = currentRecords(who), old = null, recs = [];
      var have = B.items.filter(function (x) { return x.name === SAVED_NAME && x.file; })[0];
      if (have) { var got = readRecords((await downloadItem(have.id)).buf); old = got.wb; recs = got.recs; }
      // each client's save replaces its earlier one for the same scope and dates, in this file only
      var key = function (r) { return r.scope + '|' + r.from + '|' + r.to + '|' + r.clients.join('\n'); };
      var now = mine.map(key);
      recs = recs.filter(function (r) { return now.indexOf(key(r)) < 0; }).concat(mine);
      var folder = B.crumbs[B.crumbs.length - 1].id;
      var url = GRAPH + (folder ? '/me/drive/items/' + encodeURIComponent(folder) + ':/' : '/me/drive/root:/') + encodeURIComponent(SAVED_NAME) + ':/content';
      var res = await msx.request(url, true, { method: 'PUT', body: buildBook(old, recs), headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
      var saved = JSON.parse((await res.text()) || '{}');
      if (saved.id) rememberFile(saved.id);
      var where = B.crumbs.map(function (c) { return c.name; }).join(' › ');
      close();
      S.per = []; S.checked = sigOf();
      bar('Saved the targets, RCA and period summary for ' + listOf(mine.map(function (r) { return r.clients[0]; })) + ' (' + scopeLabel() + ', ' + rangeLabel(mine[0]) + ') to ' + where + ' › ' + SAVED_NAME + '.', 'ok');
    } catch (e) { if (B) { B.busy = false; say(msalMessage(e), 'err'); draw(); } }
  }
  function startSave() {
    if (!PROD.length) return bar('Load a production sheet first, then save.', 'err');
    var pr = period();
    open(null, { mode: 'save', title: scopeLabel() + ', ' + rangeLabel(pr), who: myName(),
      start: S.folders.prodInput || S.folders.qaInput || S.lastFolder });
  }

  // every saved work file in Purview Automate's OneDrive, wherever it was saved
  async function findFiles() {
    var by = new Map();
    try {
      var r = await msx.graph("/me/drive/root/search(q='" + encodeURIComponent('Third Eye saved work') + "')?$select=id,name,file&$top=200");
      (r.value || []).forEach(function (x) { if (x.name === SAVED_NAME && x.file) by.set(x.id, 1); });
    } catch (e) { if (e.kind === 'auth' || e.kind === 'signin') throw e; }
    // search can take a few minutes to see a new file, so also look where this visit has been
    knownFiles().forEach(function (id) { by.set(id, 1); });
    var folders = Object.keys(S.folders).map(function (k) { return S.folders[k]; });
    for (var i = 0; i < folders.length; i++) {
      try { (await listFolder(folders[i][folders[i].length - 1].id)).forEach(function (x) { if (x.name === SAVED_NAME && x.file) by.set(x.id, 1); }); }
      catch (e) { /* a folder that has gone is skipped */ }
    }
    return Array.from(by.keys());
  }
  function listOf(a) { return a.length < 2 ? (a[0] || '') : a.slice(0, -1).join(', ') + ' and ' + a[a.length - 1]; }
  async function check(manual) {
    if (!page() || !PROD.length || S.busy) return;
    if (!manual && !(await msx.account().catch(function () { return null; }))) {
      S.per = [];
      return bar('Saved targets and RCA for these dates can load from OneDrive.', '', true);
    }
    var pr = period(), clients = clientsIn(pr), sig = sigOf();
    S.busy = true; bar('Looking in OneDrive for saved work for ' + rangeLabel(pr) + '…');
    try {
      var ids = await findFiles();
      // matched one client at a time: the same dates saved for another client are never loaded here
      var per = clients.map(function (c) { return { client: c, found: [], copyFrom: [], done: null }; });
      for (var i = 0; i < ids.length; i++) {
        var got; try { got = await downloadItem(ids[i]); } catch (e) { if (e.status === 404) continue; throw e; }
        var where = folderLabel(got.meta);
        readRecords(got.buf).recs.forEach(function (r) {
          if (r.scope !== scopeCode()) return;
          per.forEach(function (x) {
            if (r.clients.indexOf(x.client) < 0) return;
            var y = { rec: r, where: where, client: x.client };
            (r.from === pr.from && r.to === pr.to ? x.found : x.copyFrom).push(y);
          });
        });
      }
      var newest = function (a, b) { return b.rec.at.localeCompare(a.rec.at); };
      per.forEach(function (x) {
        x.found.sort(newest);
        // for new dates, the most recent earlier ranges to copy from
        x.copyFrom.sort(function (a, b) { return b.rec.to.localeCompare(a.rec.to) || newest(a, b); });
        x.copyFrom = x.copyFrom.slice(0, 3);
      });
      S.per = per; S.checked = sig;
      per.forEach(function (x) { if (x.found.length === 1) take(x.found[0], false); });
      if (per.some(function (x) { return x.done; })) await settle();
      say2();
    } catch (e) { S.checked = ''; bar(msalMessage(e), 'err'); }
    finally { S.busy = false; drawBar(); }
  }
  // puts one client's saved targets and RCA back, leaving every other client's as they are
  function take(x, copy) {
    var r = x.rec, c = x.client, t = r.targets || {}, keys = r.keys || {};
    ['byName', 'byProject', 'byPractice', 'byWeek', 'notes'].forEach(function (k) { if (!TARGETS[k]) TARGETS[k] = {}; });
    // the client and its practices take the saved figures as a set, so ones cleared before the save stay cleared
    delete TARGETS.byProject[c];
    (keys.practices || []).forEach(function (k) { delete TARGETS.byPractice[k]; });
    Object.keys(t.byProject || {}).forEach(function (k) { if (k === c) TARGETS.byProject[k] = t.byProject[k]; });
    Object.keys(t.byPractice || {}).forEach(function (k) { TARGETS.byPractice[k] = t.byPractice[k]; });
    Object.keys(t.byName || {}).forEach(function (k) { TARGETS.byName[k] = t.byName[k]; });
    Object.keys(t.byWeek || {}).forEach(function (wk) {
      var w = TARGETS.byWeek[wk] = TARGETS.byWeek[wk] || {}; w.tl = w.tl || {}; w.name = w.name || {};
      Object.keys(t.byWeek[wk].tl || {}).forEach(function (k) { w.tl[k] = t.byWeek[wk].tl[k]; });
      Object.keys(t.byWeek[wk].name || {}).forEach(function (k) { w.name[k] = t.byWeek[wk].name[k]; });
    });
    Object.keys(t.notes || {}).forEach(function (k) { if (k.indexOf('default|') !== 0) TARGETS.notes[k] = clone(t.notes[k]); });
    // the default covers every client, so it is only taken as is when this is the only client loaded;
    // otherwise it becomes this client's own figure
    var def = t.__default__ > 0 ? t.__default__ : 0, cur = TARGETS.__default__ || 0;
    if (def && def !== cur && !(t.byProject || {})[c]) {
      if (clientsIn(period()).length === 1) TARGETS.__default__ = def; else TARGETS.byProject[c] = def;
    }
    Object.keys(RCA).forEach(function (k) { if (k.indexOf(c + '||') === 0) delete RCA[k]; });
    Object.keys(r.rca || {}).forEach(function (k) { if (k.indexOf(c + '||') === 0) RCA[k] = clone(r.rca[k]); });
    var x2 = S.per.filter(function (y) { return y.client === c; })[0];
    if (x2) { x2.done = { x: x, copy: copy }; x2.found = []; x2.copyFrom = []; }
  }
  async function settle() {
    try { await sset(K.targets, TARGETS); await sset(K.rca, RCA); } catch (e) { /* on screen either way */ }
    try { if (typeof buildFilters === 'function') buildFilters(); render(); } catch (e) { /* the next redraw shows them */ }
  }
  async function apply(x, copy) { take(x, copy); await settle(); say2(); drawBar(); }
  // one sentence per client: loaded, copied, waiting on a pick, or nothing saved
  function say2() {
    var pr = period(), lines = [], later = [], kind = 'ok';
    S.per.forEach(function (x) {
      if (x.done) {
        var r = x.done.x.rec, src = ' (saved by ' + (r.by || 'someone') + ' on ' + when(r.at) + ', ' + x.done.x.where + ')';
        lines.push(x.done.copy ? 'Copied the targets and RCA for ' + x.client + ' from ' + rangeLabel(r) + src + '. Change what is different, then click Save to OneDrive.'
          : 'Loaded the targets and RCA for ' + x.client + src + '.');
      } else if (x.found.length > 1) { kind = 'warn'; lines.push(x.found.length + ' supervisors saved work for ' + x.client + ' for these dates. Pick whose to load below.'); }
      else { if (kind === 'ok') kind = ''; later.push('Nothing saved yet for ' + x.client + ', ' + rangeLabel(pr) + '.' + (x.copyFrom.length ? ' You can copy an earlier range below.' : '')); }
    });
    bar(lines.concat(later).join('\n'), kind);
  }

  /* the Saved work line under the upload boxes */
  function sigOf() { var pr = period(); return scopeCode() + '|' + pr.from + '|' + pr.to + '|' + clientsIn(pr).join(','); }
  var barState = { text: '', kind: '', signin: false };
  function bar(text, kind, signin) { barState = { text: text || '', kind: kind || '', signin: !!signin }; drawBar(); }
  function drawBar() {
    var host = document.getElementById('odSaved'); if (!host) return;
    var has = page() && PROD.length > 0;
    host.style.display = has ? '' : 'none';
    if (!has) return;
    var choices = [];
    S.per.forEach(function (x, ci) {
      if (x.found.length > 1) x.found.forEach(function (y, i) { choices.push('<button type="button" class="od-choice" data-sv="load" data-c="' + ci + '" data-i="' + i + '"><b>' + esc(x.client) + '</b>: load the one saved by <b>' + esc(y.rec.by || 'someone') + '</b>, ' + esc(when(y.rec.at)) + ', ' + esc(y.where) + '</button>'); });
      else if (!x.done && !x.found.length) x.copyFrom.forEach(function (y, i) { choices.push('<button type="button" class="od-choice" data-sv="copy" data-c="' + ci + '" data-i="' + i + '"><b>' + esc(x.client) + '</b>: copy from ' + esc(rangeLabel(y.rec)) + ' (' + esc(y.rec.by || 'someone') + ', ' + esc(y.where) + ')</button>'); });
    });
    host.innerHTML = '<div class="od-sv"><span class="od-sv-t">Saved work</span><span class="od-sv-msg ' + barState.kind + '">' + esc(barState.text) + '</span>' +
      (barState.signin ? '<button type="button" class="od-btn ghost" data-sv="check"' + (S.busy ? ' disabled' : '') + '>Check OneDrive</button>' : '') +
      '<button type="button" class="od-btn" data-sv="save"' + (S.busy ? ' disabled' : '') + '>Save to OneDrive</button></div>' +
      (choices.length ? '<div class="od-choices">' + choices.join('') + '</div>' : '');
  }
  function onBar(e) {
    var b = e.target.closest('[data-sv]'); if (!b || b.disabled) return;
    var a = b.dataset.sv;
    if (a === 'save') return startSave();
    if (a === 'check') return check(true);
    var x = S.per[+b.dataset.c]; if (!x) return;
    if (a === 'load') return apply(x.found[+b.dataset.i], false);
    if (a === 'copy') return apply(x.copyFrom[+b.dataset.i], true);
  }
  function watch() {
    // the dates or the loaded files changed: look for saved work once things settle
    var t = null;
    setInterval(function () {
      if (!page()) return;
      var sig = PROD.length ? sigOf() : '';
      if (sig === S.sig) return;
      S.sig = sig; drawBar();
      if (!sig || sig === S.checked) return;
      S.per = [];
      clearTimeout(t); t = setTimeout(function () { check(false); }, 1200);
    }, 1000);
  }
  function startSaved() {
    if (!page()) return;
    var after = document.getElementById('loadNote') || document.getElementById('uploadRow'); if (!after) return;
    var host = document.createElement('div'); host.id = 'odSaved'; host.className = 'no-print'; host.style.display = 'none';
    after.insertAdjacentElement('afterend', host);
    host.addEventListener('click', onBar);
    watch();
  }

  /* ---------- styles (the page's own colours, so dark mode follows) ---------- */
  var P = 'var(--panel,var(--surface,#fff))', T = 'var(--text,var(--ink,#111B4D))', D = 'var(--text-dim,var(--muted,#56618C))';
  var L = 'var(--border,var(--line,#DCE3EE))', L2 = 'var(--border-soft,var(--line,#EDF1F7))', BLUE = '#0F6CBD';
  var css = document.createElement('style');
  css.textContent = [
    '.od-up{display:inline-flex;align-items:center;gap:5px;margin-top:6px;font:inherit;font-size:11.5px;font-weight:600;line-height:1.3;color:#fff;background:' + BLUE + ';border:0;border-radius:6px;padding:4px 9px;cursor:pointer;white-space:nowrap;}',
    '.od-up:hover{background:#115EA3;}',
    'label.btn + .od-up{margin:0 0 0 6px;vertical-align:middle;}',
    '.upload-zone{flex-wrap:wrap;}',
    '.upload-zone > span:not(.u-badge){flex:1 1 0;min-width:0;}',   // the text shrinks, so only the button takes a new line
    '.upload-zone .od-row{flex-basis:100%;padding-left:28px;margin-top:-4px;}',
    '.slot .od-up,.drop .od-up{position:relative;z-index:2;align-self:flex-start;}',   // above the box-wide file input
    '.drop .od-up{align-self:center;}',
    '#odPick{position:fixed;inset:0;z-index:100000;background:rgba(10,20,50,.45);display:flex;align-items:flex-start;justify-content:center;padding:8vh 16px 16px;}',
    '#odPick .od-box{background:' + P + ';color:' + T + ';border:1px solid ' + L + ';border-radius:10px;width:min(1040px,100%);height:min(640px,84vh);display:flex;flex-direction:column;box-shadow:0 20px 50px rgba(0,0,0,.3);font:14px/1.4 "Segoe UI",system-ui,sans-serif;overflow:hidden;}',
    '#odPick .od-hd{display:flex;align-items:center;gap:14px;padding:10px 16px;background:' + BLUE + ';color:#fff;}',
    '#odPick .od-brand{font-size:16px;font-weight:600;}',
    '#odPick .od-who{flex:1;font-size:12.5px;opacity:.9;}',
    '#odPick .od-x{font:inherit;font-size:24px;line-height:1;border:0;background:none;color:#fff;cursor:pointer;padding:0 4px;}',
    '#odPick .od-for{padding:9px 24px;font-size:12.5px;color:' + D + ';border-bottom:1px solid ' + L2 + ';}',
    '#odPick .od-dim{color:' + D + ';font-size:12.5px;}',
    '#odPick .od-crumbs{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:16px 24px 10px;font-size:20px;}',
    '#odPick .od-here{font-weight:600;}',
    '#odPick .od-crumb{font:inherit;border:0;background:none;color:' + T + ';cursor:pointer;padding:0;}',
    '#odPick .od-crumbs .od-crumb:hover{color:' + BLUE + ';text-decoration:underline;}',
    '#odPick .od-ft .od-crumb{color:' + BLUE + ';font-size:12.5px;}',
    '#odPick .od-sep{color:' + D + ';font-size:16px;}',
    '#odPick .od-filter{margin-left:auto;font:inherit;font-size:13px;padding:6px 12px;border:1px solid ' + L + ';border-radius:16px;background:' + P + ';color:' + T + ';min-width:220px;}',
    '#odPick .od-list{overflow:auto;flex:1;margin:0 16px;border:1px solid ' + L2 + ';border-radius:8px;}',
    '#odPick table{width:100%;border-collapse:collapse;}',
    '#odPick th{position:sticky;top:0;background:' + P + ';text-align:left;font-weight:600;padding:12px 14px;border-bottom:1px solid ' + L + ';white-space:nowrap;}',
    '#odPick td{padding:11px 14px;border-bottom:1px solid ' + L2 + ';cursor:pointer;white-space:nowrap;color:' + D + ';}',
    '#odPick td.od-name{white-space:normal;word-break:break-word;color:' + T + ';padding-left:20px;}',
    '#odPick .od-nm{display:flex;align-items:center;gap:12px;}',
    '#odPick th.od-name{padding-left:56px;width:44%;}',
    '#odPick tr:hover td{background:rgba(15,108,189,.07);}',
    '#odPick tr.sel td{background:rgba(15,108,189,.16);}',
    '#odPick .od-ico{flex:none;}',
    '#odPick .od-empty{padding:22px 18px;color:' + D + ';}',
    '#odPick .od-empty p{margin:0 0 12px;color:' + T + ';}',
    '#odPick .od-msg{padding:0 18px;min-height:0;color:' + D + ';}',
    '#odPick .od-msg:not(:empty){padding:8px 18px 0;}',
    '#odPick .od-msg.err{color:var(--bad,var(--red,#B42318));}',
    '#odPick .od-msg.ok{color:var(--good,#127A4D);}',
    '#odPick .od-ft > span:last-child{margin-left:auto;white-space:nowrap;}',
    '#odPick .od-ft{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;padding:12px 18px;}',
    '#odPick .od-btn{font:inherit;font-weight:600;font-size:13px;padding:7px 16px;border:0;border-radius:8px;cursor:pointer;color:#fff;background:' + BLUE + ';margin-left:8px;}',
    '#odPick .od-empty .od-btn{margin-left:0;}',
    '#odPick .od-btn.ghost{background:transparent;color:' + T + ';border:1px solid ' + L + ';}',
    '#odPick .od-btn[disabled]{opacity:.5;cursor:default;}',
    '#odSaved{margin:-4px 0 14px;}',
    '#odSaved .od-sv{display:flex;align-items:center;gap:10px;flex-wrap:wrap;background:' + P + ';border:1px solid ' + L + ';border-radius:10px;padding:9px 14px;font-size:13px;}',
    '#odSaved .od-sv-t{font-weight:700;color:' + T + ';}',
    '#odSaved .od-sv-msg{flex:1;min-width:220px;color:' + D + ';white-space:pre-line;}',   // one line per client
    '#odSaved .od-sv-msg.ok{color:var(--good,#127A4D);}',
    '#odSaved .od-sv-msg.warn{color:var(--warn,#935300);}',
    '#odSaved .od-sv-msg.err{color:var(--bad,#B42318);}',
    '#odSaved .od-btn{font:inherit;font-weight:600;font-size:12.5px;padding:6px 13px;border:0;border-radius:8px;cursor:pointer;color:#fff;background:' + BLUE + ';}',
    '#odSaved .od-btn.ghost{background:transparent;color:' + BLUE + ';border:1px solid ' + BLUE + ';}',
    '#odSaved .od-btn[disabled]{opacity:.5;cursor:default;}',
    '#odSaved .od-choices{display:flex;flex-direction:column;gap:6px;margin-top:6px;}',
    '#odSaved .od-choice{font:inherit;font-size:12.5px;text-align:left;padding:7px 12px;border:1px solid ' + L + ';border-radius:8px;background:' + P + ';color:' + T + ';cursor:pointer;}',
    '#odSaved .od-choice:hover{border-color:' + BLUE + ';}',
    '#odPick .od-name-in{display:inline-flex;align-items:center;gap:6px;margin-right:10px;color:' + T + ';}',
    '#odPick .od-name-in input{font:inherit;font-size:13px;padding:5px 9px;border:1px solid ' + L + ';border-radius:7px;background:' + P + ';color:' + T + ';width:170px;}',
    '@media print{.od-up,#odPick,#odSaved{display:none !important;}}',
  ].join('\n');

  function start() {
    if (!supported()) return;
    (document.head || document.documentElement).appendChild(css);
    addButtons(document);
    startSaved();
    // Client Reporting draws its upload boxes later, so new ones get the button as they appear
    new MutationObserver(function (list) {
      list.forEach(function (m) { m.addedNodes.forEach(function (n) { if (n.nodeType === 1) addButtons(n); }); });
    }).observe(document.body, { childList: true, subtree: true });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  window.pvOneDrive = { open: open };
})();
