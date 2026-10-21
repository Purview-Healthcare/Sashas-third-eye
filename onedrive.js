/* Upload from OneDrive, for Sasha's Third Eye (Internal and Client Reporting).
 *
 * Every Excel upload box gets an "Upload from OneDrive" button. It opens a small file browser on the
 * OneDrive of Purviewautomate@purview-healthcare.com only, where supervisors keep their folders.
 * The picked file is loaded into that box exactly as if it had been chosen from the computer.
 *
 * The Microsoft sign-in only accepts the Purviewautomate account: any other account that signs in
 * is forgotten again at once, and the person's own OneDrive is never asked for. Third Eye only reads
 * files; nothing is ever written to OneDrive.
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
    request: async function (url, auth) {
      var m = mock(), doFetch = (m && m.fetch) || function (u, o) { return fetch(u, o); };
      for (var attempt = 0; ; attempt++) {
        var headers = {};
        if (auth !== false) headers.Authorization = 'Bearer ' + (await msx.token());
        var r;
        try { r = await doFetch(url, { method: 'GET', headers: headers }); }
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
  function fits(name) { var n = String(name).toLowerCase(); return B.exts.some(function (x) { return n.slice(-x.length) === x; }); }
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
    } catch (e) { say(msalMessage(e), 'err');
    }
    B.busy = false; draw();
  }
  async function pick() {
    var it = B.picked && B.items.filter(function (x) { return x.id === B.picked; })[0];
    if (!it) return;
    B.busy = true; say('Downloading ' + it.name + '…'); draw();
    try {
      var meta = await msx.graph('/me/drive/items/' + encodeURIComponent(it.id) + '?$select=id,name,@microsoft.graph.downloadUrl');
      var data = await msx.download(meta['@microsoft.graph.downloadUrl']);
      var file = new File([data], it.name, { type: it.file && it.file.mimeType || '', lastModified: Date.parse(it.lastModifiedDateTime) || Date.now() });
      var input = B.input;
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
        }).join('') + '</tbody></table>' : '</table><div class="od-empty">' + (B.busy ? 'Loading…' : q ? 'Nothing in this folder matches.' : 'This folder has no files that can be uploaded here.') + '</div>') + '</div>';
    }
    host.innerHTML = '<div class="od-box" role="dialog" aria-modal="true" aria-label="OneDrive">' +
      '<div class="od-hd"><span class="od-brand">OneDrive</span><span class="od-who">' + esc(MS.owner) + '</span>' +
      '<button type="button" class="od-x" data-od="close" aria-label="Close">×</button></div>' +
      (B.title ? '<div class="od-for">Uploading to <b>' + esc(B.title) + '</b>. Open the folder, click the file, then click Upload.</div>' : '') + body +
      '<div class="od-msg ' + B.msgKind + '">' + esc(B.msg) + '</div>' +
      '<div class="od-ft"><span class="od-dim">' + (B.signedIn ? 'Only ' + esc(B.exts.join(', ')) + ' files are listed' + (hidden ? ' (' + hidden + ' other file' + (hidden === 1 ? '' : 's') + ' hidden)' : '') + '. ' +
        '<button type="button" class="od-crumb" data-od="signout">Sign out of OneDrive</button>' : '') + '</span>' +
      '<span><button type="button" class="od-btn ghost" data-od="close">Cancel</button>' +
      (B.signedIn ? '<button type="button" class="od-btn" data-od="pick"' + (B.picked && !B.busy ? '' : ' disabled') + '>Upload</button>' : '') + '</span></div></div>';
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
    await openFolder([{ id: '', name: 'My files' }]);   // always starts at My files, as in OneDrive
  }
  async function open(input) {
    close();
    B = { input: input, title: title(input), exts: exts(input), crumbs: [], items: [], filter: '', picked: null, busy: true, signedIn: false, msg: '', msgKind: '' };
    var host = document.createElement('div'); host.id = 'odPick';
    document.body.appendChild(host);
    host.addEventListener('click', onClick);
    host.addEventListener('dblclick', onDbl);
    host.addEventListener('input', function (e) { if (e.target.id === 'odFilter') { B.filter = e.target.value; B.filterFocus = true; draw(); B.filterFocus = false; } });
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
  function onDbl(e) { var r = e.target.closest('[data-od="select"]'); if (r && B && !B.busy) { B.picked = r.dataset.id; pick(); } }

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
    '@media print{.od-up,#odPick{display:none !important;}}',
  ].join('\n');

  function start() {
    if (!supported()) return;
    (document.head || document.documentElement).appendChild(css);
    addButtons(document);
    // Client Reporting draws its upload boxes later, so new ones get the button as they appear
    new MutationObserver(function (list) {
      list.forEach(function (m) { m.addedNodes.forEach(function (n) { if (n.nodeType === 1) addButtons(n); }); });
    }).observe(document.body, { childList: true, subtree: true });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  window.pvOneDrive = { open: open };
})();
