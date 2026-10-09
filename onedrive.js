/* Push to OneDrive / Read from OneDrive for Sasha's Third Eye (Internal and Client Reporting).
 *
 * Every file picked or dropped into an upload box on the page is remembered for this visit.
 * "Push to OneDrive" signs the user in with their own Microsoft work account and saves those
 * files, plus a small manifest.json, to their OneDrive:
 *     Sashas Third Eye / <Internal Reporting | Client Reporting> / <client or name> / <date time> /
 * "Read from OneDrive" lists the saved sets, downloads the chosen one and puts each file back
 * into the same upload box, exactly as if it had been picked again.
 *
 * Setup: a Microsoft 365 admin registers this site in Microsoft Entra ID (single-page app,
 * redirect URI = <site>/auth-redirect.html, delegated permissions User.Read + Files.ReadWrite)
 * and the two IDs below are filled in. Neither ID is a secret.
 */
(function () {
  var CONFIG = {
    clientId: '',   // Application (client) ID from the Entra app registration
    tenantId: '',   // Directory (tenant) ID, so only your company's accounts can sign in
    scopes: ['User.Read', 'Files.ReadWrite'],
    root: 'Sashas Third Eye',
  };
  if (window.PV_ONEDRIVE_CONFIG) Object.assign(CONFIG, window.PV_ONEDRIVE_CONFIG);
  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var AREA = /client\.html$/i.test(location.pathname) ? 'Client Reporting' : 'Internal Reporting';
  var IS_CLIENT = AREA === 'Client Reporting';

  /* ---------- remember what was uploaded ---------- */
  var files = {};        // key -> File, in the order of the upload boxes on the page
  function keyOf(input) {
    if (!input || input.type !== 'file' || input.id === 'importInput') return null;
    var slot = input.closest('[data-slot]');
    return slot ? 'slot:' + slot.dataset.slot : input.id || null;
  }
  document.addEventListener('change', function (e) {
    var k = keyOf(e.target); if (k && e.target.files && e.target.files[0] && !restoring) files[k] = e.target.files[0];
  }, true);
  document.addEventListener('drop', function (e) {
    var box = e.target.closest && e.target.closest('label, .drop, [data-slot]'); if (!box || !e.dataTransfer || !e.dataTransfer.files[0]) return;
    var k = keyOf(box.querySelector('input[type=file]')); if (k) files[k] = e.dataTransfer.files[0];
  }, true);
  document.addEventListener('click', function (e) {
    var c = e.target.closest && e.target.closest('[data-clear]'); if (!c) return;
    var which = c.dataset.clear;
    if (which === 'all') files = {};
    else ({ qa: ['qaInput', 'baseQaInput'], prod: ['prodInput', 'baseProdInput'], roster: ['rosterInput'] }[which] || []).forEach(function (k) { delete files[k]; });
  }, true);
  var restoring = false;

  /* ---------- Microsoft sign-in ---------- */
  var msalApp = null, account = null;
  function loadScript(src) { return new Promise(function (res, rej) { var s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = function () { rej(new Error('Could not load ' + src)); }; document.head.appendChild(s); }); }
  function baseUrl() { return location.href.replace(/[?#].*$/, '').replace(/[^/]*$/, ''); }
  async function getToken() {
    if (!CONFIG.clientId || !CONFIG.tenantId) throw new Error('NOT_CONFIGURED');
    if (!msalApp) {
      if (!window.msal) await loadScript(baseUrl() + 'vendor/msal-browser-3.30.0.min.js');
      msalApp = new msal.PublicClientApplication({
        auth: { clientId: CONFIG.clientId, authority: 'https://login.microsoftonline.com/' + CONFIG.tenantId, redirectUri: baseUrl() + 'auth-redirect.html' },
        cache: { cacheLocation: 'localStorage' },
      });
      await msalApp.initialize();
      account = msalApp.getAllAccounts()[0] || null;
    }
    var req = { scopes: CONFIG.scopes, account: account };
    try {
      if (!account) throw new Error('login');
      return (await msalApp.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      var r = account ? await msalApp.acquireTokenPopup(req) : await msalApp.loginPopup({ scopes: CONFIG.scopes, prompt: 'select_account' });
      account = r.account; msalApp.setActiveAccount(account);
      return r.accessToken || (await msalApp.acquireTokenSilent({ scopes: CONFIG.scopes, account: account })).accessToken;
    }
  }

  /* ---------- OneDrive (Microsoft Graph) ---------- */
  function clean(s) { return String(s || '').replace(/[\\/:*?"<>|#%]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Untitled'; }
  function pathUrl(parts) { return GRAPH + '/me/drive/root:/' + parts.map(encodeURIComponent).join('/'); }
  async function graph(token, url, opts) {
    opts = opts || {}; opts.headers = Object.assign({ Authorization: 'Bearer ' + token }, opts.headers || {});
    var r = await fetch(url, opts);
    if (r.status === 404) return null;
    if (!r.ok) { var t = ''; try { t = (await r.json()).error.message; } catch (e) {} throw new Error('OneDrive said: ' + (t || r.status + ' ' + r.statusText)); }
    return r.status === 204 ? {} : r.json();
  }
  async function children(token, parts) {
    var out = [], url = pathUrl(parts) + ':/children?$top=200&select=id,name,folder,file,size,lastModifiedDateTime,@microsoft.graph.downloadUrl';
    while (url) { var j = await graph(token, url); if (!j) return out; out = out.concat(j.value || []); url = j['@odata.nextLink'] || null; }
    return out;
  }
  async function upload(token, parts, blob) {
    if (blob.size <= 4 * 1024 * 1024) return graph(token, pathUrl(parts) + ':/content', { method: 'PUT', body: blob });
    var s = await graph(token, pathUrl(parts) + ':/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) });
    var CHUNK = 327680 * 16, last = null;
    for (var i = 0; i < blob.size; i += CHUNK) {
      var end = Math.min(i + CHUNK, blob.size);
      var r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': 'bytes ' + i + '-' + (end - 1) + '/' + blob.size }, body: blob.slice(i, end) });
      if (!r.ok) throw new Error('OneDrive upload failed (' + r.status + ')');
      last = r;
    }
    return last.json();
  }

  /* ---------- page specifics ---------- */
  function fieldVal(id) { var el = document.getElementById(id); return el ? el.value : ''; }
  function setName() {
    if (IS_CLIENT) { var c = fieldVal('upClient'); if (c) return c; var n = document.querySelector('.client-name'); return n ? n.textContent : ''; }
    return '';
  }
  function stamp(d) { function p(n) { return String(n).padStart(2, '0'); } return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + p(d.getMinutes()); }
  function labelFor(key) {
    if (key.indexOf('slot:') === 0) { var el = document.querySelector('[data-slot="' + key.slice(5) + '"] .ttl'); return el ? el.textContent : key.slice(5); }
    var inp = document.getElementById(key); var t = inp && inp.closest('label') && inp.closest('label').querySelector('.u-title');
    return t ? t.textContent : key;
  }
  /** Make sure the upload boxes are on screen (Client Reporting shows them only before a report is built). */
  async function showUploadBoxes() {
    if (!IS_CLIENT || document.querySelector('[data-slot] input[type=file]')) return;
    var b = document.getElementById('btnUpload'); if (b) { b.click(); await wait(300); }
  }
  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  /* ---------- dialogs ---------- */
  function modal(title, bodyHtml) {
    close();
    var m = document.createElement('div'); m.id = 'odModal';
    m.innerHTML = '<div class="od-card" role="dialog" aria-label="' + esc(title) + '"><div class="od-head"><h2>' + esc(title) + '</h2><button type="button" class="od-x" aria-label="Close">×</button></div><div class="od-body">' + bodyHtml + '</div></div>';
    document.body.appendChild(m);
    m.querySelector('.od-x').onclick = close;
    m.addEventListener('click', function (e) { if (e.target === m) close(); });
    return m;
  }
  function close() { var m = document.getElementById('odModal'); if (m) m.remove(); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmtSize(n) { return n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }
  function problem(e) {
    if (e && e.message === 'NOT_CONFIGURED') return 'OneDrive isn’t connected yet. Your Microsoft 365 admin needs to register this site once (steps are in ONEDRIVE-SETUP.md), then the two IDs go into onedrive.js.';
    if (e && /user_cancelled|popup_window_error|cancel/i.test(e.errorCode || e.message || '')) return 'Microsoft sign-in was closed or blocked. Allow pop-ups for this site and try again.';
    return (e && e.message) || String(e);
  }

  async function onPush() {
    var keys = Object.keys(files);
    if (!keys.length) { modal('Push to OneDrive', '<p>Nothing to push yet. Upload your sheets on this page first, then push them.</p>'); return; }
    var name = setName();
    var m = modal('Push to OneDrive', '<p>These files go to your OneDrive in <b>' + esc(CONFIG.root + ' / ' + AREA) + '</b>.</p>' +
      '<label class="od-field">' + (IS_CLIENT ? 'Client' : 'Name for this upload (client, team or week)') + '<input id="odName" value="' + esc(name) + '" placeholder="' + (IS_CLIENT ? 'Client name' : 'For example: Week 40') + '"></label>' +
      '<ul class="od-list">' + keys.map(function (k) { return '<li><b>' + esc(labelFor(k)) + '</b><span>' + esc(files[k].name) + ' · ' + fmtSize(files[k].size) + '</span></li>'; }).join('') + '</ul>' +
      '<div class="od-msg" id="odMsg"></div><div class="od-foot"><button type="button" class="od-btn ghost" id="odCancel">Cancel</button><button type="button" class="od-btn" id="odGo">Push to OneDrive</button></div>');
    m.querySelector('#odCancel').onclick = close;
    m.querySelector('#odGo').onclick = async function () {
      var go = this, msg = m.querySelector('#odMsg'); go.disabled = true;
      var who = clean(m.querySelector('#odName').value || (IS_CLIENT ? 'Client' : 'Internal'));
      try {
        msg.textContent = 'Signing in to Microsoft…';
        var token = await getToken();
        var folder = [CONFIG.root, AREA, who, stamp(new Date())];
        var manifest = { app: "Sasha's Third Eye", area: AREA, name: who, savedAt: new Date().toISOString(), savedBy: account ? account.username : '', fields: IS_CLIENT ? { upClient: fieldVal('upClient') || who, upType: fieldVal('upType') } : {}, files: [] };
        for (var i = 0; i < keys.length; i++) {
          var f = files[keys[i]], fname = (i + 1) + ' ' + clean(labelFor(keys[i])).slice(0, 40) + ' - ' + clean(f.name.replace(/\.[^.]+$/, '')) + (f.name.match(/\.[^.]+$/) || [''])[0];
          msg.textContent = 'Uploading ' + (i + 1) + ' of ' + keys.length + ': ' + f.name;
          await upload(token, folder.concat(fname), f);
          manifest.files.push({ key: keys[i], label: labelFor(keys[i]), originalName: f.name, storedAs: fname, size: f.size });
        }
        await upload(token, folder.concat('manifest.json'), new Blob([JSON.stringify(manifest, null, 2)], { type: 'application/json' }));
        m.querySelector('.od-body').innerHTML = '<p class="od-ok">Saved ' + keys.length + ' file' + (keys.length === 1 ? '' : 's') + ' to OneDrive.</p><p class="od-path">' + esc(folder.join(' / ')) + '</p><div class="od-foot"><button type="button" class="od-btn" id="odDone">Done</button></div>';
        m.querySelector('#odDone').onclick = close;
      } catch (e) { msg.className = 'od-msg err'; msg.textContent = problem(e); go.disabled = false; }
    };
  }

  async function onRead() {
    var m = modal('Read from OneDrive', '<div class="od-msg" id="odMsg">Signing in to Microsoft…</div>');
    var msg = m.querySelector('#odMsg');
    try {
      var token = await getToken();
      msg.textContent = 'Looking for saved uploads…';
      var groups = (await children(token, [CONFIG.root, AREA])).filter(function (x) { return x.folder; });
      var sets = [];
      await Promise.all(groups.map(async function (g) {
        (await children(token, [CONFIG.root, AREA, g.name])).filter(function (x) { return x.folder; }).forEach(function (s) { sets.push({ group: g.name, name: s.name, count: s.folder.childCount, when: s.lastModifiedDateTime }); });
      }));
      sets.sort(function (a, b) { return a.when < b.when ? 1 : -1; });
      if (!document.getElementById('odModal')) return;
      if (!sets.length) { m.querySelector('.od-body').innerHTML = '<p>No saved uploads yet in <b>' + esc(CONFIG.root + ' / ' + AREA) + '</b>. Use Push to OneDrive after uploading.</p>'; return; }
      var groupsSeen = sets.map(function (s) { return s.group; }).filter(function (g, i, a) { return a.indexOf(g) === i; });
      m.querySelector('.od-body').innerHTML = (groupsSeen.length > 1 ? '<input class="od-search" id="odSearch" placeholder="Search ' + (IS_CLIENT ? 'clients' : 'names') + '">' : '') +
        '<ul class="od-list pick">' + sets.map(function (s, i) { return '<li data-i="' + i + '"><b>' + esc(s.group) + '</b><span>' + esc(s.name) + ' · ' + Math.max(0, s.count - 1) + ' file' + (s.count === 2 ? '' : 's') + '</span><button type="button" class="od-btn sm" data-load="' + i + '">Load</button></li>'; }).join('') + '</ul><div class="od-msg" id="odMsg"></div>';
      msg = m.querySelector('#odMsg');
      var search = m.querySelector('#odSearch');
      if (search) search.oninput = function () { var q = search.value.toLowerCase(); m.querySelectorAll('.od-list li').forEach(function (li) { li.style.display = li.textContent.toLowerCase().indexOf(q) >= 0 ? '' : 'none'; }); };
      m.querySelectorAll('[data-load]').forEach(function (b) { b.onclick = function () { loadSet(token, sets[+b.dataset.load], m, msg); }; });
    } catch (e) { msg.className = 'od-msg err'; msg.textContent = problem(e); }
  }

  async function loadSet(token, set, m, msg) {
    m.querySelectorAll('[data-load]').forEach(function (b) { b.disabled = true; });
    try {
      msg.className = 'od-msg'; msg.textContent = 'Downloading…';
      var items = await children(token, [CONFIG.root, AREA, set.group, set.name]);
      var man = items.find(function (x) { return x.name === 'manifest.json'; });
      if (!man) throw new Error('This folder has no manifest.json, so it was not saved by this app.');
      var manifest = await (await fetch(man['@microsoft.graph.downloadUrl'])).json();
      await showUploadBoxes();
      Object.keys(manifest.fields || {}).forEach(function (id) { var el = document.getElementById(id); if (el && manifest.fields[id]) { el.value = manifest.fields[id]; el.dispatchEvent(new Event('change', { bubbles: true })); } });
      files = {};
      var missing = [];
      for (var i = 0; i < manifest.files.length; i++) {
        var f = manifest.files[i], item = items.find(function (x) { return x.name === f.storedAs; });
        var input = f.key.indexOf('slot:') === 0 ? document.querySelector('[data-slot="' + f.key.slice(5) + '"] input[type=file]') : document.getElementById(f.key);
        if (!item || !input) { missing.push(f.label); continue; }
        msg.textContent = 'Loading ' + (i + 1) + ' of ' + manifest.files.length + ': ' + f.originalName;
        var blob = await (await fetch(item['@microsoft.graph.downloadUrl'])).blob();
        var file = new File([blob], f.originalName, { type: blob.type });
        var dt = new DataTransfer(); dt.items.add(file);
        restoring = true; input.files = dt.files; input.dispatchEvent(new Event('change', { bubbles: true })); restoring = false;
        files[f.key] = file;
        await wait(250);
      }
      close();
      toast('Loaded ' + (manifest.files.length - missing.length) + ' file' + (manifest.files.length - missing.length === 1 ? '' : 's') + ' from OneDrive: ' + set.group + ', ' + set.name + (IS_CLIENT ? '. Check the sheets, then click Build dashboards.' : '.') + (missing.length ? ' Not loaded: ' + missing.join(', ') + '.' : ''));
    } catch (e) { restoring = false; msg.className = 'od-msg err'; msg.textContent = problem(e); m.querySelectorAll('[data-load]').forEach(function (b) { b.disabled = false; }); }
  }

  function toast(text) {
    var t = document.createElement('div'); t.className = 'od-toast'; t.textContent = text; document.body.appendChild(t);
    setTimeout(function () { t.remove(); }, 7000);
  }

  /* ---------- buttons and styles ---------- */
  var css = document.createElement('style');
  css.textContent = [
    '.od-bar{display:inline-flex;gap:6px;margin-left:6px;vertical-align:middle;flex-wrap:wrap;}',
    '.od-bar button{display:inline-flex;align-items:center;gap:6px;font:600 12px/1 "Poppins",system-ui,sans-serif;padding:7px 12px;border-radius:999px;border:1px solid rgba(32,198,207,.55);background:rgba(32,198,207,.08);color:inherit;cursor:pointer;}',
    '.od-bar button:hover{background:rgba(32,198,207,.18);}',
    '.od-bar svg{width:14px;height:14px;}',
    '#odModal{position:fixed;inset:0;z-index:9000;background:rgba(10,15,44,.45);display:flex;align-items:flex-start;justify-content:center;padding:60px 16px;overflow:auto;font-family:"Poppins",system-ui,-apple-system,"Segoe UI",sans-serif;}',
    '#odModal .od-card{background:#fff;color:#111B4D;width:min(560px,100%);border-radius:16px;box-shadow:0 30px 80px rgba(10,15,44,.35);text-align:left;}',
    '#odModal .od-head{display:flex;align-items:center;padding:16px 20px;border-bottom:1px solid #E6EAF0;}',
    '#odModal h2{margin:0;font-size:17px;font-weight:600;color:#16156A;flex:1;}',
    '#odModal .od-x{border:0;background:transparent;font-size:22px;line-height:1;color:#8A94A6;cursor:pointer;width:32px;height:32px;border-radius:8px;}',
    '#odModal .od-body{padding:16px 20px 18px;font-size:13.5px;}',
    '#odModal p{margin:0 0 12px;color:#3D4874;}',
    '#odModal .od-field{display:flex;flex-direction:column;gap:5px;font-size:12.5px;font-weight:500;color:#3D4874;margin-bottom:12px;}',
    '#odModal input{font:inherit;font-size:14px;padding:9px 12px;border:1px solid #DCE3EE;border-radius:10px;color:#111B4D;}',
    '#odModal .od-search{width:100%;margin-bottom:10px;}',
    '#odModal .od-list{list-style:none;margin:0 0 12px;padding:0;border:1px solid #E6EAF0;border-radius:12px;max-height:340px;overflow:auto;}',
    '#odModal .od-list li{display:flex;flex-wrap:wrap;align-items:center;gap:2px 10px;padding:9px 12px;border-top:1px solid #E6EAF0;}',
    '#odModal .od-list li:first-child{border-top:0;}',
    '#odModal .od-list b{font-weight:600;color:#16156A;}',
    '#odModal .od-list span{color:#6A7499;font-size:12.5px;flex:1;min-width:160px;}',
    '#odModal .od-foot{display:flex;justify-content:flex-end;gap:8px;margin-top:6px;}',
    '#odModal .od-btn{font:inherit;font-weight:600;font-size:13px;padding:9px 16px;border:0;border-radius:10px;cursor:pointer;color:#08163F;background:linear-gradient(120deg,#27DA97 0%,#23D0B4 50%,#20C6CF 100%);}',
    '#odModal .od-btn.sm{padding:6px 12px;font-size:12px;}',
    '#odModal .od-btn.ghost{background:#EDF1F7;color:#3D4874;}',
    '#odModal .od-btn[disabled]{opacity:.55;cursor:default;}',
    '#odModal .od-msg{min-height:18px;font-size:12.5px;color:#56618C;}',
    '#odModal .od-msg.err{color:#BA2340;}',
    '#odModal .od-ok{font-weight:600;color:#127A4D;}',
    '#odModal .od-path{font-size:12.5px;color:#56618C;word-break:break-word;}',
    '.od-toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:9001;max-width:min(560px,calc(100% - 32px));background:#16156A;color:#fff;padding:12px 16px;border-radius:12px;font:500 13px/1.45 "Poppins",system-ui,sans-serif;box-shadow:0 12px 30px rgba(10,15,44,.3);}',
    '@media print{.od-bar,#odModal,.od-toast{display:none !important;}}',
  ].join('\n');
  (document.head || document.documentElement).appendChild(css);

  var UP = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 18a4.5 4.5 0 0 1-.5-9A6 6 0 0 1 18 9.5a4 4 0 0 1-.5 8.5"/><path d="M12 12v8M9 15l3-3 3 3"/></svg>';
  var DOWN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 18a4.5 4.5 0 0 1-.5-9A6 6 0 0 1 18 9.5a4 4 0 0 1-.5 8.5"/><path d="M12 12v8M9 17l3 3 3-3"/></svg>';
  document.addEventListener('DOMContentLoaded', function () {
    var nav = document.querySelector('.rep-tabs'); if (!nav) return;
    var bar = document.createElement('span'); bar.className = 'od-bar';
    bar.innerHTML = '<button type="button" id="odPush" title="Save everything uploaded on this page to your OneDrive">' + UP + 'Push to OneDrive</button><button type="button" id="odRead" title="Load a saved upload from your OneDrive">' + DOWN + 'Read from OneDrive</button>';
    nav.insertAdjacentElement('afterend', bar);
    bar.querySelector('#odPush').onclick = onPush;
    bar.querySelector('#odRead').onclick = onRead;
  });

  window.pvOneDrive = { reset: function () { files = {}; }, files: function () { return files; } };
})();
