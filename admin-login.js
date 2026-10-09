/* Admin login for Sasha's Third Eye (Internal Reporting and Client Reporting).
 *
 * Loaded in the <head> of index.html and client.html. Until the admin signs in, the page
 * content is hidden and a sign-in card is shown. A sign-in lasts 12 hours in that browser,
 * or until "Sign out".
 *
 * The site is static (GitHub Pages), so this is a front-door lock, not server security:
 * uploaded sheets never leave the browser either way. Only a salted PBKDF2 hash of the
 * credentials is stored here. To change the password, generate a new SALT/HASH pair:
 *   python3 -c "import hashlib,secrets;s=secrets.token_hex(16);print(s, hashlib.pbkdf2_hmac('sha256',b'admin:NEW-PASSWORD',s.encode(),150000).hex())"
 */
(function () {
  var ADMIN = {
    user: 'admin',
    salt: '798b0f4febfe5e8b3a76b02471fe0d2a',
    iterations: 150000,
    hash: '85b0202a3291b954644999ff25bba7f7ab4a1dfa9e58bfd73940efb3f67babaa',
  };
  var KEY = 'pv.adminSession';
  var HOURS = 12;

  function readSession() {
    try { var s = JSON.parse(localStorage.getItem(KEY) || 'null'); return s && s.user === ADMIN.user && s.exp > Date.now() ? s : null; } catch (e) { return null; }
  }
  function writeSession() { try { localStorage.setItem(KEY, JSON.stringify({ user: ADMIN.user, exp: Date.now() + HOURS * 3600000 })); } catch (e) {} }
  function clearSession() { try { localStorage.removeItem(KEY); } catch (e) {} }

  async function hashOf(user, pass) {
    var enc = new TextEncoder();
    var key = await crypto.subtle.importKey('raw', enc.encode(user + ':' + pass), 'PBKDF2', false, ['deriveBits']);
    var bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: enc.encode(ADMIN.salt), iterations: ADMIN.iterations, hash: 'SHA-256' }, key, 256);
    return Array.from(new Uint8Array(bits)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
  }

  var root = document.documentElement;
  var signedIn = !!readSession();
  if (!signedIn) root.classList.add('pv-locked');

  var css = document.createElement('style');
  css.textContent = [
    'html.pv-locked body > *:not(#pvLogin){display:none !important;}',
    '#pvLogin{position:fixed;inset:0;z-index:9999;display:grid;place-items:center;padding:16px;background:radial-gradient(110% 75% at 4% -10%,rgba(45,224,157,.18) 0%,rgba(45,224,157,0) 48%),radial-gradient(110% 75% at 100% -4%,rgba(75,63,209,.13) 0%,rgba(75,63,209,0) 46%),#F3F6FA;font-family:"Poppins",system-ui,-apple-system,"Segoe UI",sans-serif;}',
    '#pvLogin .card{width:min(400px,100%);background:#fff;border:1px solid #DCE3EE;border-radius:18px;padding:30px 30px 26px;box-shadow:0 20px 60px rgba(17,27,77,.12);text-align:center;}',
    '#pvLogin svg{width:56px;height:56px;display:block;margin:0 auto 8px;}',
    '#pvLogin h1{margin:0;font-size:26px;line-height:1.15;font-weight:700;letter-spacing:-.02em;color:#16156A;}',
    '#pvLogin .org{margin:6px 0 20px;font-size:11px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:#6A7499;}',
    '#pvLogin form{display:flex;flex-direction:column;gap:12px;text-align:left;}',
    '#pvLogin label{display:flex;flex-direction:column;gap:5px;font-size:12.5px;font-weight:500;color:#3D4874;}',
    '#pvLogin input{font:inherit;font-size:14px;padding:10px 12px;border:1px solid #DCE3EE;border-radius:10px;color:#111B4D;outline:none;}',
    '#pvLogin input:focus{border-color:#20C6CF;box-shadow:0 0 0 3px rgba(32,198,207,.18);}',
    '#pvLogin button{margin-top:4px;font:inherit;font-weight:600;font-size:14px;padding:11px 14px;border:0;border-radius:10px;cursor:pointer;color:#08163F;background:linear-gradient(120deg,#27DA97 0%,#23D0B4 50%,#20C6CF 100%);}',
    '#pvLogin button[disabled]{opacity:.6;cursor:default;}',
    '#pvLogin .err{min-height:18px;font-size:12.5px;color:#BA2340;}',
    '.rep-signout{margin-left:6px;font:600 12px/1 "Poppins",system-ui,sans-serif;padding:7px 12px;border-radius:999px;border:1px solid rgba(127,140,170,.35);background:transparent;color:inherit;opacity:.75;cursor:pointer;vertical-align:middle;}',
    '.rep-signout:hover{opacity:1;}',
    '@media print{#pvLogin,.rep-signout{display:none !important;}}',
  ].join('\n');
  (document.head || root).appendChild(css);

  var LOGO = '<svg viewBox="0 0 64 64" aria-hidden="true"><defs><linearGradient id="pvLoginArc" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#27DA97"/><stop offset="1" stop-color="#20C6CF"/></linearGradient></defs><circle cx="32" cy="32" r="25" fill="none" stroke="#E3E7F3" stroke-width="6"/><circle cx="32" cy="32" r="25" fill="none" stroke="url(#pvLoginArc)" stroke-width="6" stroke-linecap="round" stroke-dasharray="125.7 157.1" transform="rotate(-90 32 32)"/><text x="32" y="40" font-family="Arial,Helvetica,sans-serif" font-weight="700" text-anchor="middle" font-size="22" fill="#16156A" letter-spacing="-.5">S3</text></svg>';

  function showLogin() {
    var box = document.createElement('div');
    box.id = 'pvLogin';
    box.innerHTML = '<div class="card">' + LOGO + '<h1>Sasha&rsquo;s Third Eye</h1><div class="org">Purview Healthcare &middot; Admin sign in</div>' +
      '<form id="pvLoginForm" autocomplete="on"><label>Username<input id="pvUser" name="username" autocomplete="username" required></label>' +
      '<label>Password<input id="pvPass" name="password" type="password" autocomplete="current-password" required></label>' +
      '<div class="err" id="pvErr" role="alert"></div><button type="submit" id="pvGo">Sign in</button></form></div>';
    document.body.appendChild(box);
    var form = box.querySelector('#pvLoginForm'), err = box.querySelector('#pvErr'), go = box.querySelector('#pvGo');
    box.querySelector('#pvUser').focus();
    form.addEventListener('submit', async function (e) {
      e.preventDefault();
      var u = box.querySelector('#pvUser').value.trim().toLowerCase(), p = box.querySelector('#pvPass').value;
      go.disabled = true; err.textContent = '';
      var ok = false;
      try { ok = u === ADMIN.user && (await hashOf(u, p)) === ADMIN.hash; } catch (x) { err.textContent = 'This browser cannot check the password here. Open the published site link instead.'; go.disabled = false; return; }
      if (!ok) { err.textContent = 'Wrong username or password.'; go.disabled = false; box.querySelector('#pvPass').select(); return; }
      writeSession(); box.remove(); root.classList.remove('pv-locked'); addSignOut();
      window.dispatchEvent(new Event('resize'));
    });
  }

  function addSignOut() {
    var nav = document.querySelector('.rep-tabs');
    if (!nav || document.querySelector('.rep-signout')) return;
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'rep-signout'; b.textContent = 'Sign out'; b.title = 'Sign out of the admin login';
    b.onclick = function () { clearSession(); location.reload(); };
    nav.insertAdjacentElement('afterend', b);
  }

  document.addEventListener('DOMContentLoaded', function () { if (readSession()) addSignOut(); else showLogin(); });
})();
