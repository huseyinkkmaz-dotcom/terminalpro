const { readFileSync } = require('fs');
const { join } = require('path');
const crypto = require('crypto');

const PASSWORD = process.env.SITE_PASSWORD || '';
const SECRET = process.env.AUTH_SECRET || crypto.randomBytes(32).toString('hex');

function makeToken() {
  const exp = String(Date.now() + 7 * 24 * 60 * 60 * 1000);
  const sig = crypto.createHmac('sha256', SECRET).update(exp).digest('hex');
  return exp + '.' + sig;
}

function verifyToken(token) {
  if (!token) return false;
  const parts = token.split('.');
  if (parts.length !== 2) return false;
  const expected = crypto.createHmac('sha256', SECRET).update(parts[0]).digest('hex');
  if (!crypto.timingSafeEqual(Buffer.from(parts[1], 'utf8'), Buffer.from(expected, 'utf8'))) return false;
  return parseInt(parts[0]) > Date.now();
}

function parseCookies(str) {
  const obj = {};
  (str || '').split(';').forEach(function(pair) {
    const idx = pair.indexOf('=');
    if (idx > 0) obj[pair.substring(0, idx).trim()] = pair.substring(idx + 1).trim();
  });
  return obj;
}

function serveApp(res) {
  const paths = [
    join(__dirname, '..', '_app.html'),
    join(process.cwd(), '_app.html'),
    join(process.cwd(), 'frontend', '_app.html')
  ];
  for (const p of paths) {
    try {
      const html = readFileSync(p, 'utf8');
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.setHeader('Cache-Control', 'private, no-store');
      res.setHeader('X-Content-Type-Options', 'nosniff');
      res.setHeader('X-Frame-Options', 'DENY');
      return res.status(200).send(html);
    } catch (e) { continue; }
  }
  res.status(500).send('App file not found');
}

function renderLogin(res, error) {
  const errHtml = error ? '<div class="err">' + error + '</div>' : '';
  const html = '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Terminal Pro</title><style>' +
    'body{font-family:Inter,system-ui,sans-serif;background:#070b14;color:#f9fafb;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}' +
    '.card{background:#111827;border:1px solid #374151;border-radius:16px;padding:48px 40px;max-width:360px;width:100%;text-align:center;box-shadow:0 25px 50px rgba(0,0,0,.5)}' +
    'h1{font-size:22px;margin:0 0 6px;color:#0ea5e9;font-weight:900;letter-spacing:-0.5px}' +
    '.sub{color:#9ca3af;font-size:12px;margin-bottom:28px}' +
    'input{width:100%;padding:14px 16px;background:#1f2937;border:1px solid #374151;border-radius:10px;color:#f9fafb;font-size:14px;margin-bottom:16px;box-sizing:border-box;outline:none;transition:border-color .2s}' +
    'input:focus{border-color:#0ea5e9}' +
    'button{width:100%;padding:14px;background:#0ea5e9;color:#fff;border:none;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;transition:background .2s;letter-spacing:0.3px}' +
    'button:hover{background:#0284c7}' +
    '.err{color:#ef4444;font-size:12px;margin-bottom:16px;padding:10px;background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);border-radius:8px}' +
    '.lock{font-size:32px;margin-bottom:16px}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="lock">&#128274;</div>' +
    '<h1>Terminal Pro</h1>' +
    '<div class="sub">This dashboard is private. Enter your password.</div>' +
    errHtml +
    '<form method="POST" action="/">' +
    '<input type="password" name="password" placeholder="Password" autofocus required autocomplete="current-password">' +
    '<button type="submit">Unlock Dashboard</button>' +
    '</form></div></body></html>';
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  return res.status(401).send(html);
}

module.exports = function handler(req, res) {
  if (!PASSWORD) {
    return serveApp(res);
  }

  const cookies = parseCookies(req.headers.cookie);
  if (verifyToken(cookies.tp_auth)) {
    return serveApp(res);
  }

  if (req.method === 'POST') {
    const password = (req.body && req.body.password) || '';
    if (password === PASSWORD) {
      const token = makeToken();
      res.setHeader('Set-Cookie', 'tp_auth=' + token + '; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=' + (7 * 24 * 60 * 60));
      res.writeHead(302, { Location: '/' });
      return res.end();
    }
    return renderLogin(res, 'Incorrect password. Try again.');
  }

  return renderLogin(res);
};
