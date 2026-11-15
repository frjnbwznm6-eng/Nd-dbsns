'use strict';
process.removeAllListeners('warning'); // node:sqlite prints an ExperimentalWarning on Node 22
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { db } = require('./src/db');
const { Safe, token, safeNext } = require('./src/lib');
const views = require('./src/views');

const PORT = parseInt(process.env.PORT || '3000', 10);
const PROD = process.env.NODE_ENV === 'production';
const PUBLIC_DIR = path.join(__dirname, 'public');
const MIME = { '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const SESSION_MS = 30 * 24 * 3600 * 1000;

// --- Routing
const routes = [];
function add(method, pattern, auth, handler) {
  if (typeof auth === 'function') [handler, auth] = [auth, null];
  const keys = [];
  const re = new RegExp('^' + pattern.replace(/:(\w+)/g, (_, k) => (keys.push(k), '([^/]+)')) + '/?$');
  routes.push({ method, re, keys, auth, handler });
}
const app = {
  get: (p, a, h) => add('GET', p, a, h),
  post: (p, a, h) => add('POST', p, a, h),
};
require('./src/public')(app);
require('./src/account')(app);
require('./src/messages')(app);
require('./src/dashboard')(app);
require('./src/seed')();

function parseCookies(header) {
  const out = {};
  for (const part of String(header || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) {
      try { out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim()); } catch { /* ignore */ }
    }
  }
  return out;
}

function readBody(req, limit = 100 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) { reject(Object.assign(new Error('Body too large'), { status: 413 })); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Object.fromEntries(new URLSearchParams(Buffer.concat(chunks).toString('utf8')))));
    req.on('error', reject);
  });
}

function serveStatic(pathname, res) {
  const file = path.normalize(path.join(PUBLIC_DIR, pathname.slice('/public/'.length)));
  if (!file.startsWith(PUBLIC_DIR + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) {
    res.writeHead(404).end('Not found');
    return;
  }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Cache-Control': PROD ? 'public, max-age=3600' : 'no-cache' });
  fs.createReadStream(file).pipe(res);
}

async function handle(req, res) {
  const url = new URL(req.url, 'http://localhost');
  if (url.pathname.startsWith('/public/')) return serveStatic(url.pathname, res);

  const cookies = parseCookies(req.headers.cookie);
  const cookiesOut = [];
  const setCookie = (name, value, maxAgeSec) =>
    cookiesOut.push(`${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAgeSec}${PROD ? '; Secure' : ''}`);

  const ctx = {
    req, res, url, path: url.pathname, query: Object.fromEntries(url.searchParams), body: {}, params: {},
    user: null, mm: null, flash: cookies.flash || null, csrf: cookies.csrf, setCookie,
    ip: req.socket.remoteAddress,
  };
  if (!ctx.csrf || !/^[0-9a-f]{32}$/.test(ctx.csrf)) {
    ctx.csrf = token(16);
    setCookie('csrf', ctx.csrf, 365 * 24 * 3600);
  }

  const headers = () => ({
    'Content-Security-Policy': "default-src 'self'; style-src 'self' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; img-src 'self' data:; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'same-origin',
    'Cache-Control': 'no-store',
    ...(cookiesOut.length ? { 'Set-Cookie': cookiesOut } : {}),
  });
  ctx.send = (body, status = 200, type = 'text/html; charset=utf-8', extra = {}) => {
    if (ctx.flash && type.startsWith('text/html')) setCookie('flash', '', 0);
    res.writeHead(status, { 'Content-Type': type, ...headers(), ...extra });
    res.end(String(body));
  };
  ctx.redirect = (to, flash) => {
    if (flash) setCookie('flash', flash, 60);
    res.writeHead(303, { Location: to, ...headers() });
    res.end();
  };
  ctx.error = (status, title, message) => ctx.send(views.errorPage(ctx, status, title, message), status);
  ctx.login = (userId) => {
    const sid = token();
    db.run('INSERT INTO sessions (id, user_id, created_at) VALUES (?,?,?)', sid, userId, Date.now());
    setCookie('sid', sid, SESSION_MS / 1000);
  };
  ctx.logout = () => {
    if (cookies.sid) db.run('DELETE FROM sessions WHERE id = ?', cookies.sid);
    setCookie('sid', '', 0);
  };

  if (cookies.sid) {
    const row = db.get(
      'SELECT u.id, u.email, u.name, u.role, s.created_at AS session_at FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.id = ?',
      cookies.sid
    );
    if (row && Date.now() - row.session_at < SESSION_MS) {
      ctx.user = row;
      if (row.role === 'matchmaker') ctx.mm = db.get('SELECT * FROM matchmakers WHERE user_id = ?', row.id);
    }
  }

  if (req.method !== 'GET' && req.method !== 'POST') return ctx.error(405, 'Method not allowed', '');
  if (req.method === 'POST') {
    ctx.body = await readBody(req);
    if (ctx.body._csrf !== ctx.csrf) return ctx.error(403, 'Form expired', 'Please go back, reload the page and try again.');
  }

  for (const r of routes) {
    if (r.method !== req.method) continue;
    const m = r.re.exec(ctx.path);
    if (!m) continue;
    r.keys.forEach((k, i) => { ctx.params[k] = decodeURIComponent(m[i + 1]); });
    if (r.auth) {
      if (!ctx.user) {
        const next = req.method === 'GET' ? safeNext(url.pathname + url.search) : null;
        return ctx.redirect('/login' + (next ? '?next=' + encodeURIComponent(next) : ''), 'Please log in to continue.');
      }
      if (r.auth !== 'user' && ctx.user.role !== r.auth) {
        return ctx.error(403, 'Wrong account type',
          r.auth === 'matchmaker' ? 'This area is for matchmaker accounts.' : 'This area is for client accounts. Matchmakers manage their business from the dashboard.');
      }
    }
    const out = await r.handler(ctx);
    if (out instanceof Safe && !res.writableEnded) ctx.send(out);
    return;
  }
  ctx.error(404, 'Page not found', 'That page does not exist, or the listing has been taken down.');
}

http.createServer((req, res) => {
  handle(req, res).catch((err) => {
    if (!err.status) console.error(err);
    if (res.writableEnded) return;
    if (!res.headersSent) res.writeHead(err.status || 500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(err.status === 413 ? 'Request too large' : 'Something went wrong on our side.');
  });
}).listen(PORT, () => console.log(`matchstriker running at http://localhost:${PORT}`));
