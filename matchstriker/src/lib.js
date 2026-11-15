'use strict';
const crypto = require('node:crypto');

// --- Safe HTML templating: every interpolated value is escaped unless it is
// itself the result of an html`` template (or an array of them).
class Safe {
  constructor(s) { this.s = s; }
  toString() { return this.s; }
}
const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ESC[c]);
function render(v) {
  if (v === null || v === undefined || v === false || v === true) return '';
  if (v instanceof Safe) return v.s;
  if (Array.isArray(v)) return v.map(render).join('');
  return esc(v);
}
function html(strings, ...vals) {
  let out = '';
  strings.forEach((s, i) => {
    out += s;
    if (i < vals.length) out += render(vals[i]);
  });
  return new Safe(out);
}

// --- Money
function money(cents) {
  const d = (cents || 0) / 100;
  const opts = Number.isInteger(d) ? {} : { minimumFractionDigits: 2, maximumFractionDigits: 2 };
  return '$' + d.toLocaleString('en-US', opts);
}
function toCents(v) {
  const n = Math.round(parseFloat(String(v ?? '').replace(/[$,\s]/g, '')) * 100);
  return Number.isFinite(n) && n > 0 ? Math.min(n, 10000000) : 0;
}

// --- Passwords (scrypt)
function hashPassword(pw) {
  const salt = crypto.randomBytes(16);
  return salt.toString('hex') + ':' + crypto.scryptSync(pw, salt, 64).toString('hex');
}
function verifyPassword(pw, stored) {
  const [saltHex, hashHex] = String(stored).split(':');
  if (!saltHex || !hashHex) return false;
  const expected = Buffer.from(hashHex, 'hex');
  const actual = crypto.scryptSync(pw, Buffer.from(saltHex, 'hex'), expected.length);
  return crypto.timingSafeEqual(expected, actual);
}
const token = (bytes = 32) => crypto.randomBytes(bytes).toString('hex');

// --- Strings and dates
const clean = (v, max = 200) => String(v ?? '').trim().slice(0, max);
const slugify = (s) =>
  String(s).toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'matchmaker';
const splitTags = (v) =>
  [...new Set(String(v ?? '').split(',').map((t) => t.trim().slice(0, 40)).filter(Boolean))].slice(0, 12);
const fmtDate = (ms) =>
  new Date(ms).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
const fmtTime = (ms) =>
  new Date(ms).toLocaleString('en-US', { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
const firstName = (name) => String(name || '').trim().split(/\s+/)[0] || 'Someone';
const initials = (name) =>
  (String(name || '').match(/[\p{L}\p{N}]+/gu) || ['?']).slice(0, 2).map((w) => w[0]).join('').toUpperCase();
// Only allow same-site relative redirects.
const safeNext = (v) => (typeof v === 'string' && /^\/(?!\/)[\w\-./?=&%]*$/.test(v) ? v : null);

module.exports = {
  Safe, html, esc, money, toCents, hashPassword, verifyPassword, token,
  clean, slugify, splitTags, fmtDate, fmtTime, firstName, initials, safeNext,
};
