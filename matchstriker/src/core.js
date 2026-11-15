'use strict';
// Business rules shared by routes and the seed script.
const { db } = require('./db');
const { money, slugify } = require('./lib');

// Platform fee in basis points (500 = 5%) taken from every paid transaction.
const FEE_BPS = Math.max(0, Math.min(5000, parseInt(process.env.PLATFORM_FEE_BPS || '500', 10) || 0));
const feePercent = () => (FEE_BPS / 100).toLocaleString('en-US') + '%';
const feeFor = (amount) => Math.round((amount * FEE_BPS) / 10000);

function addMonth(ms) {
  const d = new Date(ms);
  d.setMonth(d.getMonth() + 1);
  return d.getTime();
}

function createInvoice({ membership, kind, description, amount, paid = false, at = Date.now() }) {
  if (!(amount > 0)) return null;
  return db.insert(
    `INSERT INTO invoices (membership_id, client_id, matchmaker_id, kind, description, amount, fee, status, created_at, paid_at)
     VALUES (?,?,?,?,?,?,?,?,?,?)`,
    membership.id, membership.client_id, membership.matchmaker_id, kind, description,
    amount, feeFor(amount), paid ? 'paid' : 'due', at, paid ? at : null
  );
}

// Turn a membership on: bill the sign-up fee and the first month.
function activateMembership(membership, { paid = false, at = Date.now() } = {}) {
  const plan = db.get('SELECT * FROM plans WHERE id = ?', membership.plan_id);
  createInvoice({ membership, kind: 'signup', description: `Sign-up fee: ${plan.name}`, amount: plan.signup_fee, paid, at });
  createInvoice({ membership, kind: 'monthly', description: `Membership: ${plan.name}`, amount: plan.monthly_fee, paid, at });
  db.run('UPDATE memberships SET status = ?, next_bill_at = ? WHERE id = ?',
    'active', plan.monthly_fee > 0 ? addMonth(at) : null, membership.id);
}

// Raise monthly invoices that have come due. Called lazily on page loads;
// in production this would be a scheduled job or the payment processor's subscriptions.
function runBilling(where, id) {
  const col = where === 'client' ? 'client_id' : 'matchmaker_id';
  const rows = db.all(
    `SELECT m.*, p.name AS plan_name, p.monthly_fee FROM memberships m JOIN plans p ON p.id = m.plan_id
     WHERE m.${col} = ? AND m.status = 'active' AND m.next_bill_at IS NOT NULL AND m.next_bill_at <= ?`,
    id, Date.now()
  );
  for (const m of rows) {
    let next = m.next_bill_at;
    for (let i = 0; i < 12 && next <= Date.now(); i++) {
      createInvoice({ membership: m, kind: 'monthly', description: `Membership: ${m.plan_name}`, amount: m.monthly_fee, at: next });
      next = addMonth(next);
    }
    db.run('UPDATE memberships SET next_bill_at = ? WHERE id = ?', m.monthly_fee > 0 ? next : null, m.id);
  }
}

function planTerms(plan) {
  const t = [];
  if (plan.signup_fee) t.push(`${money(plan.signup_fee)} to join`);
  if (plan.monthly_fee) t.push(`${money(plan.monthly_fee)} / month`);
  if (plan.per_match_fee) t.push(`${money(plan.per_match_fee)} per match`);
  if (plan.per_date_fee) t.push(`${money(plan.per_date_fee)} per date`);
  return t.length ? t : ['Free'];
}
const planIsFree = (p) => !p.signup_fee && !p.monthly_fee && !p.per_match_fee && !p.per_date_fee;

function uniqueSlug(name) {
  const base = slugify(name);
  let slug = base;
  for (let i = 2; db.get('SELECT 1 AS x FROM matchmakers WHERE slug = ?', slug); i++) slug = `${base}-${i}`;
  return slug;
}

function setTags(matchmakerId, kind, tags) {
  db.run('DELETE FROM matchmaker_tags WHERE matchmaker_id = ? AND kind = ?', matchmakerId, kind);
  for (const tag of tags) db.run('INSERT INTO matchmaker_tags (matchmaker_id, kind, tag) VALUES (?,?,?)', matchmakerId, kind, tag);
}
function getTags(matchmakerId) {
  const out = { community: [], language: [], specialty: [] };
  for (const r of db.all('SELECT kind, tag FROM matchmaker_tags WHERE matchmaker_id = ? ORDER BY rowid', matchmakerId)) out[r.kind].push(r.tag);
  return out;
}

function getConversation(u1, u2, create = true) {
  const [a, b] = u1 < u2 ? [u1, u2] : [u2, u1];
  const row = db.get('SELECT * FROM conversations WHERE user_a = ? AND user_b = ?', a, b);
  if (row || !create) return row;
  const id = db.insert('INSERT INTO conversations (user_a, user_b, last_at) VALUES (?,?,?)', a, b, Date.now());
  return db.get('SELECT * FROM conversations WHERE id = ?', id);
}
function sendMessage(conv, senderId, body, at = Date.now()) {
  db.run('INSERT INTO messages (conversation_id, sender_id, body, created_at) VALUES (?,?,?,?)', conv.id, senderId, body, at);
  db.run('UPDATE conversations SET last_at = ? WHERE id = ?', at, conv.id);
}

// Display name for a user: matchmakers show their business name.
function displayName(userId) {
  const r = db.get(
    'SELECT u.name, u.role, m.business_name, m.slug, m.hue FROM users u LEFT JOIN matchmakers m ON m.user_id = u.id WHERE u.id = ?',
    userId
  );
  if (!r) return { name: 'Deleted user', hue: 0 };
  return { name: r.business_name || r.name, person: r.name, role: r.role, slug: r.slug, hue: r.hue ?? (userId % 8) };
}

function connectionBetween(a, b) {
  return db.get('SELECT * FROM connections WHERE (from_id = ? AND to_id = ?) OR (from_id = ? AND to_id = ?)', a, b, b, a);
}

module.exports = {
  FEE_BPS, feePercent, feeFor, addMonth, createInvoice, activateMembership, runBilling,
  planTerms, planIsFree, uniqueSlug, setTags, getTags, getConversation, sendMessage,
  displayName, connectionBetween,
};
