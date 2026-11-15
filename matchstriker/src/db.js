'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');

const DB_PATH = process.env.DB_PATH || path.join(__dirname, '..', 'data', 'matchstriker.db');
if (DB_PATH !== ':memory:') fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

const sqlite = new DatabaseSync(DB_PATH);
sqlite.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON;');

sqlite.exec(`
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY,
  email TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  name TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('client','matchmaker')),
  created_at INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
  id TEXT PRIMARY KEY,
  user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  created_at INTEGER NOT NULL
);
-- A matchmaker's public listing and business rules.
CREATE TABLE IF NOT EXISTS matchmakers (
  user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
  slug TEXT NOT NULL UNIQUE,
  business_name TEXT NOT NULL,
  headline TEXT NOT NULL DEFAULT '',
  bio TEXT NOT NULL DEFAULT '',
  approach TEXT NOT NULL DEFAULT '',
  city TEXT NOT NULL DEFAULT '',
  years INTEGER NOT NULL DEFAULT 0,
  accepting INTEGER NOT NULL DEFAULT 1,
  requires_approval INTEGER NOT NULL DEFAULT 0,
  published INTEGER NOT NULL DEFAULT 0,
  hue INTEGER NOT NULL DEFAULT 0,
  created_at INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS matchmaker_tags (
  matchmaker_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  kind TEXT NOT NULL CHECK (kind IN ('community','language','specialty')),
  tag TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_tags ON matchmaker_tags(kind, tag);
-- Pricing plans. Any combination of the four fees; all zero = free plan.
CREATE TABLE IF NOT EXISTS plans (
  id INTEGER PRIMARY KEY,
  matchmaker_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  description TEXT NOT NULL DEFAULT '',
  signup_fee INTEGER NOT NULL DEFAULT 0,
  monthly_fee INTEGER NOT NULL DEFAULT 0,
  per_match_fee INTEGER NOT NULL DEFAULT 0,
  per_date_fee INTEGER NOT NULL DEFAULT 0,
  active INTEGER NOT NULL DEFAULT 1,
  created_at INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS client_profiles (
  user_id INTEGER PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
  age INTEGER,
  city TEXT NOT NULL DEFAULT '',
  about TEXT NOT NULL DEFAULT '',
  seeking TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS memberships (
  id INTEGER PRIMARY KEY,
  client_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  matchmaker_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  plan_id INTEGER NOT NULL REFERENCES plans(id),
  status TEXT NOT NULL CHECK (status IN ('pending','active','paused','ended')),
  notes TEXT NOT NULL DEFAULT '',
  next_bill_at INTEGER,
  created_at INTEGER NOT NULL,
  UNIQUE (client_id, matchmaker_id)
);
CREATE TABLE IF NOT EXISTS introductions (
  id INTEGER PRIMARY KEY,
  matchmaker_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  a_membership_id INTEGER NOT NULL REFERENCES memberships(id),
  b_membership_id INTEGER REFERENCES memberships(id),
  b_external TEXT NOT NULL DEFAULT '',
  note TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'introduced' CHECK (status IN ('introduced','dating','paused','success','closed')),
  created_at INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS dates (
  id INTEGER PRIMARY KEY,
  introduction_id INTEGER NOT NULL REFERENCES introductions(id) ON DELETE CASCADE,
  happens_on TEXT NOT NULL,
  venue TEXT NOT NULL DEFAULT '',
  created_at INTEGER NOT NULL
);
-- Every charge. fee = the platform's cut, fixed when the invoice is created.
CREATE TABLE IF NOT EXISTS invoices (
  id INTEGER PRIMARY KEY,
  membership_id INTEGER NOT NULL REFERENCES memberships(id),
  client_id INTEGER NOT NULL REFERENCES users(id),
  matchmaker_id INTEGER NOT NULL REFERENCES matchmakers(user_id),
  kind TEXT NOT NULL CHECK (kind IN ('signup','monthly','match','date','custom')),
  description TEXT NOT NULL,
  amount INTEGER NOT NULL,
  fee INTEGER NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('due','paid','void')),
  created_at INTEGER NOT NULL,
  paid_at INTEGER
);
CREATE INDEX IF NOT EXISTS idx_inv_mm ON invoices(matchmaker_id);
CREATE INDEX IF NOT EXISTS idx_inv_client ON invoices(client_id);
CREATE TABLE IF NOT EXISTS conversations (
  id INTEGER PRIMARY KEY,
  user_a INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  user_b INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  last_at INTEGER NOT NULL,
  UNIQUE (user_a, user_b)
);
CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY,
  conversation_id INTEGER NOT NULL REFERENCES conversations(id) ON DELETE CASCADE,
  sender_id INTEGER NOT NULL REFERENCES users(id),
  body TEXT NOT NULL,
  created_at INTEGER NOT NULL,
  read_at INTEGER
);
CREATE INDEX IF NOT EXISTS idx_msg_conv ON messages(conversation_id, id);
-- Matchmaker-to-matchmaker network.
CREATE TABLE IF NOT EXISTS connections (
  id INTEGER PRIMARY KEY,
  from_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  to_id INTEGER NOT NULL REFERENCES matchmakers(user_id) ON DELETE CASCADE,
  status TEXT NOT NULL CHECK (status IN ('pending','accepted')),
  created_at INTEGER NOT NULL,
  UNIQUE (from_id, to_id)
);
`);

const cache = new Map();
function stmt(sql) {
  let s = cache.get(sql);
  if (!s) cache.set(sql, (s = sqlite.prepare(sql)));
  return s;
}
const db = {
  get: (sql, ...p) => stmt(sql).get(...p),
  all: (sql, ...p) => stmt(sql).all(...p),
  run: (sql, ...p) => stmt(sql).run(...p),
  insert: (sql, ...p) => Number(stmt(sql).run(...p).lastInsertRowid),
  tx(fn) {
    sqlite.exec('BEGIN');
    try {
      const r = fn();
      sqlite.exec('COMMIT');
      return r;
    } catch (e) {
      sqlite.exec('ROLLBACK');
      throw e;
    }
  },
};

module.exports = { db };
