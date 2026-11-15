# Matchstriker

A marketplace for small, independent matchmakers. Clients search by community, city, language and
pricing model; matchmakers run their whole practice from one dashboard. The platform earns a
percentage of every payment that goes through it.

## Run it

Needs Node 22.13 or newer. There are no dependencies to install.

```bash
cd matchstriker
npm start
```

Open http://localhost:3000. On first run the site seeds fictional sample matchmakers and two demo
logins (password `demo1234`):

- Matchmaker: `matchmaker@demo.matchstriker.com`
- Client: `client@demo.matchstriker.com`

`npm run reset` wipes the database and reseeds.

## What is in it

For clients
- Directory search with filters for community, city, language and pricing model
- Matchmaker pages with bio, process and plans
- Join or apply on a plan, private intake profile, introductions and dates, invoices, messaging

For matchmakers
- Sign-up to published listing in three steps
- Plans that combine a sign-up fee, monthly fee, per-match fee and per-date fee (all zero = free)
- Optional approval of each client before they join, waitlist switch
- Client table with search, status filter, private notes, plan changes, one-off charges, CSV export
- Matches and dates, which raise per-match and per-date invoices automatically
- Payments ledger showing gross, platform fee and net on every line
- Network: connect with other matchmakers and message them
- One inbox for clients, prospects and peers

## Settings (environment variables)

| Variable | Default | Meaning |
| --- | --- | --- |
| `PORT` | `3000` | Port to listen on |
| `PLATFORM_FEE_BPS` | `500` | Platform fee in basis points (500 = 5%) |
| `DEMO` | on | Set to `0` to skip sample data and hide the demo logins |
| `NODE_ENV` | | Set to `production` for Secure cookies (requires HTTPS) |
| `DB_PATH` | `data/matchstriker.db` | SQLite database file |

## Before going live

Payments are in demo mode: pressing Pay marks an invoice paid and records the platform fee, but no
card is charged and no money moves. To take real money, connect a marketplace payment provider
(for example Stripe Connect) at the two marked places: where a client pays
(`src/account.js`, the join and invoice-pay handlers) and monthly billing (`runBilling` in
`src/core.js`). The fee recorded on each invoice maps directly onto the provider's application fee.

Also still to do for a real launch: email verification and password reset, email notifications,
profile photos, matchmaker vetting, terms and privacy pages, and hosting with HTTPS, backups and
`DEMO=0`.

## Layout

- `server.js` – HTTP server, routing, sessions, CSRF, security headers
- `src/db.js` – schema; `src/core.js` – billing and business rules; `src/seed.js` – sample data
- `src/public.js`, `account.js`, `messages.js`, `dashboard.js` – pages
- `src/views.js` – shared layout and components; `public/` – CSS, JS, icon
