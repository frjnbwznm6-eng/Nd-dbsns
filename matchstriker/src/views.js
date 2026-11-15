'use strict';
const { db } = require('./db');
const { html, initials } = require('./lib');
const { planTerms, planIsFree, feePercent } = require('./core');

const logo = html`<svg class="logo-mark" viewBox="0 0 32 32" aria-hidden="true">
  <path d="M21.5 3.5c1 3.2 4.5 5 4.500 9.200a5.500 5.500 0 0 1-11 0c0-2.400 1.300-3.800 2.300-5.200.6 1.500 1.400 2.200 2.300 2.600-.3-2.500.5-4.800 1.900-6.600z" fill="#E4572E"/>
  <path d="M20.500 10.500c.5 1.200 1.700 1.900 1.700 3.400a1.900 1.900 0 0 1-3.800 0c0-1.400 1.300-2 2.100-3.400z" fill="#F6B042"/>
  <path d="M17.200 17.200 6.300 28.100a1.700 1.700 0 0 1-2.400-2.400l10.900-10.900z" fill="#1E1A17"/>
</svg>`;

const csrf = (ctx) => html`<input type="hidden" name="_csrf" value="${ctx.csrf}">`;

function unreadCount(userId) {
  return db.get(
    `SELECT COUNT(*) AS n FROM messages msg JOIN conversations c ON c.id = msg.conversation_id
     WHERE (c.user_a = ? OR c.user_b = ?) AND msg.sender_id != ? AND msg.read_at IS NULL`,
    userId, userId, userId
  ).n;
}

function layout(ctx, { title, body, description }) {
  const u = ctx.user;
  const unread = u ? unreadCount(u.id) : 0;
  const here = (p) => (ctx.path === p || ctx.path.startsWith(p + '/') ? 'active' : '');
  return html`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title ? title + ' · ' : ''}Matchstriker</title>
<meta name="description" content="${description || 'Find an independent matchmaker who understands your community, culture and city.'}">
<link rel="icon" href="/public/favicon.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600&display=swap">
<link rel="stylesheet" href="/public/style.css">
<script src="/public/app.js" defer></script>
</head>
<body>
<header class="site-header">
  <div class="wrap bar">
    <a class="brand" href="/">${logo}<span>matchstriker</span></a>
    <nav class="nav">
      <a class="${here('/matchmakers')}" href="/matchmakers">Find a matchmaker</a>
      ${!u || u.role !== 'matchmaker' ? html`<a class="${here('/for-matchmakers')}" href="/for-matchmakers">For matchmakers</a>` : ''}
      ${u ? html`
        <a class="${here('/messages')}" href="/messages">Messages${unread ? html` <span class="badge">${unread}</span>` : ''}</a>
        ${u.role === 'matchmaker'
          ? html`<a class="${here('/dashboard')}" href="/dashboard">Dashboard</a>`
          : html`<a class="${here('/account')}" href="/account">My account</a>`}
        <form method="post" action="/logout" class="inline">${csrf(ctx)}<button class="linklike">Log out</button></form>
      ` : html`
        <a class="${here('/login')}" href="/login">Log in</a>
        <a class="btn btn-small" href="/signup">Sign up</a>
      `}
    </nav>
  </div>
</header>
${ctx.flash ? html`<div class="flash" role="status"><div class="wrap">${ctx.flash}</div></div>` : ''}
<main>${body}</main>
<footer class="site-footer">
  <div class="wrap foot">
    <div><a class="brand" href="/">${logo}<span>matchstriker</span></a><p class="muted">Independent matchmakers, found by the people they serve best.</p></div>
    <div class="foot-links">
      <a href="/matchmakers">Find a matchmaker</a>
      <a href="/for-matchmakers">List your practice</a>
      <a href="/signup">Create an account</a>
    </div>
  </div>
</footer>
</body>
</html>`;
}

function errorPage(ctx, status, title, message) {
  return layout(ctx, {
    title,
    body: html`<section class="wrap narrow center pad-xl">
      <p class="eyebrow">${status}</p><h1>${title}</h1><p class="lede">${message}</p>
      <p><a class="btn" href="/">Back to home</a></p></section>`,
  });
}

const avatar = (name, hue, size = '') => html`<span class="avatar hue-${(hue || 0) % 8} ${size}" aria-hidden="true">${initials(name)}</span>`;
const tagList = (tags, cls = '') => tags.map((t) => html`<span class="tag ${cls}">${t}</span>`);
const statusPill = (s) => html`<span class="pill pill-${s}">${s}</span>`;

function pricingChips(plans) {
  const chips = [];
  if (plans.some(planIsFree)) chips.push('Free plan');
  if (plans.some((p) => p.monthly_fee)) chips.push('Membership');
  if (plans.some((p) => p.per_match_fee)) chips.push('Pay per match');
  if (plans.some((p) => p.per_date_fee)) chips.push('Pay per date');
  if (plans.some((p) => p.signup_fee) && chips.length < 3) chips.push('Sign-up fee');
  return chips;
}

function matchmakerCard(m) {
  const tags = db.all("SELECT tag FROM matchmaker_tags WHERE matchmaker_id = ? AND kind = 'community' ORDER BY rowid LIMIT 3", m.user_id).map((r) => r.tag);
  const plans = db.all('SELECT * FROM plans WHERE matchmaker_id = ? AND active = 1', m.user_id);
  return html`<a class="card mm-card" href="/m/${m.slug}">
    <div class="mm-card-head">
      ${avatar(m.business_name, m.hue)}
      <div><h3>${m.business_name}</h3><p class="muted small">${m.city}${m.years ? ` · ${m.years} yrs matchmaking` : ''}</p></div>
    </div>
    <p class="mm-headline">${m.headline}</p>
    <div class="tags">${tagList(tags)}</div>
    <div class="mm-card-foot">
      <span class="small muted">${pricingChips(plans).join(' · ') || 'Pricing on request'}</span>
      ${m.accepting ? '' : html`<span class="pill pill-paused">waitlist</span>`}
    </div>
  </a>`;
}

function planCard(plan, action) {
  return html`<div class="card plan">
    <h3>${plan.name}</h3>
    <ul class="terms">${planTerms(plan).map((t) => html`<li>${t}</li>`)}</ul>
    ${plan.description ? html`<p class="muted small">${plan.description}</p>` : ''}
    ${action || ''}
  </div>`;
}

function dashNav(ctx) {
  const items = [
    ['/dashboard', 'Overview'], ['/dashboard/listing', 'Listing'], ['/dashboard/plans', 'Plans & pricing'],
    ['/dashboard/clients', 'Clients'], ['/dashboard/matches', 'Matches'], ['/dashboard/payments', 'Payments'],
    ['/dashboard/network', 'Network'],
  ];
  const on = (p) => (p === '/dashboard' ? ctx.path === p : ctx.path.startsWith(p)) ? 'active' : '';
  return html`<nav class="subnav"><div class="wrap">${items.map(([p, l]) => html`<a class="${on(p)}" href="${p}">${l}</a>`)}</div></nav>`;
}

function dashLayout(ctx, { title, body }) {
  return layout(ctx, {
    title,
    body: html`${dashNav(ctx)}<section class="wrap pad">${body}</section>`,
  });
}

const options = (list, selected) =>
  list.map((o) => {
    const [value, label] = Array.isArray(o) ? o : [o, o];
    return html`<option value="${value}" ${String(value) === String(selected ?? '') ? 'selected' : ''}>${label}</option>`;
  });

const demoNote = () => html`<p class="demo-note">Demo mode: no card is charged. Payments are recorded so you can see the full flow, including the ${feePercent()} platform fee.</p>`;

module.exports = {
  layout, errorPage, csrf, avatar, tagList, statusPill, matchmakerCard, planCard, dashLayout,
  options, demoNote, pricingChips, logo,
};
