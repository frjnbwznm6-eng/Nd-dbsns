'use strict';
const { db } = require('./db');
const { html, clean, hashPassword, verifyPassword, safeNext, money } = require('./lib');
const { getTags, uniqueSlug, feePercent, FEE_BPS, feeFor } = require('./core');
const v = require('./views');

const DEMO = process.env.DEMO !== '0';
const PRICING = [
  ['free', 'Free plan available', 'p.signup_fee = 0 AND p.monthly_fee = 0 AND p.per_match_fee = 0 AND p.per_date_fee = 0'],
  ['match', 'Pay per match', 'p.per_match_fee > 0'],
  ['date', 'Pay per date', 'p.per_date_fee > 0'],
  ['membership', 'Monthly membership', 'p.monthly_fee > 0'],
];

const distinctTags = (kind) =>
  db.all(
    `SELECT t.tag, COUNT(*) AS n FROM matchmaker_tags t JOIN matchmakers m ON m.user_id = t.matchmaker_id
     WHERE t.kind = ? AND m.published = 1 GROUP BY t.tag ORDER BY n DESC, t.tag`, kind
  ).map((r) => r.tag);
const cities = () => db.all("SELECT DISTINCT city FROM matchmakers WHERE published = 1 AND city != '' ORDER BY city").map((r) => r.city);

// Simple in-memory throttle for login attempts.
const attempts = new Map();
function throttled(ip) {
  const now = Date.now();
  const list = (attempts.get(ip) || []).filter((t) => now - t < 10 * 60 * 1000);
  list.push(now);
  attempts.set(ip, list);
  return list.length > 10;
}

module.exports = (app) => {
  app.get('/', (ctx) => {
    const featured = db.all('SELECT * FROM matchmakers WHERE published = 1 ORDER BY accepting DESC, years DESC LIMIT 6');
    const communities = distinctTags('community').slice(0, 10);
    return v.layout(ctx, {
      body: html`
<section class="hero">
  <div class="wrap hero-grid">
    <div>
      <p class="eyebrow">Independent matchmakers only</p>
      <h1>Find the matchmaker who already knows your world.</h1>
      <p class="lede">Matchstriker is a directory of small, independent matchmakers: the Jewish matchmaker in Brooklyn, the Korean matchmaker in LA, the one your cousin swears by. Search by community, city and budget, then work with them directly.</p>
      <form class="search" action="/matchmakers" method="get">
        <input type="search" name="q" placeholder="Try &quot;Jewish&quot;, &quot;Korean&quot;, &quot;over 50&quot;…" aria-label="Search matchmakers">
        <select name="city" aria-label="City"><option value="">Any city</option>${v.options(cities())}</select>
        <button class="btn">Search</button>
      </form>
      <div class="chips">${communities.map((c) => html`<a class="chip" href="/matchmakers?community=${encodeURIComponent(c)}">${c}</a>`)}</div>
    </div>
    <div class="hero-art" aria-hidden="true">
      <div class="spark-card s1">${v.avatar('Rivka Adler', 0)}<div><strong>Jewish · New York</strong><span>Free intro call, pay per match</span></div></div>
      <div class="spark-card s2">${v.avatar('Min Ji', 3)}<div><strong>Korean · Los Angeles</strong><span>Membership, family introductions</span></div></div>
      <div class="spark-card s3">${v.avatar('Sana Khan', 5)}<div><strong>Muslim · Chicago</strong><span>Pay per date, no sign-up fee</span></div></div>
    </div>
  </div>
</section>

<section class="wrap pad">
  <div class="section-head"><h2>Matchmakers taking clients now</h2><a href="/matchmakers">See all</a></div>
  <div class="grid cards">${featured.map(v.matchmakerCard)}</div>
</section>

<section class="band">
  <div class="wrap pad">
    <h2>How it works</h2>
    <div class="grid three steps">
      <div><span class="step">1</span><h3>Search by what matters</h3><p>Community, faith, language, city, age group. Every listing is a real person running their own practice, not a call center.</p></div>
      <div><span class="step">2</span><h3>Compare how they charge</h3><p>Each matchmaker sets their own terms: free plans, a sign-up fee, monthly membership, or pay only per match or per date. It is all on their page.</p></div>
      <div><span class="step">3</span><h3>Join and talk directly</h3><p>Sign up with the matchmaker you like, message them here, and keep your introductions, dates and payments in one place.</p></div>
    </div>
  </div>
</section>

<section class="wrap pad">
  <div class="cta">
    <div><h2>Are you a matchmaker?</h2><p>Run your whole practice from one place: listing, client table, messaging, pricing rules and payments. No monthly fee. We take ${feePercent()} of what you earn through the site, and nothing if you earn nothing.</p></div>
    <a class="btn btn-light" href="/for-matchmakers">See how it works</a>
  </div>
</section>`,
    });
  });

  app.get('/matchmakers', (ctx) => {
    const q = clean(ctx.query.q, 80);
    const f = { community: clean(ctx.query.community, 40), city: clean(ctx.query.city, 60), language: clean(ctx.query.language, 40), pricing: clean(ctx.query.pricing, 20) };
    const where = ['m.published = 1'];
    const params = [];
    if (q) {
      const like = '%' + q.replace(/[%_\\]/g, '\\$&') + '%';
      where.push(`(m.business_name LIKE ? ESCAPE '\\' OR m.headline LIKE ? ESCAPE '\\' OR m.bio LIKE ? ESCAPE '\\' OR m.city LIKE ? ESCAPE '\\'
        OR EXISTS (SELECT 1 FROM matchmaker_tags t WHERE t.matchmaker_id = m.user_id AND t.tag LIKE ? ESCAPE '\\'))`);
      params.push(like, like, like, like, like);
    }
    for (const kind of ['community', 'language']) {
      if (f[kind]) {
        where.push('EXISTS (SELECT 1 FROM matchmaker_tags t WHERE t.matchmaker_id = m.user_id AND t.kind = ? AND t.tag = ? COLLATE NOCASE)');
        params.push(kind, f[kind]);
      }
    }
    if (f.city) { where.push('m.city = ?'); params.push(f.city); }
    const pricing = PRICING.find((p) => p[0] === f.pricing);
    if (pricing) where.push(`EXISTS (SELECT 1 FROM plans p WHERE p.matchmaker_id = m.user_id AND p.active = 1 AND ${pricing[2]})`);
    const rows = db.all(`SELECT m.* FROM matchmakers m WHERE ${where.join(' AND ')} ORDER BY m.accepting DESC, m.years DESC, m.business_name LIMIT 60`, ...params);
    const filtered = q || Object.values(f).some(Boolean);
    return v.layout(ctx, {
      title: 'Find a matchmaker',
      body: html`
<section class="wrap pad">
  <h1>Find a matchmaker</h1>
  <form class="filters" method="get" action="/matchmakers">
    <label>Search<input type="search" name="q" value="${q}" placeholder="Community, name, specialty"></label>
    <label>Community<select name="community"><option value="">Any</option>${v.options(distinctTags('community'), f.community)}</select></label>
    <label>City<select name="city"><option value="">Any</option>${v.options(cities(), f.city)}</select></label>
    <label>Language<select name="language"><option value="">Any</option>${v.options(distinctTags('language'), f.language)}</select></label>
    <label>Pricing<select name="pricing"><option value="">Any</option>${v.options(PRICING.map((p) => [p[0], p[1]]), f.pricing)}</select></label>
    <div class="filter-actions"><button class="btn">Search</button>${filtered ? html`<a href="/matchmakers">Clear</a>` : ''}</div>
  </form>
  <p class="muted">${rows.length} ${rows.length === 1 ? 'matchmaker' : 'matchmakers'}${filtered ? ' match your search' : ''}</p>
  ${rows.length
    ? html`<div class="grid cards">${rows.map(v.matchmakerCard)}</div>`
    : html`<div class="empty"><h3>No matchmakers match that yet</h3><p>Try a broader search, or <a href="/matchmakers">browse everyone</a>. Know a matchmaker who should be here? Send them to <a href="/for-matchmakers">our matchmaker page</a>.</p></div>`}
</section>`,
    });
  });

  app.get('/m/:slug', (ctx) => {
    const m = db.get('SELECT * FROM matchmakers WHERE slug = ?', ctx.params.slug);
    const own = m && ctx.user && ctx.user.id === m.user_id;
    if (!m || (!m.published && !own)) return ctx.error(404, 'Matchmaker not found', 'This listing does not exist or is not published.');
    const tags = getTags(m.user_id);
    const plans = db.all('SELECT * FROM plans WHERE matchmaker_id = ? AND active = 1 ORDER BY signup_fee + monthly_fee + per_match_fee + per_date_fee, id', m.user_id);
    const person = db.get('SELECT name FROM users WHERE id = ?', m.user_id).name;
    const membership = ctx.user && ctx.user.role === 'client'
      ? db.get("SELECT * FROM memberships WHERE client_id = ? AND matchmaker_id = ? AND status != 'ended'", ctx.user.id, m.user_id) : null;
    const joinAction = (plan) => {
      if (own) return '';
      if (membership) return membership.plan_id === plan.id ? html`<p class="small"><strong>Your plan</strong> · ${v.statusPill(membership.status)}</p>` : '';
      if (ctx.user && ctx.user.role === 'matchmaker') return '';
      return html`<a class="btn btn-block" href="/m/${m.slug}/join/${plan.id}">${m.requires_approval ? 'Apply' : 'Join'} on this plan</a>`;
    };
    const messageAction = own ? html`<a class="btn btn-ghost" href="/dashboard/listing">Edit listing</a>`
      : ctx.user
        ? html`<form method="post" action="/messages/start">${v.csrf(ctx)}<input type="hidden" name="to" value="${m.user_id}"><button class="btn btn-ghost">Message ${person.split(' ')[0]}</button></form>`
        : html`<a class="btn btn-ghost" href="/signup?next=${encodeURIComponent('/m/' + m.slug)}">Sign up to message</a>`;
    return v.layout(ctx, {
      title: m.business_name,
      description: m.headline,
      body: html`
${!m.published ? html`<div class="flash warn"><div class="wrap">This listing is not published yet. Only you can see it. <a href="/dashboard/listing">Publish it from your dashboard.</a></div></div>` : ''}
<section class="profile-head">
  <div class="wrap profile-grid">
    ${v.avatar(m.business_name, m.hue, 'xl')}
    <div>
      <h1>${m.business_name}</h1>
      <p class="lede">${m.headline}</p>
      <p class="muted">${person}${m.city ? ` · ${m.city}` : ''}${m.years ? ` · ${m.years} years matchmaking` : ''} · ${m.accepting ? 'Taking new clients' : 'Waitlist only'}</p>
      <div class="tags">${v.tagList(tags.community)}${v.tagList(tags.language, 'tag-alt')}</div>
    </div>
    <div class="profile-actions">${messageAction}</div>
  </div>
</section>
<section class="wrap pad profile-body">
  <div>
    <h2>About</h2>
    <p class="prose">${m.bio || 'This matchmaker has not written a bio yet.'}</p>
    ${m.approach ? html`<h2>How I work</h2><p class="prose">${m.approach}</p>` : ''}
    ${tags.specialty.length ? html`<h2>Specialties</h2><div class="tags">${v.tagList(tags.specialty, 'tag-plain')}</div>` : ''}
  </div>
  <aside>
    <h2>Plans</h2>
    ${m.requires_approval ? html`<p class="small muted">${person.split(' ')[0]} reviews every application. You pay nothing until you are accepted.</p>` : ''}
    <div class="plan-stack">${plans.length ? plans.map((p) => v.planCard(p, joinAction(p))) : html`<p class="muted">No plans published yet.</p>`}</div>
  </aside>
</section>`,
    });
  });

  app.get('/for-matchmakers', (ctx) => {
    const ex = [[30000, 'Sign-up fee'], [15000, 'A match fee'], [4000, 'A date fee']];
    return v.layout(ctx, {
      title: 'For matchmakers',
      body: html`
<section class="hero hero-plain">
  <div class="wrap narrow center">
    <p class="eyebrow">For independent matchmakers</p>
    <h1>Your whole practice, set up in an afternoon.</h1>
    <p class="lede">You know your community. You should not also have to stitch together a website, a spreadsheet, a payment link and three chat apps. Matchstriker gives a solo matchmaker the back office of a big agency.</p>
    <p><a class="btn" href="/signup?role=matchmaker">List your practice</a></p>
  </div>
</section>
<section class="wrap pad">
  <div class="grid three features">
    <div class="card"><h3>Be found by the right people</h3><p>Clients search by community, language, city and budget. Tag your listing and you show up for exactly the people you serve.</p></div>
    <div class="card"><h3>Your rules, your pricing</h3><p>Offer a free plan, charge to join, bill monthly, or charge per match or per date. Mix them in as many plans as you like, and choose whether to approve each client first.</p></div>
    <div class="card"><h3>Payments built in</h3><p>Fees are invoiced automatically when a client joins, when you make a match and when you log a date. You see gross, fee and net on every line.</p></div>
    <div class="card"><h3>A client table that is yours</h3><p>Every client with plan, status, intake answers and your private notes. Filter it, search it, export it to CSV whenever you want.</p></div>
    <div class="card"><h3>One inbox</h3><p>Message clients and prospects in one place instead of across text, email and WhatsApp.</p></div>
    <div class="card"><h3>A network of peers</h3><p>Connect with other independent matchmakers to trade referrals when a client is a better fit for someone else's community.</p></div>
  </div>
</section>
<section class="band">
  <div class="wrap pad narrow">
    <h2>What it costs</h2>
    <p class="lede">No listing fee and no subscription. Matchstriker keeps ${feePercent()} of each payment your clients make through the site. If you run a free practice, it is free.</p>
    <table class="table">
      <thead><tr><th>Your client pays</th><th class="num">Amount</th><th class="num">Our ${feePercent()}</th><th class="num">You keep</th></tr></thead>
      <tbody>${ex.map(([amt, label]) => html`<tr><td>${label}</td><td class="num">${money(amt)}</td><td class="num">${money(feeFor(amt))}</td><td class="num"><strong>${money(amt - feeFor(amt))}</strong></td></tr>`)}</tbody>
    </table>
    ${FEE_BPS ? html`<p class="small muted">Card processing costs from the payment provider are separate.</p>` : ''}
    <p><a class="btn" href="/signup?role=matchmaker">Create your listing</a></p>
  </div>
</section>`,
    });
  });

  // --- Auth
  const authPage = (ctx, { title, form, error, alt }) =>
    v.layout(ctx, {
      title,
      body: html`<section class="wrap auth pad">
        <h1>${title}</h1>
        ${error ? html`<p class="error" role="alert">${error}</p>` : ''}
        ${form}
        <p class="muted">${alt}</p>
      </section>`,
    });

  const signupPage = (ctx, vals = {}, error) => {
    const role = vals.role === 'matchmaker' ? 'matchmaker' : 'client';
    const next = safeNext(vals.next);
    return authPage(ctx, {
      title: 'Create your account',
      error,
      alt: html`Already have an account? <a href="/login">Log in</a>`,
      form: html`<form method="post" action="/signup" class="form">
        ${v.csrf(ctx)}${next ? html`<input type="hidden" name="next" value="${next}">` : ''}
        <fieldset class="role-pick">
          <legend>I am…</legend>
          <label><input type="radio" name="role" value="client" ${role === 'client' ? 'checked' : ''}> Looking for a matchmaker</label>
          <label><input type="radio" name="role" value="matchmaker" ${role === 'matchmaker' ? 'checked' : ''}> A matchmaker</label>
        </fieldset>
        <label>Your name<input name="name" required maxlength="80" value="${vals.name || ''}" autocomplete="name"></label>
        <label data-role-only="matchmaker">Business name <span class="muted">(matchmakers only)</span><input name="business_name" maxlength="80" value="${vals.business_name || ''}" placeholder="e.g. Adler Introductions"></label>
        <label>Email<input type="email" name="email" required maxlength="120" value="${vals.email || ''}" autocomplete="email"></label>
        <label>Password <span class="muted">(8+ characters)</span><input type="password" name="password" required minlength="8" maxlength="200" autocomplete="new-password"></label>
        <button class="btn btn-block">Create account</button>
      </form>`,
    });
  };

  app.get('/signup', (ctx) => (ctx.user ? ctx.redirect(ctx.user.role === 'matchmaker' ? '/dashboard' : '/account') : signupPage(ctx, ctx.query)));

  app.post('/signup', (ctx) => {
    const b = ctx.body;
    const role = b.role === 'matchmaker' ? 'matchmaker' : 'client';
    const name = clean(b.name, 80);
    const email = clean(b.email, 120).toLowerCase();
    const password = String(b.password || '');
    const business = clean(b.business_name, 80) || (name ? `${name} Matchmaking` : '');
    const fail = (msg) => ctx.send(signupPage(ctx, b, msg), 400);
    if (!name) return fail('Please enter your name.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail('Please enter a valid email address.');
    if (password.length < 8) return fail('Password must be at least 8 characters.');
    if (db.get('SELECT 1 AS x FROM users WHERE email = ?', email)) return fail('An account with that email already exists. Try logging in.');
    const id = db.tx(() => {
      const uid = db.insert('INSERT INTO users (email, password_hash, name, role, created_at) VALUES (?,?,?,?,?)', email, hashPassword(password), name, role, Date.now());
      if (role === 'matchmaker') {
        db.run('INSERT INTO matchmakers (user_id, slug, business_name, hue, created_at) VALUES (?,?,?,?,?)', uid, uniqueSlug(business), business, uid % 8, Date.now());
      } else {
        db.run('INSERT INTO client_profiles (user_id) VALUES (?)', uid);
      }
      return uid;
    });
    ctx.login(id);
    if (role === 'matchmaker') return ctx.redirect('/dashboard', 'Welcome! Three steps to go live: fill in your listing, add a plan, publish.');
    ctx.redirect(safeNext(b.next) || '/account/profile', 'Welcome to Matchstriker.');
  });

  const loginPage = (ctx, vals = {}, error) => {
    const next = safeNext(vals.next);
    return authPage(ctx, {
      title: 'Log in',
      error,
      alt: html`New here? <a href="/signup${next ? '?next=' + encodeURIComponent(next) : ''}">Create an account</a>`,
      form: html`<form method="post" action="/login" class="form">
        ${v.csrf(ctx)}${next ? html`<input type="hidden" name="next" value="${next}">` : ''}
        <label>Email<input type="email" name="email" required value="${vals.email || ''}" autocomplete="email"></label>
        <label>Password<input type="password" name="password" required autocomplete="current-password"></label>
        <button class="btn btn-block">Log in</button>
      </form>
      ${DEMO ? html`<div class="demo-note"><strong>Demo accounts</strong> (password <code>demo1234</code>)<br>Matchmaker: <code>matchmaker@demo.matchstriker.com</code><br>Client: <code>client@demo.matchstriker.com</code></div>` : ''}`,
    });
  };

  app.get('/login', (ctx) => (ctx.user ? ctx.redirect(ctx.user.role === 'matchmaker' ? '/dashboard' : '/account') : loginPage(ctx, ctx.query)));

  app.post('/login', (ctx) => {
    if (throttled(ctx.ip)) return ctx.send(loginPage(ctx, ctx.body, 'Too many attempts. Please wait a few minutes and try again.'), 429);
    const email = clean(ctx.body.email, 120).toLowerCase();
    const user = db.get('SELECT * FROM users WHERE email = ?', email);
    if (!user || !verifyPassword(String(ctx.body.password || ''), user.password_hash)) {
      return ctx.send(loginPage(ctx, ctx.body, 'That email and password do not match.'), 400);
    }
    ctx.login(user.id);
    ctx.redirect(safeNext(ctx.body.next) || (user.role === 'matchmaker' ? '/dashboard' : '/account'));
  });

  app.post('/logout', (ctx) => {
    ctx.logout();
    ctx.redirect('/', 'You are logged out.');
  });
};
