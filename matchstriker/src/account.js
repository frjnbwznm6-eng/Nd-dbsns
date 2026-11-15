'use strict';
const { db } = require('./db');
const { html, clean, money, fmtDate, firstName } = require('./lib');
const { activateMembership, runBilling, planTerms } = require('./core');
const v = require('./views');

module.exports = (app) => {
  // --- Joining a matchmaker
  const loadJoin = (ctx) => {
    const m = db.get('SELECT * FROM matchmakers WHERE slug = ? AND published = 1', ctx.params.slug);
    const plan = m && db.get('SELECT * FROM plans WHERE id = ? AND matchmaker_id = ? AND active = 1', ctx.params.plan, m.user_id);
    if (!plan) { ctx.error(404, 'Plan not found', 'That plan is no longer offered.'); return null; }
    const existing = db.get('SELECT * FROM memberships WHERE client_id = ? AND matchmaker_id = ?', ctx.user.id, m.user_id);
    if (existing && existing.status !== 'ended') { ctx.redirect('/account', `You are already with ${m.business_name}.`); return null; }
    return { m, plan, existing };
  };

  app.get('/m/:slug/join/:plan', 'client', (ctx) => {
    const j = loadJoin(ctx);
    if (!j) return;
    const { m, plan } = j;
    const dueNow = plan.signup_fee + plan.monthly_fee;
    return v.layout(ctx, {
      title: `Join ${m.business_name}`,
      body: html`<section class="wrap narrow pad">
        <p><a href="/m/${m.slug}">← Back to ${m.business_name}</a></p>
        <h1>${m.requires_approval ? 'Apply to' : 'Join'} ${m.business_name}</h1>
        <div class="card checkout">
          <h3>${plan.name}</h3>
          ${plan.description ? html`<p class="muted">${plan.description}</p>` : ''}
          <table class="table">
            <tbody>
              ${plan.signup_fee ? html`<tr><td>Sign-up fee (one time)</td><td class="num">${money(plan.signup_fee)}</td></tr>` : ''}
              ${plan.monthly_fee ? html`<tr><td>Membership, first month (then monthly)</td><td class="num">${money(plan.monthly_fee)}</td></tr>` : ''}
              ${plan.per_match_fee ? html`<tr><td>Each match you are introduced to</td><td class="num">${money(plan.per_match_fee)} later</td></tr>` : ''}
              ${plan.per_date_fee ? html`<tr><td>Each date arranged for you</td><td class="num">${money(plan.per_date_fee)} later</td></tr>` : ''}
              <tr class="total"><td>${m.requires_approval ? 'Due if accepted' : 'Due today'}</td><td class="num">${money(dueNow)}</td></tr>
            </tbody>
          </table>
          ${m.requires_approval ? html`<p class="small muted">${m.business_name} reviews applications first. You are not charged anything until you are accepted.</p>` : ''}
          <form method="post" action="/m/${m.slug}/join/${plan.id}">
            ${v.csrf(ctx)}
            <button class="btn btn-block">${m.requires_approval ? 'Send application' : dueNow ? `Pay ${money(dueNow)} and join` : 'Join for free'}</button>
          </form>
          ${dueNow && !m.requires_approval ? v.demoNote() : ''}
        </div>
        <p class="small muted">Your intake profile (age, city, what you are looking for) is shared with matchmakers you join. <a href="/account/profile">Review it</a>.</p>
      </section>`,
    });
  });

  app.post('/m/:slug/join/:plan', 'client', (ctx) => {
    const j = loadJoin(ctx);
    if (!j) return;
    const { m, plan, existing } = j;
    db.tx(() => {
      let id;
      if (existing) {
        id = existing.id;
        db.run("UPDATE memberships SET plan_id = ?, status = 'pending', next_bill_at = NULL WHERE id = ?", plan.id, id);
      } else {
        id = db.insert("INSERT INTO memberships (client_id, matchmaker_id, plan_id, status, created_at) VALUES (?,?,?,'pending',?)", ctx.user.id, m.user_id, plan.id, Date.now());
      }
      if (!m.requires_approval) activateMembership(db.get('SELECT * FROM memberships WHERE id = ?', id), { paid: true });
    });
    ctx.redirect('/account', m.requires_approval
      ? `Application sent to ${m.business_name}. You will see it here once they respond.`
      : `You have joined ${m.business_name}. Say hello with a message.`);
  });

  // --- Client account
  app.get('/account', 'client', (ctx) => {
    const uid = ctx.user.id;
    runBilling('client', uid);
    const memberships = db.all(
      `SELECT ms.*, m.business_name, m.slug, m.hue, p.name AS plan_name, p.signup_fee, p.monthly_fee, p.per_match_fee, p.per_date_fee
       FROM memberships ms JOIN matchmakers m ON m.user_id = ms.matchmaker_id JOIN plans p ON p.id = ms.plan_id
       WHERE ms.client_id = ? AND ms.status != 'ended' ORDER BY ms.created_at DESC`, uid);
    const due = db.all(
      `SELECT i.*, m.business_name FROM invoices i JOIN matchmakers m ON m.user_id = i.matchmaker_id
       WHERE i.client_id = ? AND i.status = 'due' ORDER BY i.created_at`, uid);
    const paid = db.all(
      `SELECT i.*, m.business_name FROM invoices i JOIN matchmakers m ON m.user_id = i.matchmaker_id
       WHERE i.client_id = ? AND i.status = 'paid' ORDER BY i.paid_at DESC LIMIT 25`, uid);
    const intros = db.all(
      `SELECT i.*, m.business_name, ma.client_id AS a_client, ua.name AS a_name, ub.name AS b_name
       FROM introductions i JOIN matchmakers m ON m.user_id = i.matchmaker_id
       JOIN memberships ma ON ma.id = i.a_membership_id JOIN users ua ON ua.id = ma.client_id
       LEFT JOIN memberships mb ON mb.id = i.b_membership_id LEFT JOIN users ub ON ub.id = mb.client_id
       WHERE ma.client_id = ? OR mb.client_id = ? ORDER BY i.created_at DESC`, uid, uid);
    const profile = db.get('SELECT * FROM client_profiles WHERE user_id = ?', uid);
    const profileDone = profile && profile.about && profile.city;

    return v.layout(ctx, {
      title: 'My account',
      body: html`<section class="wrap pad">
  <h1>Hello, ${firstName(ctx.user.name)}</h1>
  ${!profileDone ? html`<div class="notice">Matchmakers work from your intake profile. <a href="/account/profile">Fill it in</a> so they can start.</div>` : ''}

  ${due.length ? html`
  <h2>Payments due</h2>
  <div class="table-wrap"><table class="table">
    <thead><tr><th>For</th><th>Matchmaker</th><th>Date</th><th class="num">Amount</th><th></th></tr></thead>
    <tbody>${due.map((i) => html`<tr>
      <td>${i.description}</td><td>${i.business_name}</td><td>${fmtDate(i.created_at)}</td><td class="num">${money(i.amount)}</td>
      <td class="num"><form method="post" action="/account/invoices/${i.id}/pay">${v.csrf(ctx)}<button class="btn btn-small">Pay ${money(i.amount)}</button></form></td>
    </tr>`)}</tbody>
  </table></div>${v.demoNote()}` : ''}

  <div class="section-head"><h2>My matchmakers</h2><a href="/matchmakers">Find another</a></div>
  ${memberships.length ? html`<div class="grid cards">${memberships.map((ms) => html`<div class="card">
      <div class="mm-card-head">${v.avatar(ms.business_name, ms.hue)}<div><h3><a href="/m/${ms.slug}">${ms.business_name}</a></h3><p class="small muted">Since ${fmtDate(ms.created_at)}</p></div></div>
      <p>${v.statusPill(ms.status)} <strong>${ms.plan_name}</strong></p>
      <p class="small muted">${planTerms(ms).join(' · ')}</p>
      ${ms.status === 'pending' ? html`<p class="small">Waiting for the matchmaker to review your application.</p>` : ''}
      <div class="row-actions">
        <form method="post" action="/messages/start">${v.csrf(ctx)}<input type="hidden" name="to" value="${ms.matchmaker_id}"><button class="btn btn-small btn-ghost">Message</button></form>
        <form method="post" action="/account/memberships/${ms.id}/leave" data-confirm="Leave ${ms.business_name}? Recurring billing stops; unpaid invoices stay due.">${v.csrf(ctx)}<button class="linklike danger">${ms.status === 'pending' ? 'Withdraw' : 'Leave'}</button></form>
      </div>
    </div>`)}</div>`
    : html`<div class="empty"><h3>You have not joined a matchmaker yet</h3><p><a class="btn" href="/matchmakers">Find a matchmaker</a></p></div>`}

  <h2>My introductions</h2>
  ${intros.length ? html`<div class="table-wrap"><table class="table">
    <thead><tr><th>Introduced to</th><th>By</th><th>Status</th><th>Dates</th><th>Note from your matchmaker</th></tr></thead>
    <tbody>${intros.map((i) => {
      const other = i.a_client === uid ? (i.b_name || i.b_external) : i.a_name;
      const dates = db.all('SELECT * FROM dates WHERE introduction_id = ? ORDER BY happens_on', i.id);
      return html`<tr><td><strong>${firstName(other)}</strong></td><td>${i.business_name}</td><td>${v.statusPill(i.status)}</td>
        <td>${dates.length ? dates.map((d) => html`<div class="small">${d.happens_on}${d.venue ? ` · ${d.venue}` : ''}</div>`) : html`<span class="muted small">None yet</span>`}</td>
        <td class="small">${i.note}</td></tr>`;
    })}</tbody></table></div>`
    : html`<p class="muted">Introductions your matchmakers make will appear here.</p>`}

  ${paid.length ? html`<h2>Payment history</h2>
  <div class="table-wrap"><table class="table">
    <thead><tr><th>For</th><th>Matchmaker</th><th>Paid</th><th class="num">Amount</th></tr></thead>
    <tbody>${paid.map((i) => html`<tr><td>${i.description}</td><td>${i.business_name}</td><td>${fmtDate(i.paid_at)}</td><td class="num">${money(i.amount)}</td></tr>`)}</tbody>
  </table></div>` : ''}

  <p><a href="/account/profile">Edit my intake profile</a></p>
</section>`,
    });
  });

  app.get('/account/profile', 'client', (ctx) => {
    const p = db.get('SELECT * FROM client_profiles WHERE user_id = ?', ctx.user.id) || {};
    return v.layout(ctx, {
      title: 'My intake profile',
      body: html`<section class="wrap narrow pad">
        <h1>My intake profile</h1>
        <p class="muted">Only matchmakers you join or apply to can see this. It is never shown publicly.</p>
        <form method="post" action="/account/profile" class="form">
          ${v.csrf(ctx)}
          <label>Name<input name="name" required maxlength="80" value="${ctx.user.name}"></label>
          <div class="grid two">
            <label>Age<input type="number" name="age" min="18" max="110" value="${p.age || ''}"></label>
            <label>City<input name="city" maxlength="80" value="${p.city || ''}" placeholder="e.g. Brooklyn, NY"></label>
          </div>
          <label>About me<textarea name="about" rows="5" maxlength="2000" placeholder="Background, work, faith and family, what a good weekend looks like">${p.about || ''}</textarea></label>
          <label>Who I am looking for<textarea name="seeking" rows="4" maxlength="2000" placeholder="What matters most in a partner, and any dealbreakers">${p.seeking || ''}</textarea></label>
          <button class="btn">Save profile</button>
        </form>
      </section>`,
    });
  });

  app.post('/account/profile', 'client', (ctx) => {
    const b = ctx.body;
    const name = clean(b.name, 80) || ctx.user.name;
    let age = parseInt(b.age, 10);
    if (!(age >= 18 && age <= 110)) age = null;
    db.run('UPDATE users SET name = ? WHERE id = ?', name, ctx.user.id);
    db.run(
      `INSERT INTO client_profiles (user_id, age, city, about, seeking) VALUES (?,?,?,?,?)
       ON CONFLICT(user_id) DO UPDATE SET age = excluded.age, city = excluded.city, about = excluded.about, seeking = excluded.seeking`,
      ctx.user.id, age, clean(b.city, 80), clean(b.about, 2000), clean(b.seeking, 2000));
    ctx.redirect('/account', 'Profile saved.');
  });

  app.post('/account/invoices/:id/pay', 'client', (ctx) => {
    // Demo payment. In production this is where the payment processor is called;
    // the invoice is only marked paid from the processor's confirmed webhook.
    const r = db.run("UPDATE invoices SET status = 'paid', paid_at = ? WHERE id = ? AND client_id = ? AND status = 'due'", Date.now(), ctx.params.id, ctx.user.id);
    ctx.redirect('/account', r.changes ? 'Payment recorded. Thank you.' : 'That invoice is not payable.');
  });

  app.post('/account/memberships/:id/leave', 'client', (ctx) => {
    const r = db.run("UPDATE memberships SET status = 'ended', next_bill_at = NULL WHERE id = ? AND client_id = ? AND status != 'ended'", ctx.params.id, ctx.user.id);
    ctx.redirect('/account', r.changes ? 'You have left that matchmaker.' : 'Membership not found.');
  });
};
