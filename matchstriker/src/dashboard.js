'use strict';
const { db } = require('./db');
const { html, clean, money, toCents, splitTags, fmtDate, firstName } = require('./lib');
const {
  activateMembership, runBilling, createInvoice, addMonth, planTerms, getTags, setTags,
  feePercent, connectionBetween,
} = require('./core');
const v = require('./views');

const MM = 'matchmaker';
const INTRO_STATUSES = ['introduced', 'dating', 'paused', 'success', 'closed'];
const dollars = (cents) => (cents ? (cents / 100).toString() : '');

const listingProblems = (mm) => {
  const p = [];
  if (!mm.headline) p.push('a headline');
  if (!mm.city) p.push('a city');
  if (!mm.bio) p.push('a bio');
  if (!db.get('SELECT 1 AS x FROM plans WHERE matchmaker_id = ? AND active = 1', mm.user_id)) p.push('at least one active plan');
  return p;
};

const myMembership = (ctx, id) =>
  db.get(
    `SELECT ms.*, u.name, u.email, p.name AS plan_name FROM memberships ms JOIN users u ON u.id = ms.client_id JOIN plans p ON p.id = ms.plan_id
     WHERE ms.id = ? AND ms.matchmaker_id = ?`, id, ctx.mm.user_id);

function planForm(ctx, plan) {
  const p = plan || {};
  return html`<form method="post" action="/dashboard/plans${plan ? '/' + plan.id : ''}" class="form">
    ${v.csrf(ctx)}
    <label>Plan name<input name="name" required maxlength="60" value="${p.name || ''}" placeholder="e.g. Free profile, Full search, Pay as you date"></label>
    <label>What the client gets<textarea name="description" rows="2" maxlength="400">${p.description || ''}</textarea></label>
    <div class="grid four">
      <label>Sign-up fee ($)<input name="signup_fee" inputmode="decimal" value="${dollars(p.signup_fee)}" placeholder="0"></label>
      <label>Monthly fee ($)<input name="monthly_fee" inputmode="decimal" value="${dollars(p.monthly_fee)}" placeholder="0"></label>
      <label>Per match ($)<input name="per_match_fee" inputmode="decimal" value="${dollars(p.per_match_fee)}" placeholder="0"></label>
      <label>Per date ($)<input name="per_date_fee" inputmode="decimal" value="${dollars(p.per_date_fee)}" placeholder="0"></label>
    </div>
    <button class="btn ${plan ? 'btn-small' : ''}">${plan ? 'Save changes' : 'Add plan'}</button>
  </form>`;
}
const planFields = (b) => [
  clean(b.name, 60), clean(b.description, 400), toCents(b.signup_fee), toCents(b.monthly_fee), toCents(b.per_match_fee), toCents(b.per_date_fee),
];

const csvCell = (val) => {
  let s = String(val ?? '');
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; // stop spreadsheet formula injection
  return '"' + s.replace(/"/g, '""') + '"';
};

module.exports = (app) => {
  // --- Overview
  app.get('/dashboard', MM, (ctx) => {
    const id = ctx.mm.user_id;
    runBilling('matchmaker', id);
    const counts = db.get(
      `SELECT SUM(status = 'active') AS active, SUM(status = 'pending') AS pending, SUM(status = 'paused') AS paused
       FROM memberships WHERE matchmaker_id = ?`, id);
    const money_ = db.get(
      `SELECT COALESCE(SUM(CASE WHEN status = 'paid' THEN amount END), 0) AS gross,
              COALESCE(SUM(CASE WHEN status = 'paid' THEN fee END), 0) AS fees,
              COALESCE(SUM(CASE WHEN status = 'due' THEN amount END), 0) AS due
       FROM invoices WHERE matchmaker_id = ?`, id);
    const intros = db.get('SELECT COUNT(*) AS n FROM introductions WHERE matchmaker_id = ?', id).n;
    const pending = db.all(
      `SELECT ms.id, ms.created_at, u.name, p.name AS plan_name FROM memberships ms JOIN users u ON u.id = ms.client_id JOIN plans p ON p.id = ms.plan_id
       WHERE ms.matchmaker_id = ? AND ms.status = 'pending' ORDER BY ms.created_at`, id);
    const recent = db.all(
      `SELECT i.*, u.name FROM invoices i JOIN users u ON u.id = i.client_id
       WHERE i.matchmaker_id = ? AND i.status = 'paid' ORDER BY i.paid_at DESC LIMIT 6`, id);
    const problems = listingProblems(ctx.mm);
    const hasPlan = !problems.includes('at least one active plan');
    const listingDone = problems.filter((p) => p !== 'at least one active plan').length === 0;

    return v.dashLayout(ctx, {
      title: 'Dashboard',
      body: html`
<div class="section-head"><h1>${ctx.mm.business_name}</h1>
  ${ctx.mm.published ? html`<a href="/m/${ctx.mm.slug}">View public listing →</a>` : v.statusPill('unpublished')}</div>

${!ctx.mm.published ? html`<div class="card checklist">
  <h3>Get your practice live</h3>
  <ol>
    <li class="${listingDone ? 'done' : ''}"><a href="/dashboard/listing">Fill in your listing</a>: headline, city, bio and the communities you serve.</li>
    <li class="${hasPlan ? 'done' : ''}"><a href="/dashboard/plans">Add a plan</a>: free, sign-up fee, monthly, per match, per date, or a mix.</li>
    <li><a href="/dashboard/listing">Publish</a> so clients can find you.</li>
  </ol>
</div>` : ''}

<div class="grid four stats">
  <div class="stat"><span>Active clients</span><strong>${counts.active || 0}</strong><em>${counts.paused || 0} paused</em></div>
  <div class="stat"><span>Introductions made</span><strong>${intros}</strong><em>all time</em></div>
  <div class="stat"><span>You have earned</span><strong>${money(money_.gross - money_.fees)}</strong><em>${money(money_.gross)} gross, ${money(money_.fees)} platform fee</em></div>
  <div class="stat"><span>Awaiting payment</span><strong>${money(money_.due)}</strong><em>invoiced to clients</em></div>
</div>

<div class="grid two-wide">
  <div>
    <h2>Applications to review</h2>
    ${pending.length ? html`<div class="table-wrap"><table class="table">
      <thead><tr><th>Client</th><th>Plan</th><th>Applied</th><th></th></tr></thead>
      <tbody>${pending.map((p) => html`<tr>
        <td><a href="/dashboard/clients/${p.id}">${p.name}</a></td><td>${p.plan_name}</td><td>${fmtDate(p.created_at)}</td>
        <td class="num"><form method="post" action="/dashboard/clients/${p.id}/status">${v.csrf(ctx)}<input type="hidden" name="status" value="active"><button class="btn btn-small">Accept</button></form></td>
      </tr>`)}</tbody></table></div>`
      : html`<p class="muted">Nothing waiting. ${ctx.mm.requires_approval ? '' : html`New clients join instantly; turn on approval in <a href="/dashboard/listing">your listing</a> if you want to vet them first.`}</p>`}
  </div>
  <div>
    <h2>Recent payments</h2>
    ${recent.length ? html`<div class="table-wrap"><table class="table">
      <tbody>${recent.map((i) => html`<tr><td>${i.name}<div class="small muted">${i.description}</div></td><td class="num">${money(i.amount - i.fee)}<div class="small muted">${fmtDate(i.paid_at)}</div></td></tr>`)}</tbody>
    </table></div><p><a href="/dashboard/payments">All payments</a></p>` : html`<p class="muted">Payments from clients will show here.</p>`}
  </div>
</div>`,
    });
  });

  // --- Listing
  app.get('/dashboard/listing', MM, (ctx) => {
    const m = ctx.mm;
    const tags = getTags(m.user_id);
    const problems = listingProblems(m);
    return v.dashLayout(ctx, {
      title: 'Listing',
      body: html`
<div class="section-head"><h1>Your listing</h1><a href="/m/${m.slug}">Preview →</a></div>
<div class="grid two-wide">
<form method="post" action="/dashboard/listing" class="form">
  ${v.csrf(ctx)}
  <label>Business name<input name="business_name" required maxlength="80" value="${m.business_name}"></label>
  <label>Headline <span class="muted">(one line clients see in search)</span><input name="headline" maxlength="140" value="${m.headline}" placeholder="e.g. Modern Orthodox matchmaking for NYC professionals"></label>
  <div class="grid two">
    <label>City<input name="city" maxlength="80" value="${m.city}" placeholder="e.g. New York, NY"></label>
    <label>Years matchmaking<input type="number" name="years" min="0" max="80" value="${m.years || ''}"></label>
  </div>
  <label>Communities you serve <span class="muted">(comma separated)</span><input name="communities" value="${tags.community.join(', ')}" placeholder="e.g. Jewish, Modern Orthodox, Israeli"></label>
  <label>Languages <span class="muted">(comma separated)</span><input name="languages" value="${tags.language.join(', ')}" placeholder="e.g. English, Hebrew"></label>
  <label>Specialties <span class="muted">(comma separated)</span><input name="specialties" value="${tags.specialty.join(', ')}" placeholder="e.g. Second marriages, 40+, Date coaching"></label>
  <label>About you<textarea name="bio" rows="6" maxlength="3000">${m.bio}</textarea></label>
  <label>How you work <span class="muted">(your process, what clients can expect)</span><textarea name="approach" rows="5" maxlength="3000">${m.approach}</textarea></label>
  <fieldset><legend>Business rules</legend>
    <label class="check"><input type="checkbox" name="accepting" ${m.accepting ? 'checked' : ''}> I am taking new clients</label>
    <label class="check"><input type="checkbox" name="requires_approval" ${m.requires_approval ? 'checked' : ''}> I approve each client before they join <span class="muted">(they are not charged until you accept)</span></label>
  </fieldset>
  <button class="btn">Save listing</button>
</form>
<aside class="card">
  <h3>Visibility</h3>
  <p>${m.published ? html`${v.statusPill('published')} Clients can find you in search.` : html`${v.statusPill('unpublished')} Only you can see your listing.`}</p>
  ${!m.published && problems.length ? html`<p class="small">Before publishing, add ${problems.join(', ')}.</p>` : ''}
  <form method="post" action="/dashboard/listing/publish">${v.csrf(ctx)}
    <input type="hidden" name="published" value="${m.published ? '0' : '1'}">
    <button class="btn ${m.published ? 'btn-ghost' : ''}" ${!m.published && problems.length ? 'disabled' : ''}>${m.published ? 'Unpublish' : 'Publish listing'}</button>
  </form>
  <p class="small muted">Your page: /m/${m.slug}</p>
</aside>
</div>`,
    });
  });

  app.post('/dashboard/listing', MM, (ctx) => {
    const b = ctx.body;
    const id = ctx.mm.user_id;
    let years = parseInt(b.years, 10);
    if (!(years >= 0 && years <= 80)) years = 0;
    db.tx(() => {
      db.run(
        `UPDATE matchmakers SET business_name = ?, headline = ?, city = ?, years = ?, bio = ?, approach = ?, accepting = ?, requires_approval = ? WHERE user_id = ?`,
        clean(b.business_name, 80) || ctx.mm.business_name, clean(b.headline, 140), clean(b.city, 80), years,
        clean(b.bio, 3000), clean(b.approach, 3000), b.accepting ? 1 : 0, b.requires_approval ? 1 : 0, id);
      setTags(id, 'community', splitTags(b.communities));
      setTags(id, 'language', splitTags(b.languages));
      setTags(id, 'specialty', splitTags(b.specialties));
    });
    ctx.redirect('/dashboard/listing', 'Listing saved.');
  });

  app.post('/dashboard/listing/publish', MM, (ctx) => {
    const publish = ctx.body.published === '1';
    const problems = listingProblems(ctx.mm);
    if (publish && problems.length) return ctx.redirect('/dashboard/listing', `Before publishing, add ${problems.join(', ')}.`);
    db.run('UPDATE matchmakers SET published = ? WHERE user_id = ?', publish ? 1 : 0, ctx.mm.user_id);
    ctx.redirect('/dashboard/listing', publish ? 'You are live. Clients can now find you.' : 'Listing unpublished.');
  });

  // --- Plans
  app.get('/dashboard/plans', MM, (ctx) => {
    const plans = db.all(
      `SELECT p.*, (SELECT COUNT(*) FROM memberships ms WHERE ms.plan_id = p.id AND ms.status IN ('active','paused','pending')) AS clients
       FROM plans p WHERE p.matchmaker_id = ? ORDER BY p.active DESC, p.id`, ctx.mm.user_id);
    return v.dashLayout(ctx, {
      title: 'Plans & pricing',
      body: html`
<h1>Plans &amp; pricing</h1>
<p class="lede">Set the terms of your business. Each plan can combine a sign-up fee, a monthly fee, a fee per match and a fee per date. Leave everything at 0 for a free plan. Matchstriker keeps ${feePercent()} of what clients pay.</p>
<div class="grid two-wide">
  <div>
    ${plans.length ? plans.map((p) => html`<div class="card plan-row ${p.active ? '' : 'inactive'}">
      <div class="section-head"><h3>${p.name} ${p.active ? '' : v.statusPill('hidden')}</h3><span class="small muted">${p.clients} ${p.clients === 1 ? 'client' : 'clients'}</span></div>
      <p><strong>${planTerms(p).join(' · ')}</strong></p>
      ${p.description ? html`<p class="small muted">${p.description}</p>` : ''}
      <details><summary>Edit</summary>${planForm(ctx, p)}</details>
      <form method="post" action="/dashboard/plans/${p.id}/toggle">${v.csrf(ctx)}<button class="linklike">${p.active ? 'Hide from new clients' : 'Offer again'}</button></form>
    </div>`) : html`<div class="empty"><h3>No plans yet</h3><p>Add your first plan to start taking clients.</p></div>`}
  </div>
  <aside class="card"><h3>Add a plan</h3>${planForm(ctx)}
    <p class="small muted">Changing a price applies to future charges. Invoices already raised keep their amount.</p></aside>
</div>`,
    });
  });

  app.post('/dashboard/plans', MM, (ctx) => {
    const f = planFields(ctx.body);
    if (!f[0]) return ctx.redirect('/dashboard/plans', 'Give the plan a name.');
    db.run('INSERT INTO plans (matchmaker_id, name, description, signup_fee, monthly_fee, per_match_fee, per_date_fee, created_at) VALUES (?,?,?,?,?,?,?,?)', ctx.mm.user_id, ...f, Date.now());
    ctx.redirect('/dashboard/plans', 'Plan added.');
  });

  app.post('/dashboard/plans/:id', MM, (ctx) => {
    const f = planFields(ctx.body);
    if (!f[0]) return ctx.redirect('/dashboard/plans', 'Give the plan a name.');
    db.run('UPDATE plans SET name = ?, description = ?, signup_fee = ?, monthly_fee = ?, per_match_fee = ?, per_date_fee = ? WHERE id = ? AND matchmaker_id = ?', ...f, ctx.params.id, ctx.mm.user_id);
    ctx.redirect('/dashboard/plans', 'Plan updated.');
  });

  app.post('/dashboard/plans/:id/toggle', MM, (ctx) => {
    db.run('UPDATE plans SET active = 1 - active WHERE id = ? AND matchmaker_id = ?', ctx.params.id, ctx.mm.user_id);
    if (ctx.mm.published && listingProblems(ctx.mm).length) {
      db.run('UPDATE matchmakers SET published = 0 WHERE user_id = ?', ctx.mm.user_id);
      return ctx.redirect('/dashboard/plans', 'Plan hidden. Your listing was unpublished because it has no active plan.');
    }
    ctx.redirect('/dashboard/plans', 'Plan updated.');
  });

  // --- Clients
  const clientRows = (ctx) => {
    const q = clean(ctx.query.q, 60);
    const status = ['pending', 'active', 'paused', 'ended'].includes(ctx.query.status) ? ctx.query.status : '';
    const where = ['ms.matchmaker_id = ?'];
    const params = [ctx.mm.user_id];
    if (status) { where.push('ms.status = ?'); params.push(status); }
    if (q) {
      const like = '%' + q.replace(/[%_\\]/g, '\\$&') + '%';
      where.push("(u.name LIKE ? ESCAPE '\\' OR u.email LIKE ? ESCAPE '\\' OR cp.city LIKE ? ESCAPE '\\' OR ms.notes LIKE ? ESCAPE '\\')");
      params.push(like, like, like, like);
    }
    const rows = db.all(
      `SELECT ms.*, u.name, u.email, p.name AS plan_name, cp.age, cp.city,
        (SELECT COALESCE(SUM(amount), 0) FROM invoices i WHERE i.membership_id = ms.id AND i.status = 'paid') AS paid,
        (SELECT COALESCE(SUM(amount), 0) FROM invoices i WHERE i.membership_id = ms.id AND i.status = 'due') AS due
       FROM memberships ms JOIN users u ON u.id = ms.client_id JOIN plans p ON p.id = ms.plan_id
       LEFT JOIN client_profiles cp ON cp.user_id = ms.client_id
       WHERE ${where.join(' AND ')}
       ORDER BY CASE ms.status WHEN 'pending' THEN 0 WHEN 'active' THEN 1 WHEN 'paused' THEN 2 ELSE 3 END, u.name`, ...params);
    return { rows, q, status };
  };

  app.get('/dashboard/clients', MM, (ctx) => {
    runBilling('matchmaker', ctx.mm.user_id);
    const { rows, q, status } = clientRows(ctx);
    return v.dashLayout(ctx, {
      title: 'Clients',
      body: html`
<div class="section-head"><h1>Clients</h1><a class="btn btn-small btn-ghost" href="/dashboard/clients.csv">Export CSV</a></div>
<form class="filters compact" method="get" action="/dashboard/clients">
  <label>Search<input type="search" name="q" value="${q}" placeholder="Name, email, city, notes"></label>
  <label>Status<select name="status"><option value="">All</option>${v.options(['pending', 'active', 'paused', 'ended'], status)}</select></label>
  <div class="filter-actions"><button class="btn btn-small">Filter</button>${q || status ? html`<a href="/dashboard/clients">Clear</a>` : ''}</div>
</form>
${rows.length ? html`<div class="table-wrap"><table class="table">
  <thead><tr><th>Client</th><th>Status</th><th>Plan</th><th>Age</th><th>City</th><th>Joined</th><th class="num">Paid</th><th class="num">Due</th><th>Notes</th></tr></thead>
  <tbody>${rows.map((r) => html`<tr>
    <td><a href="/dashboard/clients/${r.id}"><strong>${r.name}</strong></a><div class="small muted">${r.email}</div></td>
    <td>${v.statusPill(r.status)}</td><td>${r.plan_name}</td><td>${r.age || '–'}</td><td>${r.city || '–'}</td>
    <td>${fmtDate(r.created_at)}</td><td class="num">${money(r.paid)}</td><td class="num">${r.due ? money(r.due) : '–'}</td>
    <td class="small clip">${r.notes}</td>
  </tr>`)}</tbody></table></div>`
  : html`<div class="empty"><h3>${q || status ? 'No clients match' : 'No clients yet'}</h3><p>${q || status ? 'Try clearing the filter.' : 'When someone joins or applies from your listing they appear here.'}</p></div>`}`,
    });
  });

  app.get('/dashboard/clients.csv', MM, (ctx) => {
    ctx.query = {};
    const { rows } = clientRows(ctx);
    const lines = [['Name', 'Email', 'Status', 'Plan', 'Age', 'City', 'Joined', 'Paid', 'Due', 'Notes'].map(csvCell).join(',')];
    for (const r of rows) {
      lines.push([r.name, r.email, r.status, r.plan_name, r.age ?? '', r.city ?? '', new Date(r.created_at).toISOString().slice(0, 10),
        (r.paid / 100).toFixed(2), (r.due / 100).toFixed(2), r.notes].map(csvCell).join(','));
    }
    ctx.send(lines.join('\r\n') + '\r\n', 200, 'text/csv; charset=utf-8', { 'Content-Disposition': 'attachment; filename="clients.csv"' });
  });

  app.get('/dashboard/clients/:id', MM, (ctx) => {
    const ms = myMembership(ctx, ctx.params.id);
    if (!ms) return ctx.error(404, 'Client not found', '');
    const profile = db.get('SELECT * FROM client_profiles WHERE user_id = ?', ms.client_id) || {};
    const plans = db.all('SELECT * FROM plans WHERE matchmaker_id = ? AND (active = 1 OR id = ?) ORDER BY id', ctx.mm.user_id, ms.plan_id);
    const invoices = db.all('SELECT * FROM invoices WHERE membership_id = ? ORDER BY created_at DESC', ms.id);
    const statusChoices = ms.status === 'pending' ? [['active', 'Accept'], ['ended', 'Decline']]
      : ms.status === 'active' ? [['paused', 'Pause'], ['ended', 'End']]
      : ms.status === 'paused' ? [['active', 'Resume'], ['ended', 'End']] : [];
    return v.dashLayout(ctx, {
      title: ms.name,
      body: html`
<p><a href="/dashboard/clients">← All clients</a></p>
<div class="section-head"><h1>${ms.name} ${v.statusPill(ms.status)}</h1>
  <form method="post" action="/messages/start">${v.csrf(ctx)}<input type="hidden" name="to" value="${ms.client_id}"><button class="btn btn-ghost btn-small">Message</button></form></div>
<div class="grid two-wide">
  <div>
    <div class="card">
      <h3>Intake profile</h3>
      <dl class="facts">
        <dt>Email</dt><dd>${ms.email}</dd>
        <dt>Age</dt><dd>${profile.age || 'Not given'}</dd>
        <dt>City</dt><dd>${profile.city || 'Not given'}</dd>
        <dt>About</dt><dd class="prose">${profile.about || 'Not filled in yet'}</dd>
        <dt>Looking for</dt><dd class="prose">${profile.seeking || 'Not filled in yet'}</dd>
      </dl>
    </div>
    <h2>Invoices</h2>
    ${invoices.length ? html`<div class="table-wrap"><table class="table">
      <thead><tr><th>For</th><th>Date</th><th>Status</th><th class="num">Amount</th><th></th></tr></thead>
      <tbody>${invoices.map((i) => html`<tr><td>${i.description}</td><td>${fmtDate(i.created_at)}</td><td>${v.statusPill(i.status)}</td><td class="num">${money(i.amount)}</td>
        <td class="num">${i.status === 'due' ? html`<form method="post" action="/dashboard/invoices/${i.id}/void">${v.csrf(ctx)}<input type="hidden" name="back" value="/dashboard/clients/${ms.id}"><button class="linklike danger">Cancel</button></form>` : ''}</td></tr>`)}</tbody>
    </table></div>` : html`<p class="muted">No invoices.</p>`}
  </div>
  <aside>
    ${ms.status !== 'ended' ? html`<div class="card">
      <h3>Membership</h3>
      <p class="small muted">Joined ${fmtDate(ms.created_at)}${ms.next_bill_at ? ` · next monthly bill ${fmtDate(ms.next_bill_at)}` : ''}</p>
      <div class="row-actions">${statusChoices.map(([s, label]) => html`<form method="post" action="/dashboard/clients/${ms.id}/status">${v.csrf(ctx)}<input type="hidden" name="status" value="${s}"><button class="btn btn-small ${s === 'active' ? '' : 'btn-ghost'}">${label}</button></form>`)}</div>
      <form method="post" action="/dashboard/clients/${ms.id}/plan" class="form">${v.csrf(ctx)}
        <label>Plan<select name="plan_id">${v.options(plans.map((p) => [p.id, `${p.name} (${planTerms(p).join(', ')})`]), ms.plan_id)}</select></label>
        <button class="btn btn-small btn-ghost">Change plan</button>
      </form>
    </div>` : html`<div class="card"><p>This membership has ended. The client can rejoin from your listing.</p></div>`}
    <div class="card">
      <h3>Private notes</h3>
      <form method="post" action="/dashboard/clients/${ms.id}/notes" class="form">${v.csrf(ctx)}
        <textarea name="notes" rows="5" maxlength="4000" aria-label="Private notes" placeholder="Only you can see this">${ms.notes}</textarea>
        <button class="btn btn-small btn-ghost">Save notes</button>
      </form>
    </div>
    ${ms.status === 'active' || ms.status === 'paused' ? html`<div class="card">
      <h3>One-off charge</h3>
      <form method="post" action="/dashboard/clients/${ms.id}/charge" class="form">${v.csrf(ctx)}
        <label>For<input name="description" required maxlength="120" placeholder="e.g. Date coaching session"></label>
        <label>Amount ($)<input name="amount" required inputmode="decimal"></label>
        <button class="btn btn-small btn-ghost">Send invoice</button>
      </form>
    </div>` : ''}
  </aside>
</div>`,
    });
  });

  app.post('/dashboard/clients/:id/status', MM, (ctx) => {
    const ms = myMembership(ctx, ctx.params.id);
    if (!ms) return ctx.error(404, 'Client not found', '');
    const to = ctx.body.status;
    const back = '/dashboard/clients/' + ms.id;
    const allowed = { pending: ['active', 'ended'], active: ['paused', 'ended'], paused: ['active', 'ended'], ended: [] }[ms.status];
    if (!allowed.includes(to)) return ctx.redirect(back, 'That change is not available.');
    db.tx(() => {
      if (ms.status === 'pending' && to === 'active') {
        activateMembership(ms);
      } else if (ms.status === 'paused' && to === 'active') {
        const plan = db.get('SELECT * FROM plans WHERE id = ?', ms.plan_id);
        createInvoice({ membership: ms, kind: 'monthly', description: `Membership: ${plan.name}`, amount: plan.monthly_fee });
        db.run("UPDATE memberships SET status = 'active', next_bill_at = ? WHERE id = ?", plan.monthly_fee ? addMonth(Date.now()) : null, ms.id);
      } else {
        db.run('UPDATE memberships SET status = ?, next_bill_at = NULL WHERE id = ?', to, ms.id);
      }
    });
    const msg = { active: ms.status === 'pending' ? `${ms.name} accepted. Any joining fees have been invoiced.` : 'Membership resumed.', paused: 'Membership paused. Monthly billing is on hold.', ended: 'Membership ended.' }[to];
    ctx.redirect(ctx.body.back === 'dashboard' ? '/dashboard' : back, msg);
  });

  app.post('/dashboard/clients/:id/plan', MM, (ctx) => {
    const ms = myMembership(ctx, ctx.params.id);
    const plan = ms && db.get('SELECT * FROM plans WHERE id = ? AND matchmaker_id = ?', ctx.body.plan_id, ctx.mm.user_id);
    if (!plan || ms.status === 'ended') return ctx.error(404, 'Not found', '');
    const next = ms.status === 'active' && plan.monthly_fee ? (ms.next_bill_at || Date.now()) : null;
    db.run('UPDATE memberships SET plan_id = ?, next_bill_at = ? WHERE id = ?', plan.id, next, ms.id);
    ctx.redirect('/dashboard/clients/' + ms.id, `Plan changed to ${plan.name}.`);
  });

  app.post('/dashboard/clients/:id/notes', MM, (ctx) => {
    const ms = myMembership(ctx, ctx.params.id);
    if (!ms) return ctx.error(404, 'Client not found', '');
    db.run('UPDATE memberships SET notes = ? WHERE id = ?', clean(ctx.body.notes, 4000), ms.id);
    ctx.redirect('/dashboard/clients/' + ms.id, 'Notes saved.');
  });

  app.post('/dashboard/clients/:id/charge', MM, (ctx) => {
    const ms = myMembership(ctx, ctx.params.id);
    if (!ms || !['active', 'paused'].includes(ms.status)) return ctx.error(404, 'Client not found', '');
    const amount = toCents(ctx.body.amount);
    const description = clean(ctx.body.description, 120);
    if (!amount || !description) return ctx.redirect('/dashboard/clients/' + ms.id, 'Enter a description and an amount above zero.');
    createInvoice({ membership: ms, kind: 'custom', description, amount });
    ctx.redirect('/dashboard/clients/' + ms.id, `Invoice for ${money(amount)} sent.`);
  });

  // --- Matches (introductions) and dates
  app.get('/dashboard/matches', MM, (ctx) => {
    const id = ctx.mm.user_id;
    const roster = db.all(
      `SELECT ms.id, u.name, p.per_match_fee FROM memberships ms JOIN users u ON u.id = ms.client_id JOIN plans p ON p.id = ms.plan_id
       WHERE ms.matchmaker_id = ? AND ms.status IN ('active','paused') ORDER BY u.name`, id);
    const intros = db.all(
      `SELECT i.*, ua.name AS a_name, ub.name AS b_name FROM introductions i
       JOIN memberships ma ON ma.id = i.a_membership_id JOIN users ua ON ua.id = ma.client_id
       LEFT JOIN memberships mb ON mb.id = i.b_membership_id LEFT JOIN users ub ON ub.id = mb.client_id
       WHERE i.matchmaker_id = ? ORDER BY i.created_at DESC`, id);
    const rosterOpts = roster.map((r) => [r.id, r.per_match_fee ? `${r.name} (${money(r.per_match_fee)} match fee)` : r.name]);
    return v.dashLayout(ctx, {
      title: 'Matches',
      body: html`
<h1>Matches</h1>
<div class="grid two-wide">
  <div>
    ${intros.length ? intros.map((i) => {
      const dates = db.all('SELECT * FROM dates WHERE introduction_id = ? ORDER BY happens_on', i.id);
      return html`<div class="card match">
        <div class="section-head"><h3>${i.a_name} <span class="amp">&amp;</span> ${i.b_name || i.b_external}${i.b_name ? '' : html` <span class="small muted">(not a client)</span>`}</h3>${v.statusPill(i.status)}</div>
        <p class="small muted">Introduced ${fmtDate(i.created_at)}</p>
        ${i.note ? html`<p class="small">${i.note}</p>` : ''}
        ${dates.length ? html`<ul class="dates">${dates.map((d) => html`<li>${d.happens_on}${d.venue ? ` · ${d.venue}` : ''}</li>`)}</ul>` : ''}
        <details><summary>Log a date or update status</summary>
          <form method="post" action="/dashboard/matches/${i.id}/dates" class="form inline-form">${v.csrf(ctx)}
            <label>Date<input type="date" name="happens_on" required></label>
            <label>Where<input name="venue" maxlength="120" placeholder="Optional"></label>
            <button class="btn btn-small">Log date</button>
          </form>
          <form method="post" action="/dashboard/matches/${i.id}/status" class="form inline-form">${v.csrf(ctx)}
            <label>Status<select name="status">${v.options(INTRO_STATUSES, i.status)}</select></label>
            <button class="btn btn-small btn-ghost">Update</button>
          </form>
        </details>
      </div>`;
    }) : html`<div class="empty"><h3>No matches yet</h3><p>Introduce two of your clients, or a client and someone outside your roster.</p></div>`}
  </div>
  <aside class="card">
    <h3>Make a match</h3>
    ${roster.length ? html`<form method="post" action="/dashboard/matches" class="form">${v.csrf(ctx)}
      <label>Client<select name="a" required><option value="">Choose…</option>${v.options(rosterOpts)}</select></label>
      <label>Introduce to<select name="b"><option value="">Someone outside my roster</option>${v.options(rosterOpts)}</select></label>
      <label>Their name <span class="muted">(if outside your roster)</span><input name="b_external" maxlength="80"></label>
      <label>Note to the client(s)<textarea name="note" rows="3" maxlength="1000" placeholder="Why you think they will get along"></textarea></label>
      <button class="btn">Create match</button>
    </form>
    <p class="small muted">Clients on a plan with a per-match fee are invoiced when you create the match, and per-date fees when you log a date.</p>`
    : html`<p class="muted">You need at least one active client to make a match.</p>`}
  </aside>
</div>`,
    });
  });

  const rosterMember = (ctx, id) =>
    db.get(
      `SELECT ms.*, p.per_match_fee, p.per_date_fee, u.name FROM memberships ms JOIN plans p ON p.id = ms.plan_id JOIN users u ON u.id = ms.client_id
       WHERE ms.id = ? AND ms.matchmaker_id = ? AND ms.status IN ('active','paused')`, id, ctx.mm.user_id);

  app.post('/dashboard/matches', MM, (ctx) => {
    const a = rosterMember(ctx, ctx.body.a);
    const b = ctx.body.b ? rosterMember(ctx, ctx.body.b) : null;
    const external = clean(ctx.body.b_external, 80);
    if (!a) return ctx.redirect('/dashboard/matches', 'Choose a client to match.');
    if (ctx.body.b && !b) return ctx.redirect('/dashboard/matches', 'That second client is not on your active roster.');
    if (b && b.id === a.id) return ctx.redirect('/dashboard/matches', 'Choose two different people.');
    if (!b && !external) return ctx.redirect('/dashboard/matches', 'Choose a second client or enter a name.');
    db.tx(() => {
      db.run('INSERT INTO introductions (matchmaker_id, a_membership_id, b_membership_id, b_external, note, created_at) VALUES (?,?,?,?,?,?)',
        ctx.mm.user_id, a.id, b ? b.id : null, b ? '' : external, clean(ctx.body.note, 1000), Date.now());
      createInvoice({ membership: a, kind: 'match', description: `Match fee: introduction to ${firstName(b ? b.name : external)}`, amount: a.per_match_fee });
      if (b) createInvoice({ membership: b, kind: 'match', description: `Match fee: introduction to ${firstName(a.name)}`, amount: b.per_match_fee });
    });
    ctx.redirect('/dashboard/matches', 'Match created.');
  });

  const myIntro = (ctx) => db.get('SELECT * FROM introductions WHERE id = ? AND matchmaker_id = ?', ctx.params.id, ctx.mm.user_id);

  app.post('/dashboard/matches/:id/status', MM, (ctx) => {
    const intro = myIntro(ctx);
    if (!intro || !INTRO_STATUSES.includes(ctx.body.status)) return ctx.error(404, 'Match not found', '');
    db.run('UPDATE introductions SET status = ? WHERE id = ?', ctx.body.status, intro.id);
    ctx.redirect('/dashboard/matches', 'Match updated.');
  });

  app.post('/dashboard/matches/:id/dates', MM, (ctx) => {
    const intro = myIntro(ctx);
    if (!intro) return ctx.error(404, 'Match not found', '');
    const on = clean(ctx.body.happens_on, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(on) || Number.isNaN(Date.parse(on))) return ctx.redirect('/dashboard/matches', 'Enter a valid date.');
    db.tx(() => {
      db.run('INSERT INTO dates (introduction_id, happens_on, venue, created_at) VALUES (?,?,?,?)', intro.id, on, clean(ctx.body.venue, 120), Date.now());
      if (intro.status === 'introduced') db.run("UPDATE introductions SET status = 'dating' WHERE id = ?", intro.id);
      for (const mid of [intro.a_membership_id, intro.b_membership_id]) {
        const ms = mid && rosterMember(ctx, mid);
        if (ms) createInvoice({ membership: ms, kind: 'date', description: `Date fee: ${on}`, amount: ms.per_date_fee });
      }
    });
    ctx.redirect('/dashboard/matches', 'Date logged.');
  });

  // --- Payments
  app.get('/dashboard/payments', MM, (ctx) => {
    const id = ctx.mm.user_id;
    runBilling('matchmaker', id);
    const rows = db.all(
      `SELECT i.*, u.name FROM invoices i JOIN users u ON u.id = i.client_id WHERE i.matchmaker_id = ? AND i.status != 'void'
       ORDER BY i.created_at DESC LIMIT 200`, id);
    const t = db.get(
      `SELECT COALESCE(SUM(CASE WHEN status = 'paid' THEN amount END), 0) AS gross,
              COALESCE(SUM(CASE WHEN status = 'paid' THEN fee END), 0) AS fees,
              COALESCE(SUM(CASE WHEN status = 'due' THEN amount END), 0) AS due
       FROM invoices WHERE matchmaker_id = ?`, id);
    return v.dashLayout(ctx, {
      title: 'Payments',
      body: html`
<h1>Payments</h1>
<div class="grid four stats">
  <div class="stat"><span>Clients paid</span><strong>${money(t.gross)}</strong><em>gross</em></div>
  <div class="stat"><span>Platform fee (${feePercent()})</span><strong>${money(t.fees)}</strong><em>kept by Matchstriker</em></div>
  <div class="stat"><span>Your earnings</span><strong>${money(t.gross - t.fees)}</strong><em>net</em></div>
  <div class="stat"><span>Awaiting payment</span><strong>${money(t.due)}</strong><em>invoiced, not yet paid</em></div>
</div>
${v.demoNote()}
${rows.length ? html`<div class="table-wrap"><table class="table">
  <thead><tr><th>Date</th><th>Client</th><th>For</th><th>Status</th><th class="num">Amount</th><th class="num">Fee</th><th class="num">You get</th><th></th></tr></thead>
  <tbody>${rows.map((i) => html`<tr>
    <td>${fmtDate(i.created_at)}</td><td><a href="/dashboard/clients/${i.membership_id}">${i.name}</a></td><td>${i.description}</td><td>${v.statusPill(i.status)}</td>
    <td class="num">${money(i.amount)}</td><td class="num muted">−${money(i.fee)}</td><td class="num"><strong>${money(i.amount - i.fee)}</strong></td>
    <td class="num">${i.status === 'due' ? html`<form method="post" action="/dashboard/invoices/${i.id}/void">${v.csrf(ctx)}<button class="linklike danger">Cancel</button></form>` : ''}</td>
  </tr>`)}</tbody></table></div>`
  : html`<div class="empty"><h3>No payments yet</h3><p>Invoices appear here as clients join, get matched and go on dates.</p></div>`}`,
    });
  });

  app.post('/dashboard/invoices/:id/void', MM, (ctx) => {
    const r = db.run("UPDATE invoices SET status = 'void' WHERE id = ? AND matchmaker_id = ? AND status = 'due'", ctx.params.id, ctx.mm.user_id);
    const back = /^\/dashboard\/clients\/\d+$/.test(ctx.body.back || '') ? ctx.body.back : '/dashboard/payments';
    ctx.redirect(back, r.changes ? 'Invoice cancelled.' : 'Only unpaid invoices can be cancelled.');
  });

  // --- Network
  app.get('/dashboard/network', MM, (ctx) => {
    const id = ctx.mm.user_id;
    const conns = db.all(
      `SELECT c.*, m.user_id, m.business_name, m.slug, m.city, m.hue, m.headline FROM connections c
       JOIN matchmakers m ON m.user_id = CASE WHEN c.from_id = ? THEN c.to_id ELSE c.from_id END
       WHERE c.from_id = ? OR c.to_id = ? ORDER BY c.created_at DESC`, id, id, id);
    const connected = conns.filter((c) => c.status === 'accepted');
    const incoming = conns.filter((c) => c.status === 'pending' && c.to_id === id);
    const sent = conns.filter((c) => c.status === 'pending' && c.from_id === id);
    const known = new Set(conns.map((c) => c.user_id));
    const others = db.all('SELECT * FROM matchmakers WHERE published = 1 AND user_id != ? ORDER BY business_name', id).filter((m) => !known.has(m.user_id));
    const row = (m, action) => html`<div class="net-row">${v.avatar(m.business_name, m.hue)}
      <div class="conv-main"><a href="/m/${m.slug}"><strong>${m.business_name}</strong></a><span class="muted small">${m.city}${m.headline ? ` · ${m.headline}` : ''}</span></div>
      <div class="row-actions">${action}</div></div>`;
    const post = (action, label, cls = '', hidden) => html`<form method="post" action="${action}">${v.csrf(ctx)}${hidden || ''}<button class="${cls || 'btn btn-small'}">${label}</button></form>`;
    return v.dashLayout(ctx, {
      title: 'Network',
      body: html`
<h1>Network</h1>
<p class="lede">Connect with other independent matchmakers to trade referrals when a client fits someone else's community better than yours.</p>
${incoming.length ? html`<h2>Requests to connect</h2><div class="card list">${incoming.map((c) => row(c, html`${post(`/dashboard/network/${c.id}/accept`, 'Accept')}${post(`/dashboard/network/${c.id}/remove`, 'Decline', 'linklike')}`))}</div>` : ''}
<h2>Your connections</h2>
${connected.length ? html`<div class="card list">${connected.map((c) => row(c, html`${post('/messages/start', 'Message', 'btn btn-small btn-ghost', html`<input type="hidden" name="to" value="${c.user_id}">`)}${post(`/dashboard/network/${c.id}/remove`, 'Remove', 'linklike danger')}`))}</div>`
  : html`<p class="muted">No connections yet. Send a request below.</p>`}
${sent.length ? html`<h2>Requests you sent</h2><div class="card list">${sent.map((c) => row(c, html`<span class="small muted">Pending</span>${post(`/dashboard/network/${c.id}/remove`, 'Withdraw', 'linklike')}`))}</div>` : ''}
<h2>Other matchmakers</h2>
${others.length ? html`<div class="card list">${others.map((m) => row(m, post('/dashboard/network/connect', 'Connect', 'btn btn-small btn-ghost', html`<input type="hidden" name="to" value="${m.user_id}">`)))}</div>` : html`<p class="muted">You are connected with everyone listed.</p>`}`,
    });
  });

  app.post('/dashboard/network/connect', MM, (ctx) => {
    const to = db.get('SELECT user_id FROM matchmakers WHERE user_id = ? AND published = 1', ctx.body.to);
    if (!to || to.user_id === ctx.mm.user_id) return ctx.error(404, 'Matchmaker not found', '');
    if (!connectionBetween(ctx.mm.user_id, to.user_id)) {
      db.run("INSERT INTO connections (from_id, to_id, status, created_at) VALUES (?,?,'pending',?)", ctx.mm.user_id, to.user_id, Date.now());
    }
    ctx.redirect('/dashboard/network', 'Request sent.');
  });

  app.post('/dashboard/network/:id/accept', MM, (ctx) => {
    db.run("UPDATE connections SET status = 'accepted' WHERE id = ? AND to_id = ? AND status = 'pending'", ctx.params.id, ctx.mm.user_id);
    ctx.redirect('/dashboard/network', 'Connected.');
  });

  app.post('/dashboard/network/:id/remove', MM, (ctx) => {
    db.run('DELETE FROM connections WHERE id = ? AND (from_id = ? OR to_id = ?)', ctx.params.id, ctx.mm.user_id, ctx.mm.user_id);
    ctx.redirect('/dashboard/network', 'Removed.');
  });
};
