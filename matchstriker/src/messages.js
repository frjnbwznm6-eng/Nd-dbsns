'use strict';
const { db } = require('./db');
const { html, clean, fmtTime } = require('./lib');
const { getConversation, sendMessage, displayName, connectionBetween } = require('./core');
const v = require('./views');

// Who may open a new conversation with whom:
//  client -> any published matchmaker
//  matchmaker -> their own clients, and matchmakers they are connected with
function mayStart(user, targetId) {
  const target = db.get('SELECT u.id, u.role, m.published FROM users u LEFT JOIN matchmakers m ON m.user_id = u.id WHERE u.id = ?', targetId);
  if (!target || target.id === user.id) return false;
  if (user.role === 'client') return target.role === 'matchmaker' && target.published === 1;
  if (target.role === 'client') return !!db.get('SELECT 1 AS x FROM memberships WHERE client_id = ? AND matchmaker_id = ?', target.id, user.id);
  const c = connectionBetween(user.id, target.id);
  return !!c && c.status === 'accepted';
}

const myConversation = (ctx) => {
  const c = db.get('SELECT * FROM conversations WHERE id = ? AND (user_a = ? OR user_b = ?)', ctx.params.id, ctx.user.id, ctx.user.id);
  if (!c) ctx.error(404, 'Conversation not found', '');
  return c;
};

function thread(ctx, conv) {
  const msgs = db.all('SELECT * FROM messages WHERE conversation_id = ? ORDER BY id DESC LIMIT 200', conv.id).reverse();
  db.run('UPDATE messages SET read_at = ? WHERE conversation_id = ? AND sender_id != ? AND read_at IS NULL', Date.now(), conv.id, ctx.user.id);
  return msgs.length
    ? msgs.map((m) => html`<div class="msg ${m.sender_id === ctx.user.id ? 'mine' : ''}"><p>${m.body}</p><time>${fmtTime(m.created_at)}</time></div>`)
    : html`<p class="muted center">No messages yet. Say hello.</p>`;
}

module.exports = (app) => {
  app.get('/messages', 'user', (ctx) => {
    const uid = ctx.user.id;
    const convs = db.all(
      `SELECT c.*, (SELECT body FROM messages WHERE conversation_id = c.id ORDER BY id DESC LIMIT 1) AS last_body,
        (SELECT COUNT(*) FROM messages WHERE conversation_id = c.id AND sender_id != ? AND read_at IS NULL) AS unread
       FROM conversations c WHERE c.user_a = ? OR c.user_b = ? ORDER BY c.last_at DESC`, uid, uid, uid);
    return v.layout(ctx, {
      title: 'Messages',
      body: html`<section class="wrap narrow pad">
        <h1>Messages</h1>
        ${convs.length ? html`<div class="conv-list">${convs.map((c) => {
          const other = displayName(c.user_a === uid ? c.user_b : c.user_a);
          return html`<a class="conv ${c.unread ? 'unread' : ''}" href="/messages/${c.id}">
            ${v.avatar(other.name, other.hue)}
            <div class="conv-main"><strong>${other.name}</strong><span class="muted">${c.last_body || 'No messages yet'}</span></div>
            <div class="conv-meta"><time class="small muted">${fmtTime(c.last_at)}</time>${c.unread ? html`<span class="badge">${c.unread}</span>` : ''}</div>
          </a>`;
        })}</div>`
        : html`<div class="empty"><h3>No conversations yet</h3><p>${ctx.user.role === 'client'
            ? html`Open a matchmaker's page and press Message to start one. <a href="/matchmakers">Browse matchmakers</a>.`
            : 'Messages from clients and prospects will land here. You can also message a client from their row in your client table.'}</p></div>`}
      </section>`,
    });
  });

  app.post('/messages/start', 'user', (ctx) => {
    const to = parseInt(ctx.body.to, 10);
    if (!mayStart(ctx.user, to)) return ctx.error(403, 'Cannot start that conversation', 'You can message matchmakers listed on the site, your own clients, and matchmakers in your network.');
    ctx.redirect('/messages/' + getConversation(ctx.user.id, to).id);
  });

  app.get('/messages/:id', 'user', (ctx) => {
    const conv = myConversation(ctx);
    if (!conv) return;
    if (ctx.query.partial) return ctx.send(html`${thread(ctx, conv)}`);
    const otherId = conv.user_a === ctx.user.id ? conv.user_b : conv.user_a;
    const other = displayName(otherId);
    return v.layout(ctx, {
      title: `Messages · ${other.name}`,
      body: html`<section class="wrap narrow pad">
        <p><a href="/messages">← All messages</a></p>
        <div class="thread-head">${v.avatar(other.name, other.hue)}<div><h1>${other.name}</h1>
          ${other.slug ? html`<a class="small" href="/m/${other.slug}">View listing</a>` : html`<span class="small muted">Client</span>`}</div></div>
        <div class="thread" id="thread" data-poll="/messages/${conv.id}?partial=1">${thread(ctx, conv)}</div>
        <form method="post" action="/messages/${conv.id}" class="composer">
          ${v.csrf(ctx)}
          <textarea name="body" rows="2" maxlength="4000" required placeholder="Write a message…" aria-label="Message"></textarea>
          <button class="btn">Send</button>
        </form>
      </section>`,
    });
  });

  app.post('/messages/:id', 'user', (ctx) => {
    const conv = myConversation(ctx);
    if (!conv) return;
    const body = clean(ctx.body.body, 4000);
    if (body) sendMessage(conv, ctx.user.id, body);
    ctx.redirect('/messages/' + conv.id);
  });
};
