/* FFCSSTRAT accounts API (Vercel serverless function, POST /api/ffcs  {action, ...}).
   Storage: Upstash Redis REST (env KV_REST_API_URL/KV_REST_API_TOKEN or UPSTASH_REDIS_REST_URL/UPSTASH_REDIS_REST_TOKEN).
   For local testing only: FFCS_FILE_STORE=/path/to/store.json uses a JSON file instead.

   Keys:  ffcs:users        {id: {id, name, login, admin, active, salt, hash, created}}
          ffcs:dist         {targetId: sourceId}   faculty-tag distribution chosen by an admin
          ffcs:data:<id>    {data, rev, at}        one user's planner state (tags, queue, courses, session)
          ffcs:sess:<token> userId                  (expires after 30 days)
          ffcs:fail:<login> failed sign-in count    (expires after 15 minutes)

   Every user's data lives under its own key and is written only by that user, except when an admin
   distributes tags: the source's tags are COPIED into the target's data (the source is only read). */
const crypto = require('crypto');
const fs = require('fs');
const BOOT = require('./_bootstrap.json');

const SESSION_TTL = 60 * 60 * 24 * 30, FAIL_TTL = 15 * 60, FAIL_MAX = 10;

/* ---------- storage ---------- */
const KV_URL = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const KV_TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
const FILE = process.env.FFCS_FILE_STORE;

async function redis(cmd) {
  const r = await fetch(KV_URL, { method: 'POST', headers: { Authorization: 'Bearer ' + KV_TOKEN, 'Content-Type': 'application/json' }, body: JSON.stringify(cmd) });
  const j = await r.json();
  if (j.error) throw new Error('storage: ' + j.error);
  return j.result;
}
function fileDb() { try { return JSON.parse(fs.readFileSync(FILE, 'utf8')); } catch (e) { return {}; } }
const store = {
  ok: () => !!(KV_URL && KV_TOKEN) || !!FILE,
  async get(k) {
    if (FILE) { const db = fileDb(), e = db[k]; return e && (!e.exp || e.exp > Date.now()) ? e.v : null; }
    const v = await redis(['GET', k]); return v == null ? null : JSON.parse(v);
  },
  async set(k, v, ttl) {
    if (FILE) { const db = fileDb(); db[k] = { v, exp: ttl ? Date.now() + ttl * 1000 : 0 }; fs.writeFileSync(FILE, JSON.stringify(db)); return; }
    await redis(ttl ? ['SET', k, JSON.stringify(v), 'EX', ttl] : ['SET', k, JSON.stringify(v)]);
  },
  async del(k) {
    if (FILE) { const db = fileDb(); delete db[k]; fs.writeFileSync(FILE, JSON.stringify(db)); return; }
    await redis(['DEL', k]);
  }
};

/* ---------- helpers ---------- */
const hashPw = (pw, salt) => crypto.scryptSync(String(pw), salt, 32).toString('hex');
function checkPw(pw, u) {
  const a = Buffer.from(hashPw(pw, u.salt), 'hex'), b = Buffer.from(u.hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
const pub = u => ({ id: u.id, name: u.name, login: u.login, admin: !!u.admin, active: u.active !== false });
const normLogin = s => String(s || '').trim().toLowerCase();
class HttpError extends Error { constructor(code, msg) { super(msg); this.code = code; } }
const fail = (code, msg) => { throw new HttpError(code, msg); };

async function getUsers() {
  let users = await store.get('ffcs:users');
  if (!users) {   // first run: create the initial accounts (passwords were generated offline, only hashes are stored)
    users = {};
    BOOT.users.forEach(u => { users[u.id] = Object.assign({ active: true, created: Date.now() }, u); });
    await store.set('ffcs:users', users);
    for (const id of Object.keys(BOOT.data || {})) await store.set('ffcs:data:' + id, { data: BOOT.data[id], rev: 1, at: Date.now() });
  }
  return users;
}
async function getDist() { return (await store.get('ffcs:dist')) || {}; }
async function getData(id) { return (await store.get('ffcs:data:' + id)) || { data: null, rev: 0, at: 0 }; }
async function putData(id, data) {
  const cur = await getData(id), rec = { data, rev: cur.rev + 1, at: Date.now() };
  await store.set('ffcs:data:' + id, rec); return rec;
}
/* copy the source's faculty tags into the target's own data (source is only read) */
async function copyTags(sourceId, targetId) {
  const src = await getData(sourceId), tgt = await getData(targetId);
  const data = Object.assign({}, tgt.data || {}, { tags: JSON.parse(JSON.stringify((src.data && src.data.tags) || {})), tagsFrom: sourceId, tagsAt: Date.now() });
  return putData(targetId, data);
}
async function resetTags(targetId) {
  const tgt = await getData(targetId);
  const data = Object.assign({}, tgt.data || {}, { tags: {}, tagsFrom: null, tagsAt: Date.now() });
  return putData(targetId, data);
}

async function auth(req) {
  const h = req.headers.authorization || '', tok = h.startsWith('Bearer ') ? h.slice(7) : '';
  if (!/^[0-9a-f]{64}$/.test(tok)) fail(401, 'Please sign in');
  const id = await store.get('ffcs:sess:' + tok); if (!id) fail(401, 'Session expired — please sign in again');
  const users = await getUsers(), u = users[id];
  if (!u || u.active === false) { await store.del('ffcs:sess:' + tok); fail(401, 'Your access has been removed'); }
  return { u, users, tok };
}
const needAdmin = me => { if (!me.u.admin) fail(403, 'Admins only'); };

/* ---------- actions ---------- */
const ACTIONS = {
  async login(b) {
    const login = normLogin(b.login), fk = 'ffcs:fail:' + login;
    const fails = (await store.get(fk)) || 0;
    if (fails >= FAIL_MAX) fail(429, 'Too many wrong passwords — try again in 15 minutes');
    const users = await getUsers(), u = Object.values(users).find(x => x.login === login);
    if (!u || !checkPw(b.password || '', u)) { await store.set(fk, fails + 1, FAIL_TTL); fail(401, 'Wrong login ID or password'); }
    if (u.active === false) fail(403, 'Your access to this platform is disabled — ask an admin');
    await store.del(fk);
    const token = crypto.randomBytes(32).toString('hex');
    await store.set('ffcs:sess:' + token, u.id, SESSION_TTL);
    return { token, user: pub(u) };
  },
  async logout(b, req) { const me = await auth(req); await store.del('ffcs:sess:' + me.tok); return { ok: true }; },
  async me(b, req) { const me = await auth(req); return { user: pub(me.u) }; },
  async getData(b, req) { const me = await auth(req); const d = await getData(me.u.id); return Object.assign({ user: pub(me.u) }, d); },
  async putData(b, req) {
    const me = await auth(req);
    if (!b.data || typeof b.data !== 'object') fail(400, 'Missing data');
    if (JSON.stringify(b.data).length > 900000) fail(413, 'Data too large');
    const cur = await getData(me.u.id);
    if (!b.force && b.baseRev !== cur.rev) return { conflict: true, rev: cur.rev, data: cur.data };
    return putData(me.u.id, b.data);
  },
  async changePassword(b, req) {
    const me = await auth(req);
    if (!checkPw(b.old || '', me.u)) fail(400, 'Current password is wrong');
    if (String(b.password || '').length < 6) fail(400, 'New password must be at least 6 characters');
    me.u.salt = crypto.randomBytes(16).toString('hex'); me.u.hash = hashPw(b.password, me.u.salt);
    await store.set('ffcs:users', me.users); return { ok: true };
  },

  /* ----- admin ----- */
  async adminList(b, req) {
    const me = await auth(req); needAdmin(me);
    const dist = await getDist(), out = [];
    for (const u of Object.values(me.users)) {
      const d = await getData(u.id), data = d.data || {};
      out.push(Object.assign(pub(u), {
        tags: Object.keys(data.tags || {}).length, added: (data.added || []).length, rev: d.rev, at: d.at,
        tagsFrom: data.tagsFrom || null, receivesFrom: dist[u.id] || null
      }));
    }
    out.sort((a, b) => (b.admin - a.admin) || a.name.localeCompare(b.name));
    return { users: out, dist };
  },
  async addUser(b, req) {
    const me = await auth(req); needAdmin(me);
    const login = normLogin(b.login), name = String(b.name || '').trim();
    if (!name) fail(400, 'Name is required');
    if (!/^[a-z0-9._-]{3,32}$/.test(login)) fail(400, 'Login ID: 3–32 characters, letters, digits, . _ -');
    if (Object.values(me.users).some(u => u.login === login)) fail(400, 'Login ID already taken');
    if (String(b.password || '').length < 6) fail(400, 'Password must be at least 6 characters');
    const id = 'u' + crypto.randomBytes(6).toString('hex'), salt = crypto.randomBytes(16).toString('hex');
    me.users[id] = { id, name, login, admin: !!b.admin, active: true, salt, hash: hashPw(b.password, salt), created: Date.now() };
    await store.set('ffcs:users', me.users);
    return { user: pub(me.users[id]) };
  },
  async updateUser(b, req) {
    const me = await auth(req); needAdmin(me);
    const u = me.users[b.id]; if (!u) fail(404, 'No such user');
    if (b.name !== undefined) { const n = String(b.name).trim(); if (!n) fail(400, 'Name is required'); u.name = n; }
    if (b.login !== undefined) {
      const l = normLogin(b.login);
      if (!/^[a-z0-9._-]{3,32}$/.test(l)) fail(400, 'Login ID: 3–32 characters, letters, digits, . _ -');
      if (Object.values(me.users).some(x => x.login === l && x.id !== u.id)) fail(400, 'Login ID already taken');
      u.login = l;
    }
    if (b.password) { if (String(b.password).length < 6) fail(400, 'Password must be at least 6 characters'); u.salt = crypto.randomBytes(16).toString('hex'); u.hash = hashPw(b.password, u.salt); }
    const admins = () => Object.values(me.users).filter(x => x.admin && x.active !== false).length;
    if (b.admin !== undefined) { u.admin = !!b.admin; if (!admins()) fail(400, 'At least one active admin is required'); }
    if (b.active !== undefined) {
      if (u.id === me.u.id && !b.active) fail(400, 'You cannot remove your own access');
      u.active = !!b.active; if (!admins()) fail(400, 'At least one active admin is required');
    }
    await store.set('ffcs:users', me.users);
    return { user: pub(u) };
  },
  async deleteUser(b, req) {
    const me = await auth(req); needAdmin(me);
    const u = me.users[b.id]; if (!u) fail(404, 'No such user');
    if (u.id === me.u.id) fail(400, 'You cannot delete your own account');
    delete me.users[u.id];
    if (!Object.values(me.users).some(x => x.admin && x.active !== false)) fail(400, 'At least one active admin is required');
    const dist = await getDist();
    delete dist[u.id];
    for (const t of Object.keys(dist)) if (dist[t] === u.id) delete dist[t];   // receivers keep their copy; they just stop being linked
    await store.set('ffcs:users', me.users); await store.set('ffcs:dist', dist); await store.del('ffcs:data:' + u.id);
    return { ok: true };
  },
  /* source's tags -> each target in `targets` (fresh copy for newly added targets, or for all with resend);
     targets that were receiving from this source and are no longer selected are reset to Neutral. */
  async setDist(b, req) {
    const me = await auth(req); needAdmin(me);
    const src = b.source; if (!me.users[src]) fail(404, 'No such source user');
    const targets = [...new Set((b.targets || []).filter(t => me.users[t] && t !== src))];
    const dist = await getDist(), sent = [], reset = [];
    for (const t of Object.keys(dist)) if (dist[t] === src && !targets.includes(t)) { delete dist[t]; await resetTags(t); reset.push(t); }
    for (const t of targets) {
      if (dist[t] !== src || b.resend) { dist[t] = src; await copyTags(src, t); sent.push(t); }
    }
    await store.set('ffcs:dist', dist);
    return { dist, sent, reset };
  }
};

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  try {
    if (!store.ok()) fail(503, 'Account database is not connected (add Upstash Redis to the Vercel project)');
    if (req.method !== 'POST') fail(405, 'POST only');
    let b = req.body;
    if (typeof b === 'string') { try { b = JSON.parse(b); } catch (e) { b = {}; } }
    b = b || {};
    const fn = ACTIONS[b.action]; if (!fn) fail(400, 'Unknown action');
    res.status(200).json(await fn(b, req));
  } catch (e) {
    res.status(e.code || 500).json({ error: e.code ? e.message : 'Server error: ' + e.message });
  }
};
