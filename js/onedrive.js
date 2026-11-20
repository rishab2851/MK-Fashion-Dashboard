// OneDrive connection through Microsoft Graph (OAuth 2.0 authorisation code flow with PKCE, no library needed).
// The Microsoft sign in is kept per device in this browser. Microsoft limits it to 24 hours for web apps,
// after which one tap on "Reconnect OneDrive" renews it (usually without typing a password).
import { msConfig } from './config.js';

const TOK = 'mkf_ms_tokens';
const AUTH_MSG = 'mkf_ms_auth';
const PKCE = 'mkf_ms_pkce';
const SCOPES = 'openid profile offline_access User.Read Files.Read';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const base = () => `https://login.microsoftonline.com/${msConfig.authority || 'consumers'}/oauth2/v2.0`;

export const redirectUri = () => location.origin + location.pathname.replace(/[^/]*$/, '') + 'ms-auth.html';
export const configured = () => !!msConfig.clientId && !String(msConfig.clientId).startsWith('PASTE');

const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const rand = (n = 32) => b64url(crypto.getRandomValues(new Uint8Array(n)));
const load = () => { try { return JSON.parse(localStorage.getItem(TOK)) || null; } catch { return null; } };
const save = (t) => localStorage.setItem(TOK, JSON.stringify(t));

export function status() {
  const t = load();
  if (!t) return { connected: false };
  const live = t.rtExp > Date.now() + 60000;
  return { connected: live, expired: !live, account: t.account, rtExp: t.rtExp };
}
export function disconnect() { localStorage.removeItem(TOK); }

// Must be called directly from a tap or click, so the browser allows the sign in window.
export async function connect({ chooseAccount = false } = {}) {
  if (!configured()) throw new Error('Microsoft app id is not set in js/config.js yet.');
  const w = window.open('', 'mkf_ms_login', 'width=520,height=680');
  const verifier = rand(48);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = rand(16);
  const prev = load();
  const p = new URLSearchParams({
    client_id: msConfig.clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query',
    scope: SCOPES, state, code_challenge: challenge, code_challenge_method: 'S256',
  });
  if (chooseAccount || !prev?.account) p.set('prompt', 'select_account');
  else if (prev.account?.email) p.set('login_hint', prev.account.email);
  localStorage.removeItem(AUTH_MSG);
  localStorage.setItem(PKCE, JSON.stringify({ verifier, state, t: Date.now() }));
  const url = `${base()}/authorize?${p}`;
  if (w) w.location.href = url; else { location.href = url; return null; }

  const params = await new Promise((resolve, reject) => {
    const t0 = Date.now();
    const iv = setInterval(() => {
      const raw = localStorage.getItem(AUTH_MSG);
      if (raw) { clearInterval(iv); localStorage.removeItem(AUTH_MSG); resolve(new URLSearchParams(JSON.parse(raw).search)); }
      else if (Date.now() - t0 > 5 * 60000) { clearInterval(iv); reject(new Error('Microsoft sign in timed out.')); }
    }, 400);
  });
  return finish(params);
}

// If the sign in went through a full page redirect (popup blocked), finish it when the app loads.
export async function completePending() {
  const raw = localStorage.getItem(AUTH_MSG);
  if (!raw || !localStorage.getItem(PKCE)) return null;
  localStorage.removeItem(AUTH_MSG);
  return finish(new URLSearchParams(JSON.parse(raw).search));
}

async function finish(params) {
  const pk = JSON.parse(localStorage.getItem(PKCE) || 'null');
  localStorage.removeItem(PKCE);
  if (params.get('error')) throw new Error(params.get('error_description') || params.get('error'));
  if (!pk || params.get('state') !== pk.state) throw new Error('Sign in response did not match. Please try again.');
  const prev = load();
  const tok = await tokenCall({ grant_type: 'authorization_code', code: params.get('code'), redirect_uri: redirectUri(), code_verifier: pk.verifier });
  const t = fromResponse(tok, null);
  t.account = prev?.account || null;
  save(t);
  try {
    const me = await graph('/me?$select=displayName,mail,userPrincipalName');
    t.account = { name: me.displayName, email: me.mail || me.userPrincipalName };
    save(t);
  } catch { /* account name is cosmetic */ }
  return t.account;
}

async function tokenCall(body) {
  const r = await fetch(`${base()}/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: msConfig.clientId, scope: SCOPES, ...body }),
  });
  const j = await r.json();
  if (!r.ok) { const e = new Error(j.error_description || j.error || 'Token error'); e.code = j.error; throw e; }
  return j;
}
function fromResponse(j, prev) {
  return {
    access: j.access_token, exp: Date.now() + (j.expires_in - 120) * 1000,
    refresh: j.refresh_token || prev?.refresh,
    // Single page app refresh tokens live 24 hours and do not slide, so keep the first expiry.
    rtExp: prev?.rtExp || Date.now() + (j.refresh_token_expires_in || 86400) * 1000,
    account: prev?.account || null,
  };
}

export class ReconnectNeeded extends Error { constructor() { super('OneDrive sign in has expired. Tap Reconnect OneDrive.'); this.reconnect = true; } }

async function accessToken() {
  const t = load();
  if (!t) throw new ReconnectNeeded();
  if (t.exp > Date.now()) return t.access;
  if (!t.refresh || t.rtExp < Date.now()) throw new ReconnectNeeded();
  try {
    const j = await tokenCall({ grant_type: 'refresh_token', refresh_token: t.refresh });
    const n = fromResponse(j, t); save(n); return n.access;
  } catch (e) {
    if (e.code === 'invalid_grant' || e.code === 'interaction_required') { t.rtExp = 0; save(t); throw new ReconnectNeeded(); }
    throw e;
  }
}

export async function graph(path, { raw = false } = {}) {
  let tok = await accessToken();
  let r = await fetch(GRAPH + path, { headers: { Authorization: 'Bearer ' + tok } });
  if (r.status === 401) {
    const t = load(); if (t) { t.exp = 0; save(t); }
    tok = await accessToken();
    r = await fetch(GRAPH + path, { headers: { Authorization: 'Bearer ' + tok } });
  }
  if (!r.ok) {
    let msg = `OneDrive error ${r.status}`;
    try { const j = await r.json(); msg = j.error?.message || msg; } catch { /* ignore */ }
    const e = new Error(msg); e.status = r.status; throw e;
  }
  return raw ? r : r.json();
}

const SEL = 'id,name,cTag,eTag,size,lastModifiedDateTime,lastModifiedBy,parentReference,webUrl';
const pathOf = (it) => ((it.parentReference?.path || '').replace(/^\/drive\/root:?/, '') || '/') ;

export async function findWorkbooks(q) {
  const term = (q || 'xlsx').replace(/'/g, "''");
  const j = await graph(`/me/drive/root/search(q='${encodeURIComponent(term)}')?$top=100&$select=${SEL},file`);
  return (j.value || []).filter((it) => it.file && /\.xls[xm]$/i.test(it.name))
    .map((it) => ({ itemId: it.id, driveId: it.parentReference?.driveId, name: it.name, path: pathOf(it), modified: it.lastModifiedDateTime, size: it.size }))
    .sort((a, b) => (b.modified || '').localeCompare(a.modified || ''));
}
export async function fileInfo(ref) {
  const it = await graph(`/drives/${ref.driveId}/items/${ref.itemId}?$select=${SEL}`);
  return { cTag: it.cTag, eTag: it.eTag, name: it.name, path: pathOf(it), modified: it.lastModifiedDateTime,
    modifiedBy: it.lastModifiedBy?.user?.displayName || '', size: it.size, webUrl: it.webUrl };
}
export async function download(ref) {
  const it = await graph(`/drives/${ref.driveId}/items/${ref.itemId}`);
  const url = it['@microsoft.graph.downloadUrl'];
  if (url) {
    try { const r = await fetch(url); if (r.ok) return { buf: await r.arrayBuffer(), cTag: it.cTag, modified: it.lastModifiedDateTime }; }
    catch { /* fall through to the authenticated route */ }
  }
  const r = await graph(`/drives/${ref.driveId}/items/${ref.itemId}/content`, { raw: true });
  return { buf: await r.arrayBuffer(), cTag: it.cTag, modified: it.lastModifiedDateTime };
}
