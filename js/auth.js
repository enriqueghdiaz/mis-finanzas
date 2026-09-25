/* Inicio de sesión con Microsoft (OAuth 2.0 + PKCE, sin librerías externas).
   Los tokens se guardan solo en este dispositivo. */
(function () {
  const LS = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} }
  };
  const SCOPES = "openid profile offline_access User.Read Files.ReadWrite";

  function cfg() {
    const s = LS.get("fin_settings") || {};
    return {
      clientId: (s.clientId || window.APP_CONFIG.CLIENT_ID || "").trim(),
      authority: (s.authority || window.APP_CONFIG.AUTHORITY || "consumers").trim()
    };
  }
  function redirectUri() { return location.origin + location.pathname; }
  function b64url(bytes) {
    let s = ""; bytes.forEach(b => s += String.fromCharCode(b));
    return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  function randomString(n) { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a); }
  async function sha256(str) {
    const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
    return new Uint8Array(d);
  }
  function endpoint(kind) {
    return `https://login.microsoftonline.com/${cfg().authority}/oauth2/v2.0/${kind}`;
  }

  async function login() {
    const { clientId } = cfg();
    if (!clientId) throw new Error("Falta el Client ID. Añádelo en Ajustes.");
    const verifier = randomString(48);
    const state = randomString(16);
    LS.set("fin_pkce", { verifier, state, t: Date.now() });
    const challenge = b64url(await sha256(verifier));
    const p = new URLSearchParams({
      client_id: clientId, response_type: "code", redirect_uri: redirectUri(),
      scope: SCOPES, code_challenge: challenge, code_challenge_method: "S256",
      state, response_mode: "query", prompt: "select_account"
    });
    location.assign(endpoint("authorize") + "?" + p.toString());
  }

  async function tokenRequest(params) {
    const { clientId } = cfg();
    const body = new URLSearchParams(Object.assign({ client_id: clientId, scope: SCOPES }, params));
    const r = await fetch(endpoint("token"), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error_description || j.error || ("Error " + r.status));
    const tok = {
      access: j.access_token,
      refresh: j.refresh_token || (LS.get("fin_token") || {}).refresh,
      exp: Date.now() + (j.expires_in || 3600) * 1000
    };
    LS.set("fin_token", tok);
    return tok;
  }

  // Se llama al arrancar: si volvemos de Microsoft con ?code=..., canjea el código.
  async function handleRedirect() {
    const q = new URLSearchParams(location.search);
    if (q.get("error")) {
      history.replaceState(null, "", redirectUri());
      throw new Error(q.get("error_description") || q.get("error"));
    }
    const code = q.get("code");
    if (!code) return false;
    const saved = LS.get("fin_pkce");
    history.replaceState(null, "", redirectUri());
    if (!saved || saved.state !== q.get("state")) throw new Error("La respuesta de inicio de sesión no coincide. Inténtalo de nuevo.");
    LS.del("fin_pkce");
    await tokenRequest({
      grant_type: "authorization_code", code, redirect_uri: redirectUri(), code_verifier: saved.verifier
    });
    return true;
  }

  let refreshing = null;
  async function getToken() {
    const t = LS.get("fin_token");
    if (!t) throw new AuthError("Sin sesión");
    if (t.access && t.exp - Date.now() > 90 * 1000) return t.access;
    if (!t.refresh) throw new AuthError("Sesión caducada");
    if (!refreshing) {
      refreshing = tokenRequest({ grant_type: "refresh_token", refresh_token: t.refresh })
        .catch(e => { throw new AuthError("Sesión caducada: " + e.message); })
        .finally(() => { refreshing = null; });
    }
    return (await refreshing).access;
  }
  class AuthError extends Error {}

  function isSignedIn() { const t = LS.get("fin_token"); return !!(t && (t.refresh || t.exp > Date.now())); }
  function logout() { LS.del("fin_token"); LS.del("fin_pkce"); }

  window.Auth = { login, handleRedirect, getToken, isSignedIn, logout, redirectUri, cfg, AuthError, LS };
})();
