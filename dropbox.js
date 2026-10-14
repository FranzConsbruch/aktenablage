/* Dropbox: Anmeldung (OAuth mit PKCE), Ordnerliste, Dateien lesen und schreiben.
   Kein App secret - alles läuft im Browser.
   Zwei Anmeldewege: login() im Outlook-Taskpane (Office-Dialog),
   loginWeb()/completeWebLogin() auf der Aktenseite (Weiterleitung im selben Fenster). */

const DBX = (() => {
  const TOKEN_KEY = "dbx_tokens_v1";
  const API = "https://api.dropboxapi.com/2";
  const CONTENT = "https://content.dropboxapi.com/2";
  const VERIFIER_KEY = "dbx_pkce_verifier_web";

  let tokens = null; // { access_token, refresh_token, expires_at }

  function load() {
    if (tokens) return tokens;
    try { tokens = JSON.parse(localStorage.getItem(TOKEN_KEY) || "null"); } catch { tokens = null; }
    return tokens;
  }
  function save(t) {
    tokens = t;
    try { localStorage.setItem(TOKEN_KEY, JSON.stringify(t)); } catch {}
  }
  function forget() {
    tokens = null;
    try { localStorage.removeItem(TOKEN_KEY); } catch {}
  }

  /* Dropbox verlangt die Argumente als HTTP-Header. Header dürfen nur ASCII enthalten,
     Umlaute in Pfaden wie "03_Süschendorf" müssen also escaped werden. */
  function asciiJSON(obj) {
    return JSON.stringify(obj).replace(/[\u0080-￿]/g, c =>
      "\\u" + c.charCodeAt(0).toString(16).padStart(4, "0"));
  }

  async function ensureToken() {
    const t = load();
    if (!t) throw new Error("Nicht bei Dropbox angemeldet.");
    if (t.expires_at && Date.now() < t.expires_at - 60000) return t.access_token;
    if (!t.refresh_token) throw new Error("Anmeldung abgelaufen. Bitte neu anmelden.");

    const body = new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: t.refresh_token,
      client_id: CONFIG.DROPBOX_APP_KEY,
    });
    const res = await fetch("https://api.dropbox.com/oauth2/token", { method: "POST", body });
    if (!res.ok) { forget(); throw new Error("Anmeldung abgelaufen. Bitte neu anmelden."); }
    const j = await res.json();
    save({ ...t, access_token: j.access_token, expires_at: Date.now() + (j.expires_in || 14400) * 1000 });
    return j.access_token;
  }

  async function rpc(path, arg) {
    const token = await ensureToken();
    const res = await fetch(API + path, {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify(arg),
    });
    if (!res.ok) throw new Error("Dropbox " + path + ": " + res.status + " " + (await res.text()).slice(0, 300));
    return res.json();
  }

  /* ---------- PKCE-Helfer (gemeinsam für beide Anmeldewege) ---------- */

  function b64url(buf) {
    return btoa(String.fromCharCode(...new Uint8Array(buf)))
      .replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  async function challenge(verifier) {
    return b64url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
  }
  function newVerifier() {
    const a = new Uint8Array(64); crypto.getRandomValues(a);
    return b64url(a.buffer);
  }
  async function authorizeUrl(verifier, redirect) {
    const url = new URL("https://www.dropbox.com/oauth2/authorize");
    url.searchParams.set("client_id", CONFIG.DROPBOX_APP_KEY);
    url.searchParams.set("response_type", "code");
    url.searchParams.set("code_challenge", await challenge(verifier));
    url.searchParams.set("code_challenge_method", "S256");
    url.searchParams.set("redirect_uri", redirect);
    url.searchParams.set("token_access_type", "offline");
    return url.href;
  }
  async function exchangeCode(code, verifier, redirect) {
    const body = new URLSearchParams({
      code, grant_type: "authorization_code",
      client_id: CONFIG.DROPBOX_APP_KEY,
      code_verifier: verifier, redirect_uri: redirect,
    });
    const res = await fetch("https://api.dropbox.com/oauth2/token", { method: "POST", body });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error("Dropbox hat die Anmeldung abgelehnt: " + (j.error_description || res.status));
    return j;
  }

  /* ---------- Anmeldung im Outlook-Taskpane (Office-Dialog) ---------- */

  function login() {
    return new Promise((resolve, reject) => {
      const url = new URL("auth.html", window.location.href).href;
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 40, promptBeforeOpen: false }, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(result.error.message));
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
          dialog.close();
          let msg;
          try { msg = JSON.parse(arg.message); } catch { return reject(new Error("Unerwartete Antwort der Anmeldung.")); }
          if (msg.error) return reject(new Error(msg.error));
          save({
            access_token: msg.access_token,
            refresh_token: msg.refresh_token,
            expires_at: Date.now() + (msg.expires_in || 14400) * 1000,
          });
          resolve();
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => reject(new Error("Anmeldung abgebrochen.")));
      });
    });
  }

  /* ---------- Anmeldung im Browser (Aktenseite): Weiterleitung im selben Fenster ---------- */

  async function loginWeb() {
    const verifier = newVerifier();
    sessionStorage.setItem(VERIFIER_KEY, verifier);
    const redirect = window.location.origin + window.location.pathname;
    window.location.assign(await authorizeUrl(verifier, redirect));
  }

  /* Beim Laden der Seite aufrufen. Gibt true zurück, wenn gerade eine Anmeldung
     abgeschlossen wurde (die Seite ist dann angemeldet und die URL bereinigt). */
  async function completeWebLogin() {
    const q = new URLSearchParams(window.location.search);
    const code = q.get("code");
    if (!code) return false;
    const verifier = sessionStorage.getItem(VERIFIER_KEY);
    sessionStorage.removeItem(VERIFIER_KEY);
    const redirect = window.location.origin + window.location.pathname;
    if (!verifier) throw new Error("Anmeldung abgelaufen, bitte erneut versuchen.");
    const j = await exchangeCode(code, verifier, redirect);
    save({
      access_token: j.access_token,
      refresh_token: j.refresh_token,
      expires_at: Date.now() + (j.expires_in || 14400) * 1000,
    });
    history.replaceState(null, "", redirect);
    return true;
  }

  function isLoggedIn() { return !!load(); }

  /* ---------- Lesen ---------- */

  async function listChildren(path) {
    const out = [];
    let r = await rpc("/files/list_folder", { path, recursive: false, limit: 500 });
    out.push(...r.entries);
    while (r.has_more) { r = await rpc("/files/list_folder/continue", { cursor: r.cursor }); out.push(...r.entries); }
    return out;
  }

  /* Baut die Aktenliste: Bereiche -> Themenordner -> deren Unterordner. */
  async function loadAkten() {
    const akten = [];
    for (const root of CONFIG.ROOTS) {
      let level1;
      try { level1 = await listChildren(root.path); }
      catch (e) { console.warn("Bereich nicht lesbar: " + root.path, e); continue; }

      const folders1 = level1.filter(e => e[".tag"] === "folder" && !CONFIG.EXCLUDE.includes(e.name));
      const jobs = folders1.map(async f => {
        akten.push({ label: root.label + " › " + f.name, path: f.path_display || (root.path + "/" + f.name), root: root.path });
        if (CONFIG.DEPTH < 2) return;
        try {
          const level2 = await listChildren(f.path_display || (root.path + "/" + f.name));
          for (const g of level2) {
            if (g[".tag"] !== "folder" || CONFIG.EXCLUDE.includes(g.name)) continue;
            akten.push({ label: root.label + " › " + f.name + "/" + g.name, path: g.path_display, root: root.path });
          }
        } catch (e) { console.warn("Unterordner nicht lesbar: " + f.name, e); }
      });
      await Promise.all(jobs);
    }
    akten.sort((a, b) => a.label.localeCompare(b.label, "de"));
    return akten;
  }

  /* Lädt eine Datei herunter. Gibt einen ArrayBuffer zurück. */
  async function download(path) {
    const token = await ensureToken();
    const res = await fetch(CONTENT + "/files/download", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Dropbox-API-Arg": asciiJSON({ path }) },
      body: "",
    });
    if (!res.ok) throw new Error("Download fehlgeschlagen (" + res.status + "): " + (await res.text()).slice(0, 300));
    return res.arrayBuffer();
  }

  /* Wie download(), aber null statt Fehler, wenn die Datei nicht existiert. */
  async function downloadIfExists(path) {
    try { return await download(path); }
    catch (e) { if (String(e.message).includes("not_found") || String(e.message).includes("409")) return null; throw e; }
  }

  /* Kurzlebiger Link (4 Stunden) zum Öffnen einer Datei im Browser. */
  async function tempLink(path) {
    const r = await rpc("/files/get_temporary_link", { path });
    return r.link;
  }

  /* ---------- Schreiben ---------- */

  async function uploadRaw(path, bytes, mode) {
    const token = await ensureToken();
    const res = await fetch(CONTENT + "/files/upload", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/octet-stream",
        "Dropbox-API-Arg": asciiJSON({ path, mode, autorename: mode === "add", mute: true, strict_conflict: false }),
      },
      body: bytes,
    });
    if (!res.ok) throw new Error("Upload fehlgeschlagen (" + res.status + "): " + (await res.text()).slice(0, 300));
    return res.json();
  }

  /* Neue Datei: nie überschreiben, bei Namensgleichheit nummerieren. */
  function upload(path, bytes) { return uploadRaw(path, bytes, "add"); }

  /* Bekannte Datei ersetzen (z. B. beteiligte.json). */
  function uploadOverwrite(path, bytes) { return uploadRaw(path, bytes, "overwrite"); }

  async function deleteFile(path) { return rpc("/files/delete_v2", { path }); }

  return {
    login, loginWeb, completeWebLogin, isLoggedIn, forget,
    loadAkten, listChildren, download, downloadIfExists, tempLink,
    upload, uploadOverwrite, deleteFile,
  };
})();
