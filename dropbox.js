/* Dropbox: Anmeldung (OAuth mit PKCE), Ordnerliste, Upload.
   Kein App secret - das Add-in läuft im Browser. */

const DBX = (() => {
  const TOKEN_KEY = "dbx_tokens_v1";
  const API = "https://api.dropboxapi.com/2";
  const CONTENT = "https://content.dropboxapi.com/2";

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
     Umlaute in Pfaden wie "03_Süschendorf" müssen also escaped werden. Ohne das
     scheitert jeder Upload in einen Ordner mit Umlaut - der klassische Stolperstein. */
  function asciiJSON(obj) {
    return JSON.stringify(obj).replace(/[\u0080-\uffff]/g, c =>
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

  /* Anmeldung über die Office-Dialog-API. Ein Add-in läuft in einem iframe,
     Popups sind blockiert - deshalb der Umweg über auth.html auf derselben Domain. */
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

  function isLoggedIn() { return !!load(); }

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
        akten.push({ label: root.label + " › " + f.name, path: f.path_display || (root.path + "/" + f.name) });
        if (CONFIG.DEPTH < 2) return;
        try {
          const level2 = await listChildren(f.path_display || (root.path + "/" + f.name));
          for (const g of level2) {
            if (g[".tag"] !== "folder" || CONFIG.EXCLUDE.includes(g.name)) continue;
            akten.push({ label: root.label + " › " + f.name + "/" + g.name, path: g.path_display });
          }
        } catch (e) { console.warn("Unterordner nicht lesbar: " + f.name, e); }
      });
      await Promise.all(jobs);
    }
    akten.sort((a, b) => a.label.localeCompare(b.label, "de"));
    return akten;
  }

  /* Lädt eine Datei hoch. bytes = ArrayBuffer oder Uint8Array.
     autorename verhindert, dass eine bestehende Datei überschrieben wird. */
  async function upload(path, bytes) {
    const token = await ensureToken();
    const res = await fetch(CONTENT + "/files/upload", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/octet-stream",
        "Dropbox-API-Arg": asciiJSON({ path, mode: "add", autorename: true, mute: true, strict_conflict: false }),
      },
      body: bytes,
    });
    if (!res.ok) throw new Error("Upload fehlgeschlagen (" + res.status + "): " + (await res.text()).slice(0, 300));
    return res.json();
  }

  return { login, isLoggedIn, forget, loadAkten, upload };
})();
