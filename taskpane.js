/* Taskpane des Outlook-Add-ins.
   Lesemodus:    Mail als .eml plus Anhänge in die gewählte Dropbox-Akte ablegen,
                 Kategorie "Abgelegt" setzen, Zielordner an der Mail vermerken.
   Verfassen:    Datei aus einer Dropbox-Akte als Anhang in die Mail legen. */

const AKTEN_CACHE_KEY = "akten_cache_v2";
const PROP_KEY = "ablage"; // benutzerdefinierte Eigenschaft an der Mail

let akten = [];
let selected = null;
let customProps = null;
let composeMode = false;
let akteFiles = [];

const $ = id => document.getElementById(id);

/* ---------- Anzeige ---------- */

function setStatus(text) { $("status").textContent = text || ""; }

function banner(text, kind, el) {
  const b = $(el || "banner");
  if (!text) { b.className = "hidden"; b.textContent = ""; return; }
  b.className = "note " + (kind || "info");
  b.textContent = text;
}

function fmtSize(n) {
  if (n == null) return "";
  if (n < 1024) return n + " B";
  if (n < 1048576) return (n / 1024).toFixed(0) + " KB";
  return (n / 1048576).toFixed(1) + " MB";
}

function renderList() {
  const q = $("filter").value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const hits = akten.filter(a => {
    const hay = a.label.toLowerCase();
    return words.every(w => hay.includes(w));
  }).slice(0, 200);

  const ul = $("list");
  ul.innerHTML = "";
  for (const a of hits) {
    const li = document.createElement("li");
    li.textContent = a.label;
    li.setAttribute("role", "option");
    li.setAttribute("aria-selected", selected && selected.path === a.path ? "true" : "false");
    li.onclick = () => {
      selected = a;
      if (composeMode) { openAkteFiles(a); return; }
      renderList();
      $("btn-file").disabled = false;
    };
    ul.appendChild(li);
  }
  if (!hits.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = akten.length ? "Keine Akte passt zu diesem Suchbegriff." : "Noch keine Akten geladen.";
    ul.appendChild(li);
  }
}

/* ---------- Dateinamen ---------- */

function sanitize(s) {
  return (s || "")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+|\.+$/g, "");
}

function shorten(s, max) {
  return s.length <= max ? s : s.slice(0, max).trim().replace(/[\s\-_]+$/, "");
}

function localDate(d) {
  const p = n => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
}

function baseName(item) {
  const date = localDate(item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date());
  const from = item.from ? (item.from.displayName || item.from.emailAddress) : "unbekannt";
  const subject = item.subject || "ohne Betreff";
  return date + "_" + sanitize(from) + "_" + shorten(sanitize(subject), CONFIG.SUBJECT_MAX);
}

/* ---------- Outlook: Mail lesen ---------- */

function restToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, r => {
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message));
    });
  });
}

function supportsSet(v) {
  try { return Office.context.requirements.isSetSupported("Mailbox", v); } catch { return false; }
}

/* Bevorzugter Weg (Requirement Set 1.14): Outlook liefert die Nachricht direkt
   als EML, base64-kodiert. Braucht kein Token und keinen REST-Aufruf. */
function emlViaOfficeApi() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAsFileAsync(r => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(b64ToBytes(r.value));
      else reject(new Error("Outlook konnte die Nachricht nicht liefern: " + r.error.message));
    });
  });
}

/* Alter Weg über die Outlook-REST-Schnittstelle - nur noch Rückfallebene.
   In neuen Tenants sind die Legacy-Tokens aus (Fehler 9018). */
async function emlViaRest() {
  const token = await restToken();
  const restId = Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  const url = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + restId + "/$value";
  const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) throw new Error("Nachricht konnte nicht gelesen werden (" + res.status + ").");
  return res.arrayBuffer();
}

async function fetchMime() {
  const item = Office.context.mailbox.item;
  if (item && typeof item.getAsFileAsync === "function" && supportsSet("1.14")) {
    return emlViaOfficeApi();
  }
  return emlViaRest();
}

function attachmentContent(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, r => {
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message));
    });
  });
}

/* ---------- base64 <-> Bytes ---------- */

function b64ToBytes(b64) {
  const bin = atob(b64);
  const a = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
  return a;
}

/* Stückweise, sonst reißt bei großen Dateien der Aufruf-Stack. */
function bytesToB64(buf) {
  const u = new Uint8Array(buf);
  let s = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < u.length; i += CHUNK) {
    s += String.fromCharCode.apply(null, u.subarray(i, i + CHUNK));
  }
  return btoa(s);
}

/* ---------- Outlook: Kategorie und Vermerk ---------- */

async function ensureCategory() {
  const name = CONFIG.CATEGORY;
  await new Promise(resolve => {
    Office.context.mailbox.masterCategories.getAsync(r => {
      const exists = r.status === Office.AsyncResultStatus.Succeeded &&
        (r.value || []).some(c => c.displayName === name);
      if (exists) return resolve();
      Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: name, color: Office.MailboxEnums.CategoryColor.Preset0 }],
        () => resolve());
    });
  });
  await new Promise((resolve, reject) => {
    Office.context.mailbox.item.categories.addAsync([name], r => {
      r.status === Office.AsyncResultStatus.Succeeded ? resolve()
        : reject(new Error("Kategorie konnte nicht gesetzt werden: " + r.error.message));
    });
  });
}

function loadProps() {
  return new Promise(resolve => {
    Office.context.mailbox.item.loadCustomPropertiesAsync(r => {
      customProps = r.status === Office.AsyncResultStatus.Succeeded ? r.value : null;
      resolve(customProps);
    });
  });
}

function saveProps() {
  return new Promise(resolve => customProps ? customProps.saveAsync(() => resolve()) : resolve());
}

/* ---------- Lesemodus: Ablegen ---------- */

async function fileToAkte() {
  if (!selected) return;
  const item = Office.context.mailbox.item;
  const withAttachments = $("opt-attachments").checked;

  $("btn-file").disabled = true;
  banner("");
  let schritt = "Start";
  try {
    schritt = "Nachricht lesen";
    setStatus("Nachricht wird gelesen …");
    const mime = await fetchMime();

    const base = baseName(item);
    schritt = "Mail hochladen";
    setStatus("Mail wird hochgeladen …");
    await DBX.upload(selected.path + "/" + base + ".eml", mime);

    let count = 0;
    if (withAttachments) {
      const list = (item.attachments || []).filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline);
      for (let i = 0; i < list.length; i++) {
        const a = list[i];
        schritt = "Anhang " + (i + 1) + " lesen/hochladen";
        setStatus("Anhang " + (i + 1) + " von " + list.length + " …");
        const content = await attachmentContent(a.id);
        if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) continue;
        const name = sanitize(a.name) || ("Anhang" + (i + 1));
        await DBX.upload(selected.path + "/" + base + "_A" + (i + 1) + "_" + name, b64ToBytes(content.content));
        count++;
      }
    }

    schritt = "Kategorie setzen";
    setStatus("Mail wird markiert …");
    await ensureCategory();

    schritt = "Vermerk speichern";
    if (customProps) {
      customProps.set(PROP_KEY, JSON.stringify({ path: selected.path, label: selected.label, at: new Date().toISOString() }));
      await saveProps();
    }

    setStatus("");
    banner("Abgelegt in " + selected.label + (withAttachments ? " (Mail + " + count + " Anhänge)" : " (nur die Mail)"), "ok");
  } catch (e) {
    setStatus("");
    banner("Schritt „" + schritt + "\" fehlgeschlagen: " + (e.message || String(e)), "err");
    console.error("[Akte]", schritt, e);
    $("btn-file").disabled = false;
  }
}

/* ---------- Verfassen-Modus: Anhang aus der Akte ---------- */

function renderFiles() {
  const q = $("file-filter").value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const hits = akteFiles.filter(f => {
    const hay = f.name.toLowerCase();
    return words.every(w => hay.includes(w));
  }).slice(0, 200);

  const ul = $("file-list");
  ul.innerHTML = "";
  for (const f of hits) {
    const li = document.createElement("li");
    li.textContent = f.name;
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = fmtSize(f.size) + (f.server_modified ? " · " + f.server_modified.slice(0, 10) : "");
    li.appendChild(meta);
    li.onclick = () => attachToMail(f);
    ul.appendChild(li);
  }
  if (!hits.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = akteFiles.length ? "Keine Datei passt zu diesem Suchbegriff." : "Diese Akte enthält keine Dateien.";
    ul.appendChild(li);
  }
}

async function openAkteFiles(akte) {
  $("main").className = "hidden";
  $("files").className = "";
  $("files-title").textContent = akte.label;
  banner("", null, "files-banner");
  $("files-status").textContent = "Dateien werden geladen …";
  try {
    const entries = await DBX.listChildren(akte.path);
    akteFiles = entries
      .filter(e => e[".tag"] === "file" && e.name !== CONFIG.BETEILIGTE_DATEI)
      .sort((a, b) => b.name.localeCompare(a.name, "de")); // Datumspräfix: neueste oben
    $("file-filter").value = "";
    renderFiles();
    $("files-status").textContent = "Klick auf eine Datei hängt sie an die Mail an.";
  } catch (e) {
    banner("Dateien konnten nicht geladen werden: " + e.message, "err", "files-banner");
    $("files-status").textContent = "";
  }
}

async function attachToMail(f) {
  const maxBytes = (CONFIG.ATTACH_MAX_MB || 20) * 1048576;
  if (f.size > maxBytes) {
    banner(f.name + " ist " + fmtSize(f.size) + " groß – als Mail-Anhang ungeeignet (Grenze " +
      (CONFIG.ATTACH_MAX_MB || 20) + " MB). Besser verkleinern oder auf anderem Weg teilen.", "err", "files-banner");
    return;
  }
  banner("", null, "files-banner");
  try {
    $("files-status").textContent = f.name + " wird geladen …";
    const buf = await DBX.download(f.path_display);
    $("files-status").textContent = f.name + " wird angehängt …";
    const b64 = bytesToB64(buf);
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.addFileAttachmentFromBase64Async(b64, f.name, r => {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve()
          : reject(new Error(r.error.message));
      });
    });
    $("files-status").textContent = "";
    banner(f.name + " ist angehängt. Weitere Datei anklicken oder Panel schließen.", "ok", "files-banner");
  } catch (e) {
    $("files-status").textContent = "";
    banner("Anhängen fehlgeschlagen: " + (e.message || String(e)), "err", "files-banner");
  }
}

/* ---------- Aktenliste ---------- */

async function loadAkten(force) {
  setStatus("Aktenliste wird geladen …");
  try {
    if (!force) {
      const cached = JSON.parse(localStorage.getItem(AKTEN_CACHE_KEY) || "null");
      if (cached && Date.now() - cached.at < 24 * 3600 * 1000) {
        akten = cached.akten; renderList(); setStatus(akten.length + " Akten");
        return;
      }
    }
    akten = await DBX.loadAkten();
    try { localStorage.setItem(AKTEN_CACHE_KEY, JSON.stringify({ at: Date.now(), akten })); } catch {}
    renderList();
    setStatus(akten.length + " Akten");
  } catch (e) {
    setStatus("");
    banner("Aktenliste konnte nicht geladen werden: " + e.message, "err");
  }
}

/* ---------- Start ---------- */

async function refreshForCurrentItem() {
  banner("");
  setStatus("");
  selected = null;
  $("btn-file").disabled = true;
  renderList();

  const item = Office.context.mailbox.item;
  if (!item) { banner("Keine Nachricht ausgewählt.", "info"); return; }
  if (composeMode) return;

  const props = await loadProps();
  if (!props) return;
  const raw = props.get(PROP_KEY);
  if (!raw) return;
  try {
    const prev = JSON.parse(raw);
    banner("Diese Mail wurde bereits abgelegt: " + prev.label + ". Erneutes Ablegen legt eine zweite Kopie an.", "info");
  } catch {}
}

async function showMain() {
  $("signin").className = "hidden";
  $("main").className = "";
  await loadAkten(false);
  await refreshForCurrentItem();

  if (!composeMode) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => { refreshForCurrentItem().catch(e => banner(e.message, "err")); },
      r => { if (r.status !== Office.AsyncResultStatus.Succeeded) console.warn("ItemChanged nicht registriert:", r.error); }
    );
  }
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;

  const item = Office.context.mailbox.item;
  composeMode = !!(item && typeof item.addFileAttachmentFromBase64Async === "function");

  if (composeMode) {
    $("title").textContent = "Anhang aus Akte";
    $("filter").placeholder = "Akte suchen …";
    $("opt-row").classList.add("hidden");
    $("btn-file").classList.add("hidden");
  }

  $("filter").addEventListener("input", renderList);
  $("file-filter").addEventListener("input", renderFiles);
  $("btn-file").addEventListener("click", fileToAkte);
  $("btn-reload").addEventListener("click", () => loadAkten(true));
  $("btn-back").addEventListener("click", () => {
    $("files").className = "hidden";
    $("main").className = "";
  });
  $("btn-login").addEventListener("click", async () => {
    try { await DBX.login(); await showMain(); }
    catch (e) { banner(e.message, "err"); }
  });

  if (DBX.isLoggedIn()) showMain();
  else { $("signin").className = ""; $("main").className = "hidden"; $("files").className = "hidden"; }
});
