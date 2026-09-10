/* Taskpane: Akte wählen, Mail als .eml plus Anhänge nach Dropbox schreiben,
   Mail mit der Kategorie "Abgelegt" markieren und den Zielordner an der Mail vermerken. */

const AKTEN_CACHE_KEY = "akten_cache_v1";
const PROP_KEY = "ablage"; // benutzerdefinierte Eigenschaft an der Mail

let akten = [];
let selected = null;
let customProps = null;

const $ = id => document.getElementById(id);

/* ---------- Anzeige ---------- */

function setStatus(text) { $("status").textContent = text || ""; }

function banner(text, kind) {
  const b = $("banner");
  if (!text) { b.className = "hidden"; b.textContent = ""; return; }
  b.className = "note " + (kind || "info");
  b.textContent = text;
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
    li.onclick = () => { selected = a; renderList(); $("btn-file").disabled = false; };
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
    .replace(/[\\/:*?"<>|]/g, "-")   // unter Windows verbotene Zeichen
    .replace(/[\u0000-\u001f\u007f]/g, "")   // Steuerzeichen
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+|\.+$/g, "");      // führende/abschließende Punkte
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

/* ---------- Outlook ---------- */

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

/* Alter Weg über die Outlook-REST-Schnittstelle. Funktioniert nur, solange im
   Tenant die "Legacy Exchange Online Tokens" aktiv sind - in neuen Tenants sind
   sie ab Werk aus, dann scheitert getCallbackTokenAsync mit Code 9018
   ("Ein interner Fehler ist aufgetreten"). Deshalb nur noch Rückfallebene. */
async function emlViaRest() {
  const token = await restToken();
  const restId = Office.context.mailbox.convertToRestId(
    Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  const url = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + restId + "/$value";
  const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) throw new Error("Nachricht konnte nicht gelesen werden (" + res.status + ").");
  return res.arrayBuffer();
}

/* Holt die Nachricht als rohes MIME - das ist bereits eine .eml-Datei. */
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

function b64ToBytes(b64) {
  const bin = atob(b64);
  const a = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
  return a;
}

async function ensureCategory() {
  const name = CONFIG.CATEGORY;
  await new Promise(resolve => {
    Office.context.mailbox.masterCategories.getAsync(r => {
      const exists = r.status === Office.AsyncResultStatus.Succeeded &&
        (r.value || []).some(c => c.displayName === name);
      if (exists) return resolve();
      Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: name, color: Office.MailboxEnums.CategoryColor.Preset0 }],
        () => resolve()); // schlägt das fehl, versuchen wir das Setzen trotzdem
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

/* ---------- Ablauf ---------- */

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

/* Zeigt für die gerade ausgewählte Mail an, ob sie schon abgelegt wurde,
   und setzt Auswahl und Knopf zurück. Läuft beim Öffnen und bei jedem Mailwechsel,
   damit das angeheftete Panel mitwandert. */
async function refreshForCurrentItem() {
  banner("");
  setStatus("");
  selected = null;
  $("btn-file").disabled = true;
  renderList();

  const item = Office.context.mailbox.item;
  if (!item) { banner("Keine Nachricht ausgewählt.", "info"); return; }

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

  /* Angeheftetes Panel: Outlook meldet den Wechsel der ausgewählten Mail. */
  Office.context.mailbox.addHandlerAsync(
    Office.EventType.ItemChanged,
    () => { refreshForCurrentItem().catch(e => banner(e.message, "err")); },
    r => { if (r.status !== Office.AsyncResultStatus.Succeeded) console.warn("ItemChanged nicht registriert:", r.error); }
  );
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Outlook) return;

  $("filter").addEventListener("input", renderList);
  $("btn-file").addEventListener("click", fileToAkte);
  $("btn-reload").addEventListener("click", () => loadAkten(true));
  $("btn-login").addEventListener("click", async () => {
    try { await DBX.login(); await showMain(); }
    catch (e) { banner(e.message, "err"); }
  });

  if (DBX.isLoggedIn()) showMain();
  else { $("signin").className = ""; $("main").className = "hidden"; }
});
