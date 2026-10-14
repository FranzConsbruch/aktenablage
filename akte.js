/* Aktenseite: Akte wählen, Dokumente, Beteiligte, Wiedervorlagen, Brief erzeugen.
   Reines Frontend gegen die Dropbox-API. Läuft auf jedem statischen Hosting;
   die Adresse der Seite muss als Redirect-URI in der Dropbox-App eingetragen sein. */

const AKTEN_CACHE_KEY = "akten_cache_v2";

let akten = [];
let currentAkte = null;   // { label, path, root }
let beteiligte = [];      // Inhalt von beteiligte.json der aktuellen Akte
let betEditIndex = -1;
let dokumente = [];
let wvAll = [];           // alle offenen Wiedervorlagen über alle Bereiche

const $ = id => document.getElementById(id);

/* ---------- Kleinkram ---------- */

function banner(elId, text, kind) {
  const b = $(elId);
  b.innerHTML = "";
  if (!text) return;
  const d = document.createElement("div");
  d.className = "note " + (kind || "info");
  d.textContent = text;
  b.appendChild(d);
}

function fmtSize(n) {
  if (n == null) return "";
  if (n < 1024) return n + " B";
  if (n < 1048576) return (n / 1024).toFixed(0) + " KB";
  return (n / 1048576).toFixed(1) + " MB";
}

function fmtDate(iso) { // "2026-10-20" -> "20.10.2026"
  if (!iso || iso.length < 10) return iso || "";
  return iso.slice(8, 10) + "." + iso.slice(5, 7) + "." + iso.slice(0, 4);
}

function todayISO() {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
}

function sanitize(s) {
  return (s || "")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+|\.+$/g, "");
}

function rootOf(path) {
  const r = CONFIG.ROOTS.find(r => path === r.path || path.startsWith(r.path + "/"));
  return r ? r.path : null;
}

function show(view) {
  for (const v of ["v-login", "v-home", "v-akte"]) $(v).classList.add("hidden");
  $(view).classList.remove("hidden");
  $("btn-logout").classList.toggle("hidden", view === "v-login");
}

const enc = new TextEncoder();

/* ---------- Startansicht ---------- */

async function loadAkten(force) {
  $("home-status").textContent = "Aktenliste wird geladen …";
  try {
    if (!force) {
      const cached = JSON.parse(localStorage.getItem(AKTEN_CACHE_KEY) || "null");
      if (cached && Date.now() - cached.at < 24 * 3600 * 1000) {
        akten = cached.akten; renderAkten(); $("home-status").textContent = akten.length + " Akten"; return;
      }
    }
    akten = await DBX.loadAkten();
    try { localStorage.setItem(AKTEN_CACHE_KEY, JSON.stringify({ at: Date.now(), akten })); } catch {}
    renderAkten();
    $("home-status").textContent = akten.length + " Akten";
  } catch (e) {
    $("home-status").textContent = "";
    banner("wv-banner", "Aktenliste konnte nicht geladen werden: " + e.message, "err");
  }
}

function renderAkten() {
  const q = $("home-filter").value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const hits = akten.filter(a => words.every(w => a.label.toLowerCase().includes(w))).slice(0, 200);
  const ul = $("home-list");
  ul.innerHTML = "";
  for (const a of hits) {
    const li = document.createElement("li");
    li.textContent = a.label;
    li.onclick = () => openAkte(a);
    ul.appendChild(li);
  }
  if (!hits.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = akten.length ? "Keine Akte passt zu diesem Suchbegriff." : "Noch keine Akten geladen.";
    ul.appendChild(li);
  }
}

/* ---------- Wiedervorlagen (zentral je Bereich, eine Datei je Eintrag) ---------- */

async function loadWiedervorlagen() {
  wvAll = [];
  for (const root of CONFIG.ROOTS) {
    let entries;
    try { entries = await DBX.listChildren(root.path + "/" + CONFIG.WIEDERVORLAGEN_DIR); }
    catch { continue; } // Ordner existiert noch nicht
    const files = entries.filter(e => e[".tag"] === "file" && e.name.endsWith(".json"));
    const loaded = await Promise.all(files.map(async f => {
      try {
        const buf = await DBX.download(f.path_display);
        const j = JSON.parse(new TextDecoder().decode(buf));
        return { ...j, file: f.path_display, rootLabel: root.label };
      } catch { return null; }
    }));
    wvAll.push(...loaded.filter(Boolean));
  }
  wvAll.sort((a, b) => (a.datum || "").localeCompare(b.datum || ""));
}

function wvLi(e, showAkte) {
  const li = document.createElement("li");
  const overdue = e.datum && e.datum <= todayISO();
  const date = document.createElement("span");
  date.textContent = fmtDate(e.datum) + "  ";
  if (overdue) date.className = "overdue";
  li.appendChild(date);
  li.appendChild(document.createTextNode(e.notiz || "(ohne Notiz)"));
  const meta = document.createElement("span");
  meta.className = "meta";
  meta.textContent = showAkte ? (e.akteLabel || e.aktePath || "") : "";
  if (showAkte) li.appendChild(meta);
  const done = document.createElement("button");
  done.className = "mini ghost";
  done.style.float = "right";
  done.textContent = "Erledigt";
  done.onclick = async ev => {
    ev.stopPropagation();
    done.disabled = true;
    try { await wvErledigen(e); } catch (err) { alert("Erledigen fehlgeschlagen: " + err.message); done.disabled = false; }
  };
  li.insertBefore(done, li.firstChild);
  return li;
}

function renderWvHome() {
  const ul = $("wv-all");
  ul.innerHTML = "";
  for (const e of wvAll) {
    const li = wvLi(e, true);
    li.onclick = () => {
      const a = akten.find(a => a.path === e.aktePath) ||
        { label: e.akteLabel || e.aktePath, path: e.aktePath, root: rootOf(e.aktePath) };
      if (a.path) openAkte(a);
    };
    ul.appendChild(li);
  }
  if (!wvAll.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "Keine offenen Wiedervorlagen.";
    ul.appendChild(li);
  }
}

function renderWvAkte() {
  const ul = $("wv-list");
  ul.innerHTML = "";
  const mine = wvAll.filter(e => e.aktePath === currentAkte.path);
  for (const e of mine) ul.appendChild(wvLi(e, false));
  if (!mine.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "Keine offenen Wiedervorlagen in dieser Akte.";
    ul.appendChild(li);
  }
}

async function wvAnlegen() {
  const datum = $("wv-datum").value;
  const notiz = $("wv-notiz").value.trim();
  if (!datum) { banner("akte-banner", "Bitte ein Datum wählen.", "err"); return; }
  const root = currentAkte.root || rootOf(currentAkte.path);
  if (!root) { banner("akte-banner", "Bereich der Akte nicht erkannt.", "err"); return; }
  const id = Math.random().toString(36).slice(2, 8);
  const entry = {
    datum, notiz,
    akteLabel: currentAkte.label, aktePath: currentAkte.path,
    angelegt: new Date().toISOString(),
  };
  const path = root + "/" + CONFIG.WIEDERVORLAGEN_DIR + "/" + datum + "__" + id + ".json";
  await DBX.upload(path, enc.encode(JSON.stringify(entry, null, 2)));
  $("wv-notiz").value = "";
  await loadWiedervorlagen();
  renderWvAkte(); renderWvHome();
  banner("akte-banner", "Wiedervorlage für den " + fmtDate(datum) + " angelegt.", "ok");
}

async function wvErledigen(e) {
  const root = rootOf(e.file);
  const name = e.file.split("/").pop();
  const ziel = root + "/" + CONFIG.WIEDERVORLAGEN_DIR + "/erledigt/" + name;
  const body = enc.encode(JSON.stringify({ ...e, erledigt: new Date().toISOString() }, null, 2));
  await DBX.uploadOverwrite(ziel, body);
  await DBX.deleteFile(e.file);
  wvAll = wvAll.filter(x => x.file !== e.file);
  renderWvHome();
  if (currentAkte) renderWvAkte();
}

/* ---------- Aktenansicht ---------- */

async function openAkte(a) {
  currentAkte = { ...a, root: a.root || rootOf(a.path) };
  show("v-akte");
  $("akte-title").textContent = a.label;
  $("akte-path").textContent = "Dropbox: " + a.path;
  banner("akte-banner", "");
  selectTab("dok");

  dokumente = []; renderDok();
  beteiligte = []; betEditIndex = -1; renderBet(); betFormReset();

  try {
    const entries = await DBX.listChildren(a.path);
    dokumente = entries
      .filter(e => e[".tag"] === "file" && e.name !== CONFIG.BETEILIGTE_DATEI)
      .sort((x, y) => y.name.localeCompare(x.name, "de"));
    renderDok();
  } catch (e) {
    banner("akte-banner", "Dokumente konnten nicht geladen werden: " + e.message, "err");
  }

  try {
    const buf = await DBX.downloadIfExists(a.path + "/" + CONFIG.BETEILIGTE_DATEI);
    if (buf) {
      const j = JSON.parse(new TextDecoder().decode(buf));
      beteiligte = Array.isArray(j) ? j : (j.beteiligte || []);
    }
    renderBet(); fillBriefDropdowns();
  } catch (e) {
    banner("akte-banner", "Beteiligtenliste konnte nicht gelesen werden: " + e.message, "err");
  }

  renderWvAkte();
  loadVorlagen();
}

function selectTab(tab) {
  for (const b of document.querySelectorAll(".tabs button")) {
    b.setAttribute("aria-selected", b.dataset.tab === tab ? "true" : "false");
  }
  for (const id of ["dok", "bet", "wv", "brief"]) {
    $("tab-" + id).classList.toggle("hidden", id !== tab);
  }
}

/* ---------- Dokumente ---------- */

function renderDok() {
  const q = $("dok-filter").value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  const hits = dokumente.filter(f => words.every(w => f.name.toLowerCase().includes(w))).slice(0, 300);
  const ul = $("dok-list");
  ul.innerHTML = "";
  for (const f of hits) {
    const li = document.createElement("li");
    li.textContent = f.name;
    const meta = document.createElement("span");
    meta.className = "meta";
    meta.textContent = fmtSize(f.size) + (f.server_modified ? " · " + fmtDate(f.server_modified.slice(0, 10)) : "");
    li.appendChild(meta);
    li.onclick = async () => {
      try { window.open(await DBX.tempLink(f.path_display), "_blank"); }
      catch (e) { banner("akte-banner", "Öffnen fehlgeschlagen: " + e.message, "err"); }
    };
    ul.appendChild(li);
  }
  if (!hits.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = dokumente.length ? "Keine Datei passt zu diesem Suchbegriff." : "Diese Akte enthält keine Dateien.";
    ul.appendChild(li);
  }
}

/* ---------- Beteiligte ---------- */

const BET_FIELDS = ["name", "rolle", "anrede", "firma", "telefon", "strasse", "plzort", "email"];

function renderBet() {
  const tb = $("bet-table").querySelector("tbody");
  tb.innerHTML = "";
  beteiligte.forEach((b, i) => {
    const tr = document.createElement("tr");
    const tdName = document.createElement("td");
    tdName.textContent = b.name || "";
    if (b.firma) { const s = document.createElement("div"); s.className = "muted small"; s.textContent = b.firma; tdName.appendChild(s); }
    const tdRolle = document.createElement("td"); tdRolle.textContent = b.rolle || "";
    const tdKontakt = document.createElement("td"); tdKontakt.textContent = [b.email, b.telefon].filter(Boolean).join(" · ");
    const tdAdr = document.createElement("td"); tdAdr.textContent = [b.strasse, b.plzort].filter(Boolean).join(", ");
    const tdBtn = document.createElement("td");
    const be = document.createElement("button"); be.className = "mini ghost"; be.textContent = "Bearbeiten";
    be.onclick = () => betEdit(i);
    const bd = document.createElement("button"); bd.className = "mini ghost"; bd.textContent = "Entfernen";
    bd.style.marginLeft = "4px";
    bd.onclick = () => { if (confirm((b.name || "Eintrag") + " aus der Beteiligtenliste entfernen?")) { beteiligte.splice(i, 1); betSpeichern(); } };
    tdBtn.append(be, bd);
    tr.append(tdName, tdRolle, tdKontakt, tdAdr, tdBtn);
    tb.appendChild(tr);
  });
  if (!beteiligte.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 5; td.className = "muted";
    td.textContent = "Noch keine Beteiligten in dieser Akte.";
    tr.appendChild(td); tb.appendChild(tr);
  }
}

function betFormReset() {
  betEditIndex = -1;
  for (const f of BET_FIELDS) $("bet-" + f).value = "";
  $("bet-form-title").textContent = "Beteiligten hinzufügen";
  $("bet-cancel").classList.add("hidden");
}

function betEdit(i) {
  betEditIndex = i;
  for (const f of BET_FIELDS) $("bet-" + f).value = beteiligte[i][f] || "";
  $("bet-form-title").textContent = "Beteiligten bearbeiten";
  $("bet-cancel").classList.remove("hidden");
  selectTab("bet");
}

async function betSpeichernAusFormular() {
  const b = {};
  for (const f of BET_FIELDS) b[f] = $("bet-" + f).value.trim();
  if (!b.name) { banner("akte-banner", "Ein Beteiligter braucht mindestens einen Namen.", "err"); return; }
  if (betEditIndex >= 0) beteiligte[betEditIndex] = b; else beteiligte.push(b);
  await betSpeichern();
  betFormReset();
}

async function betSpeichern() {
  try {
    const body = enc.encode(JSON.stringify({ beteiligte }, null, 2));
    await DBX.uploadOverwrite(currentAkte.path + "/" + CONFIG.BETEILIGTE_DATEI, body);
    renderBet(); fillBriefDropdowns();
    banner("akte-banner", "Beteiligtenliste gespeichert.", "ok");
  } catch (e) {
    banner("akte-banner", "Speichern fehlgeschlagen: " + e.message, "err");
  }
}

/* ---------- Brief erzeugen ---------- */

let vorlagen = [];

async function loadVorlagen() {
  vorlagen = [];
  const root = currentAkte.root || rootOf(currentAkte.path);
  if (root) {
    try {
      const entries = await DBX.listChildren(root + "/" + CONFIG.VORLAGEN_DIR);
      vorlagen = entries.filter(e => e[".tag"] === "file" && /\.docx$/i.test(e.name));
    } catch { /* Ordner existiert noch nicht */ }
  }
  fillBriefDropdowns();
}

function fillBriefDropdowns() {
  const sv = $("brief-vorlage");
  sv.innerHTML = "";
  if (!vorlagen.length) {
    sv.appendChild(new Option("Keine Vorlagen gefunden – .docx in den Ordner " + CONFIG.VORLAGEN_DIR + " des Bereichs legen", ""));
  } else {
    for (const v of vorlagen) sv.appendChild(new Option(v.name, v.path_display));
  }
  const sb = $("brief-beteiligter");
  sb.innerHTML = "";
  if (!beteiligte.length) {
    sb.appendChild(new Option("Keine Beteiligten in dieser Akte", ""));
  } else {
    beteiligte.forEach((b, i) => sb.appendChild(new Option(b.name + (b.rolle ? " (" + b.rolle + ")" : ""), String(i))));
  }
}

function xmlEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

/* Ersetzt {{Platzhalter}} im document.xml. Word zerlegt getippten Text gern in
   mehrere Läufe; das Muster erlaubt deshalb XML-Tags zwischen den Klammern. */
function mergeXml(xml, data) {
  const lower = {};
  for (const k of Object.keys(data)) lower[k.toLowerCase()] = data[k];
  let count = 0;
  const out = xml.replace(/\{\{(?:[^{}<]|<[^>]+>)*?\}\}/g, m => {
    const key = m.replace(/<[^>]+>/g, "").replace(/[{}]/g, "").trim().toLowerCase();
    if (key in lower) { count++; return xmlEscape(lower[key] ?? ""); }
    return m;
  });
  return { out, count };
}

async function briefErzeugen() {
  const vPath = $("brief-vorlage").value;
  const bIdx = $("brief-beteiligter").value;
  if (!vPath) { banner("akte-banner", "Keine Vorlage gewählt.", "err"); return; }
  if (bIdx === "") { banner("akte-banner", "Kein Beteiligter gewählt – erst unter „Beteiligte“ anlegen.", "err"); return; }
  const b = beteiligte[Number(bIdx)];
  const st = $("brief-status");
  $("brief-go").disabled = true;
  try {
    st.textContent = "Vorlage wird geladen …";
    const buf = await DBX.download(vPath);
    const zip = await JSZip.loadAsync(buf);
    const docFile = zip.file("word/document.xml");
    if (!docFile) throw new Error("Vorlage ist keine gültige Word-Datei.");
    const xml = await docFile.async("string");

    const d = new Date(), p = n => String(n).padStart(2, "0");
    const data = {
      Anrede: b.anrede || "Sehr geehrte Damen und Herren",
      Name: b.name || "", Firma: b.firma || "",
      Strasse: b.strasse || "", PLZOrt: b.plzort || "",
      Email: b.email || "", Telefon: b.telefon || "",
      Datum: p(d.getDate()) + "." + p(d.getMonth() + 1) + "." + d.getFullYear(),
      Akte: currentAkte.label.split(" › ").pop(),
      Betreff: $("brief-betreff").value.trim(),
    };
    const { out, count } = mergeXml(xml, data);
    zip.file("word/document.xml", out);

    st.textContent = "Brief wird erzeugt …";
    const result = await zip.generateAsync({ type: "arraybuffer" });

    const vBase = $("brief-vorlage").selectedOptions[0].textContent.replace(/\.docx$/i, "");
    const name = todayISO() + "_" + sanitize(b.name) + "_" + sanitize(vBase) + ".docx";
    st.textContent = "Brief wird in der Akte gespeichert …";
    const up = await DBX.upload(currentAkte.path + "/" + name, result);

    // Zusätzlich als Download anbieten
    const blob = new Blob([result], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = up.name || name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 30000);

    st.textContent = "";
    banner("akte-banner", "Brief „" + (up.name || name) + "“ liegt in der Akte (" + count + " Platzhalter gefüllt) und wurde heruntergeladen.", "ok");
    dokumente.unshift({ ".tag": "file", name: up.name || name, path_display: up.path_display, size: result.byteLength, server_modified: new Date().toISOString() });
    renderDok();
  } catch (e) {
    st.textContent = "";
    banner("akte-banner", "Brief konnte nicht erzeugt werden: " + e.message, "err");
  } finally {
    $("brief-go").disabled = false;
  }
}

/* ---------- Start ---------- */

async function enterHome() {
  show("v-home");
  await loadAkten(false);
  renderWvHome();
  try { await loadWiedervorlagen(); renderWvHome(); }
  catch (e) { banner("wv-banner", "Wiedervorlagen konnten nicht geladen werden: " + e.message, "err"); }
}

(async function init() {
  $("home-filter").addEventListener("input", renderAkten);
  $("dok-filter").addEventListener("input", renderDok);
  $("btn-reload-akten").addEventListener("click", () => loadAkten(true));
  $("btn-home").addEventListener("click", () => { currentAkte = null; show("v-home"); renderWvHome(); });
  $("bet-save").addEventListener("click", betSpeichernAusFormular);
  $("bet-cancel").addEventListener("click", betFormReset);
  $("wv-save").addEventListener("click", () => wvAnlegen().catch(e => banner("akte-banner", e.message, "err")));
  $("brief-go").addEventListener("click", briefErzeugen);
  $("btn-logout").addEventListener("click", () => { DBX.forget(); location.reload(); });
  $("btn-login").addEventListener("click", () => DBX.loginWeb());
  $("wv-datum").min = todayISO();
  for (const b of document.querySelectorAll(".tabs button")) {
    b.addEventListener("click", () => selectTab(b.dataset.tab));
  }

  try { await DBX.completeWebLogin(); }
  catch (e) { show("v-login"); banner("login-banner", e.message, "err"); return; }

  if (DBX.isLoggedIn()) enterHome();
  else show("v-login");
})();
