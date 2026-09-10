# In Akte ablegen — Outlook-Add-in

Legt die geöffnete Mail als `.eml` samt Anhängen in einer Dropbox-Akte ab und markiert sie
in Outlook mit der Kategorie **Abgelegt**. Ersetzt den Kleos-Handgriff „Mail markieren → Akte wählen".

Kein Server, keine Datenbank, kein Hintergrunddienst: reine statische Dateien, alles passiert
im Moment des Klicks.

## Voraussetzungen

- **Exchange-/Microsoft-365-Postfach.** Add-ins funktionieren nicht mit Gmail- oder IMAP-Konten —
  auch nicht im neuen Outlook. Zum Testen also `@vonconsbruch.de` verwenden.
- Dropbox-App im Development-Modus, Zugriffstyp **Full Dropbox**, App key steht in `config.js`.

## Einrichtung

**1. Repository und GitHub Pages**

Diese Dateien in das Repository `FranzConsbruch/aktenablage` legen, dann unter
*Settings → Pages* die Quelle auf den Branch `main`, Ordner `/` stellen.
Nach ein bis zwei Minuten ist alles erreichbar unter
`https://franzconsbruch.github.io/aktenablage/`.

**2. Redirect URI in der Dropbox-App eintragen**

Unter *dropbox.com/developers/apps → deine App → Settings → OAuth 2 → Redirect URIs*:

```
https://franzconsbruch.github.io/aktenablage/auth.html
```

Exakt so, inklusive `auth.html`. Im selben Reiter unter *Permissions* müssen gesetzt und
mit **Submit** gespeichert sein: `files.metadata.read`, `files.metadata.write`,
`files.content.read`, `files.content.write`.

**3. Manifest anpassen — bereits erledigt**

In `manifest.xml` stehen die Adressen schon drin
(`https://franzconsbruch.github.io/aktenablage`). Nichts zu tun.

**4. Add-in bereitstellen**

Über das Admin Center (empfohlen, und der Weg für weitere Nutzer):
**admin.microsoft.com → Einstellungen → Integrierte Apps → Benutzerdefinierte Apps hochladen**
→ *Office-Add-In* → `manifest.xml`. Danach beim Zuweisen die gewünschten Benutzer wählen.
Die Bereitstellung braucht bis zu 24 Stunden, meist 15 bis 60 Minuten.

Alternativ als Endnutzer: eine Mail öffnen → **Apps** in der Nachrichten-Symbolleiste →
im Dialog *Add-Ins für Outlook* unter **Meine Add-Ins → Benutzerdefinierte Add-Ins →
Benutzerdefiniertes Add-In hinzufügen → Aus Datei hinzufügen**. In den Outlook-Einstellungen
gibt es diesen Punkt **nicht** – dort sucht man vergeblich.

Der Knopf erscheint anschließend im Menüband einer geöffneten Mail unter **Weitere Apps**.

**5. Erster Lauf**

Eine Mail öffnen → Schaltfläche **In Akte ablegen** → *Bei Dropbox anmelden*.
Die Anmeldung läuft einmalig; danach hält sie über das Refresh-Token.

## Änderungen veröffentlichen

Im Projektordner liegt **`push.cmd`**. Doppelklick genügt: Beim ersten Lauf legt es das lokale
Repository an, verbindet es mit `github.com/FranzConsbruch/aktenablage` und übernimmt die
vorhandene Historie; danach committet und pusht es alles Geänderte. Für eine eigene
Commit-Nachricht in der Eingabeaufforderung `push.cmd "Was geaendert wurde"` aufrufen.

Beim allerersten Push fragt Git nach der GitHub-Anmeldung — im Browser bestätigen und
`push.cmd` noch einmal starten. GitHub Pages übernimmt die Änderung nach ein bis zwei Minuten;
danach im Outlook das Panel schließen und neu öffnen.

**Wichtig:** Änderungen an `manifest.xml` wirken **nicht** über Pages. Das Manifest muss im
Admin Center unter *Integrierte Apps* neu hochgeladen werden — und dort auch nur, wenn sich
das Manifest tatsächlich geändert hat. Änderungen an HTML, JS und Symbolen brauchen nur den Push.

## Dateien

| Datei | Zweck |
|---|---|
| `manifest.xml` | Add-in-Manifest, meldet Schaltfläche und Taskpane bei Outlook an |
| `taskpane.html` | Oberfläche: Suchfeld, Aktenliste, Anhang-Schalter |
| `taskpane.js` | Ablauf: MIME holen, hochladen, Anhänge, Kategorie, Vermerk an der Mail |
| `dropbox.js` | Dropbox: Anmeldung mit PKCE, Ordnerliste, Upload |
| `auth.html` | Landeseite der Anmeldung, gibt die Tokens ans Taskpane zurück |
| `config.js` | **Die einzige Datei zum Anpassen:** App key, Bereiche, Ausschlussliste, Namensregeln |
| `assets/` | Symbole |

## Namensregeln

```
2026-09-09_refurbed_Noch unschlüssig.eml
2026-09-09_refurbed_Noch unschlüssig_A1_Rechnung.pdf
```

Datum, Absender, Betreff (auf 60 Zeichen gekürzt), unter Windows verbotene Zeichen ersetzt.
Anhänge liegen **flach daneben**, damit sie nach Datum sortiert direkt bei der Mail stehen
und in der Aktenansicht auftauchen.

## Angeheftetes Panel

Das Taskpane kann **angeheftet** werden (Stecknadel oben im Panel). Dann bleibt die Leiste
rechts offen und wechselt automatisch mit, wenn eine andere Mail ausgewählt wird — dasselbe
Verhalten wie in Kleos. Technisch: `<SupportsPinning>` im Manifest plus ein Handler auf
`Office.EventType.ItemChanged`, der Auswahl, Knopf und den Hinweis „bereits abgelegt" neu setzt.

## Was das Add-in bewusst nicht tut

- **Nichts verschieben.** Die Mail bleibt im Posteingang, sie bekommt nur die Kategorie.
- **Nichts überschreiben.** Uploads laufen mit `autorename`; eine gleichnamige Datei wird
  nie ersetzt, sondern bekommt eine Nummer.
- **Nichts stillschweigend halb erledigen.** Schlägt ein Schritt fehl, bleibt die Mail
  unmarkiert und die Meldung steht im Taskpane.

## Gegenprobe

In Outlook einen Suchordner auf die Kategorie **Abgelegt** anlegen. Dessen Anzahl gegen die
Zahl der `.eml`-Dateien in Dropbox stellen — laufen sie auseinander, fehlt etwas.

## Bekannte Grenzen

- **Nur Lesemodus.** Gesendete Mails legt man vorerst genauso von Hand aus „Gesendete Elemente" ab.
- **Anhängen aus der Akte beim Verfassen** ist noch nicht gebaut (zweite Ausbaustufe).
- Das MIME wird über den REST-Weg von Outlook geholt (`getCallbackTokenAsync`). Microsoft
  bewegt sich langfristig Richtung Graph mit *Nested App Auth*; falls der REST-Weg wegfällt,
  ist `fetchMime()` in `taskpane.js` die einzige Stelle, die sich ändert.
- Ordner, die nicht angeboten werden sollen, stehen in `config.js` unter `EXCLUDE`.
