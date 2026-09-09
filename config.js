/* Zentrale Konfiguration. Einzige Datei, die du anfassen musst, wenn sich etwas ändert. */

const CONFIG = {
  // Dropbox-App "vC-Aktenablage", Development-Modus, Zugriffstyp: Full Dropbox.
  // Kein Geheimnis - der App key steckt in jeder ausgelieferten Client-App.
  DROPBOX_APP_KEY: "a82pf6r4daprb46",

  // Die drei Bereiche. Reihenfolge = Reihenfolge im Picker.
  ROOTS: [
    { label: "Charly und Franz", path: "/01_Charly und Franz" },
    { label: "Hiddenhausen",     path: "/02_Hiddenhausen" },
    { label: "Süschendorf",      path: "/03_Süschendorf" },
  ],

  // Ordnernamen, die nie als Ablageziel angeboten werden (exakter Name, auf jeder Ebene).
  EXCLUDE: ["00_Claude", "_Archiv", "_to_delete", "Claude outputs", "Ad Acta"],

  // Wie tief unter den Bereichen Ordner angeboten werden (1 = nur Themenordner, 2 = auch deren Unterordner).
  DEPTH: 2,

  // Betreff im Dateinamen auf so viele Zeichen kürzen (Windows-Pfadgrenze von 260 Zeichen).
  SUBJECT_MAX: 60,

  // Name der Outlook-Kategorie, die eine abgelegte Mail bekommt.
  CATEGORY: "Abgelegt",
};
