/**
 * Mysterious Worlds – Reservierungen (Google Apps Script) · Version 3 (mit Logo in der Mail)
 *
 * Einrichtung am Computer ODER iPad (ca. 10 Minuten):
 * 1. Im Browser script.google.com öffnen (iPad: Safari, „Desktop-Website anfordern“)
 *    → Neues Projekt. Den gesamten Inhalt dieser Datei einfügen und speichern.
 *    (Alternativ am Computer: Google-Tabelle → Erweiterungen → Apps Script.)
 * 2. Oben die Funktion "setup" auswählen → Ausführen → Zugriff erlauben.
 *    Das Script legt die Tabelle "Mysterious Worlds – Reservierungen" mit den Blättern
 *    "Reservierungen" und "Einstellungen" selbst an (Link steht im Ausführungsprotokoll).
 * 3. Bereitstellen → Neue Bereitstellung → Typ "Web-App".
 *    Ausführen als: Ich · Zugriff: Jeder → Bereitstellen.
 *    Die Web-App-URL (endet auf /exec) kommt in die Reservierungsseite.
 *
 * Im Alltag (geht auch in der Google-Tabellen-App):
 * - Reservierung öffnen/schließen, Plätze, Beginn, Einlass, Ort und Mail-Adresse
 *   im Blatt "Einstellungen" ändern. Wirkt sofort, ohne neue Bereitstellung.
 * - Stornieren von Hand: in "Reservierungen" die Spalte Status auf "storniert" setzen.
 *   Die Plätze werden automatisch wieder frei.
 * - Nur wenn dieser Code geändert wird: Bereitstellen → Bereitstellungen verwalten →
 *   Bearbeiten → Version "Neue Version" → Bereitstellen (die URL bleibt gleich).
 */

const NIGHTS = [
  { id: "2027-03-17", label: "Mittwoch, 17. März 2027", kurz: "Mi 17.03." },
  { id: "2027-03-18", label: "Donnerstag, 18. März 2027", kurz: "Do 18.03." }
];
const SEITE = "https://jonahweh.github.io/Mysterious-Worlds/reservierung/";
const LOGO = "https://jonahweh.github.io/Mysterious-Worlds/mail-header.jpg"; // Kopfbild der Bestätigungsmail
const RES = "Reservierungen", SET = "Einstellungen";
const COLS = ["Eingang", "Code", "Abend", "Plätze", "Name", "E-Mail", "Hinweis", "Status", "Storno-Token", "Storniert am"];

/* ---------- Tabelle finden (an eine Tabelle gebunden oder eigenständig) ---------- */
function book_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("Noch keine Tabelle: bitte zuerst die Funktion setup ausführen.");
  return SpreadsheetApp.openById(id);
}

/* ---------- Einrichtung ---------- */
function setup() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (!ss) {
    const props = PropertiesService.getScriptProperties(), id = props.getProperty("SHEET_ID");
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create("Mysterious Worlds – Reservierungen"); props.setProperty("SHEET_ID", ss.getId()); }
  }
  let r = ss.getSheetByName(RES);
  if (!r) { r = ss.getSheets()[0]; r.setName(RES); }
  if (r.getLastRow() === 0) {
    r.appendRow(COLS);
    r.getRange(1, 1, 1, COLS.length).setFontWeight("bold").setBackground("#ece6fb");
    r.setFrozenRows(1);
    r.getRange("A:A").setNumberFormat("dd.mm.yyyy hh:mm");
    r.getRange("H2:H").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["aktiv", "storniert"], true).build());
  }
  let s = ss.getSheetByName(SET);
  if (!s) {
    s = ss.insertSheet(SET);
    const me = Session.getEffectiveUser().getEmail();
    const rows = [
      ["Reservierung offen", false, "Häkchen setzen, sobald Gäste reservieren dürfen"],
      ["Plätze " + NIGHTS[0].kurz, 300, ""],
      ["Plätze " + NIGHTS[1].kurz, 300, ""],
      ["Max. Plätze pro Reservierung", 6, ""],
      ["Beginn", "", "z. B. 19:00 Uhr – leer lassen = „folgt“"],
      ["Einlass", "", "z. B. ab 18:30 Uhr"],
      ["Ort", "", "z. B. Aula, mit Adresse"],
      ["Benachrichtigung an", me, "An diese Adresse geht bei jeder Reservierung eine Mail"],
      ["Mail bei Stornierung", true, ""],
      ["", "", ""],
      ["Reserviert " + NIGHTS[0].kurz, `=SUMIFS(${RES}!D:D,${RES}!C:C,"${NIGHTS[0].id}",${RES}!H:H,"aktiv")`, "wird automatisch berechnet"],
      ["Reserviert " + NIGHTS[1].kurz, `=SUMIFS(${RES}!D:D,${RES}!C:C,"${NIGHTS[1].id}",${RES}!H:H,"aktiv")`, "wird automatisch berechnet"]
    ];
    s.getRange(1, 1, rows.length, 3).setValues(rows);
    s.getRange("B1").insertCheckboxes(); s.getRange("B1").setValue(false);
    s.getRange("B9").insertCheckboxes(); s.getRange("B9").setValue(true);
    s.getRange("A:A").setFontWeight("bold");
    s.getRange("C:C").setFontColor("#625b78");
    s.setColumnWidth(1, 230); s.setColumnWidth(2, 260); s.setColumnWidth(3, 380);
  }
  r.getRange("C2:C").setNumberFormat("@");
  try { ss.setActiveSheet(s); } catch (e) {}
  console.log("Fertig! Deine Tabelle: " + ss.getUrl());
}

/* ---------- Hilfsfunktionen ---------- */
function settings_() {
  const vals = book_().getSheetByName(SET).getDataRange().getValues();
  const m = {};
  vals.forEach(r => { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  const cap = {};
  NIGHTS.forEach(n => { cap[n.id] = Number(m["Plätze " + n.kurz]) || 0; });
  return {
    open: m["Reservierung offen"] === true,
    cap: cap,
    max: Math.max(1, Number(m["Max. Plätze pro Reservierung"]) || 6),
    beginn: String(m["Beginn"] || "").trim(),
    einlass: String(m["Einlass"] || "").trim(),
    ort: String(m["Ort"] || "").trim(),
    notify: String(m["Benachrichtigung an"] || "").trim(),
    notifyStorno: m["Mail bei Stornierung"] !== false
  };
}
function rows_() {
  const sh = book_().getSheetByName(RES);
  const v = sh.getDataRange().getValues();
  return { sh: sh, rows: v.slice(1).map((r, i) => ({ row: i + 2, code: String(r[1]), abend: nightId_(r[2]), plaetze: Number(r[3]) || 0, email: String(r[5]).toLowerCase(), status: String(r[7]), token: String(r[8]) })) };
}
const nightId_ = d => (d instanceof Date) ? Utilities.formatDate(d, "Europe/Berlin", "yyyy-MM-dd") : String(d).slice(0, 10);
function booked_(rows) {
  const b = {}; NIGHTS.forEach(n => b[n.id] = 0);
  rows.forEach(r => { if (r.status === "aktiv" && b[r.abend] !== undefined) b[r.abend] += r.plaetze; });
  return b;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function code_() {
  const a = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let c = "MW-";
  for (let i = 0; i < 5; i++) c += a.charAt(Math.floor(Math.random() * a.length));
  return c;
}
const escH_ = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const isPast_ = id => nightId_(new Date()) > id;

/* ---------- Status (GET) ---------- */
function doGet() {
  const s = settings_(), b = booked_(rows_().rows);
  return out_({
    ok: true, open: s.open, max: s.max, beginn: s.beginn, einlass: s.einlass, ort: s.ort,
    nights: NIGHTS.map(n => ({ id: n.id, label: n.label, total: s.cap[n.id], booked: b[n.id], free: Math.max(0, s.cap[n.id] - b[n.id]), past: isPast_(n.id) }))
  });
}

/* ---------- Reservieren / Stornieren (POST) ---------- */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "Ungültige Anfrage." }); }
  if (d.website) return out_({ ok: true, code: "MW-OK" }); // Honeypot gegen Spam-Bots
  if (d.action === "cancel") return cancel_(d);
  return reserve_(d);
}

function reserve_(d) {
  const s = settings_();
  if (!s.open) return out_({ ok: false, error: "Die Reservierung ist noch nicht geöffnet." });
  const name = String(d.name || "").trim().replace(/\s+/g, " ");
  const email = String(d.email || "").trim().toLowerCase();
  const hinweis = String(d.hinweis || "").trim().slice(0, 300);
  const plaetze = Math.floor(Number(d.plaetze));
  const night = NIGHTS.find(n => n.id === d.abend);
  if (!night) return out_({ ok: false, error: "Bitte wähle einen Abend." });
  if (isPast_(night.id)) return out_({ ok: false, error: "Dieser Abend ist schon vorbei." });
  if (name.length < 2 || name.length > 80) return out_({ ok: false, error: "Bitte gib deinen Namen an." });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 120) return out_({ ok: false, error: "Bitte prüfe deine E-Mail-Adresse." });
  if (!(plaetze >= 1 && plaetze <= s.max)) return out_({ ok: false, error: "Pro Reservierung sind 1 bis " + s.max + " Plätze möglich." });
  if (d.datenschutz !== true) return out_({ ok: false, error: "Bitte stimme dem Datenschutzhinweis zu." });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ ok: false, error: "Gerade ist viel los. Bitte versuch es gleich noch einmal." });
  let code, token, free;
  try {
    const { sh, rows } = rows_();
    const dup = rows.find(r => r.status === "aktiv" && r.abend === night.id && r.email === email);
    if (dup) return out_({ ok: false, error: "Mit dieser E-Mail-Adresse gibt es für diesen Abend schon eine Reservierung (" + dup.code + "). Für Änderungen antworte einfach auf die Bestätigungsmail." });
    free = s.cap[night.id] - booked_(rows)[night.id];
    if (plaetze > free) return out_({ ok: false, full: free <= 0, free: Math.max(0, free), error: free <= 0 ? "Dieser Abend ist leider ausgebucht." : "Für diesen Abend sind nur noch " + free + " Plätze frei." });
    const codes = new Set(rows.map(r => r.code));
    do { code = code_(); } while (codes.has(code));
    token = Utilities.getUuid().replace(/-/g, "");
    sh.appendRow([new Date(), code, "'" + night.id, plaetze, name, email, hinweis, "aktiv", token, ""]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  const storno = SEITE + "?storno=" + encodeURIComponent(code) + "&t=" + token;
  const wann = night.label + (s.beginn ? ", Beginn " + s.beginn : "") + (s.einlass ? " (Einlass " + s.einlass + ")" : "");
  let mailed = false;
  try {
    MailApp.sendEmail({
      to: email, name: "Mysterious Worlds", replyTo: s.notify || undefined,
      subject: "Deine Reservierung " + code + " – Mysterious Worlds",
      body: "Hallo " + name + ",\n\ndeine Plätze sind reserviert. Wir freuen uns auf dich!\n\nCode: " + code + "\nAbend: " + wann +
        "\nOrt: " + (s.ort || "folgt") + "\nPlätze: " + plaetze + "\n\nDer Eintritt ist frei. Am Einlass reicht dein Name oder der Code." +
        "\n\nDoch keine Zeit? Bitte storniere, damit jemand anderes den Platz bekommt:\n" + storno +
        "\n\nFragen? Antworte einfach auf diese Mail.\n\nDein Mysterious-Worlds-Team",
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:520px;color:#1d1830">' +
        '<img src="' + LOGO + '" width="520" height="165" alt="Mysterious Worlds – Musical, 17. &amp; 18. März 2027" style="display:block;width:100%;max-width:520px;height:auto;border:0;border-radius:8px;margin:0 0 18px;background:#1d1036;color:#ffd97a;font-size:22px;font-weight:bold;text-align:center">' +
        "<p>Hallo " + escH_(name) + ",</p><p>deine Plätze sind reserviert. Wir freuen uns auf dich!</p>" +
        '<table style="border-collapse:collapse;background:#f3f1f7;border-radius:6px;width:100%">' +
        '<tr><td style="padding:10px 14px;color:#625b78">Code</td><td style="padding:10px 14px;font-size:22px;font-weight:bold;letter-spacing:.08em;color:#5b2bb5">' + code + "</td></tr>" +
        '<tr><td style="padding:6px 14px;color:#625b78">Abend</td><td style="padding:6px 14px"><b>' + escH_(wann) + "</b></td></tr>" +
        '<tr><td style="padding:6px 14px;color:#625b78">Ort</td><td style="padding:6px 14px">' + escH_(s.ort || "folgt") + "</td></tr>" +
        '<tr><td style="padding:6px 14px 12px;color:#625b78">Plätze</td><td style="padding:6px 14px 12px">' + plaetze + "</td></tr></table>" +
        "<p>Der Eintritt ist frei. Am Einlass reicht dein Name oder der Code.</p>" +
        '<p>Doch keine Zeit? Bitte <a href="' + storno + '" style="color:#5b2bb5">storniere hier</a>, damit jemand anderes den Platz bekommt.</p>' +
        '<p style="color:#625b78;font-size:13px">Fragen? Antworte einfach auf diese Mail.</p></div>'
    });
    mailed = true;
  } catch (err) { console.warn("Bestätigungsmail fehlgeschlagen: " + err); }
  try {
    if (s.notify) MailApp.sendEmail({
      to: s.notify, name: "Reservierungen Mysterious Worlds",
      subject: "Neue Reservierung: " + plaetze + " × " + night.kurz + " – " + name,
      body: "Neue Reservierung " + code + "\n\nName: " + name + "\nE-Mail: " + email + "\nAbend: " + night.label + "\nPlätze: " + plaetze +
        (hinweis ? "\nHinweis: " + hinweis : "") + "\n\nNoch frei an diesem Abend: " + (free - plaetze) + " von " + s.cap[night.id] +
        (mailed ? "" : "\n\nACHTUNG: Die Bestätigungsmail an den Gast konnte nicht verschickt werden (Tageslimit?). Der Gast hat den Code auf der Seite gesehen.") +
        "\n\nAlle Reservierungen: " + book_().getUrl()
    });
  } catch (err) { console.warn("Benachrichtigung fehlgeschlagen: " + err); }

  return out_({ ok: true, code: code, abend: night.id, label: night.label, plaetze: plaetze, mailed: mailed, beginn: s.beginn, einlass: s.einlass, ort: s.ort, storno: storno });
}

function cancel_(d) {
  const code = String(d.code || "").trim().toUpperCase(), token = String(d.token || "").trim();
  if (!code || !token) return out_({ ok: false, error: "Der Storno-Link ist unvollständig." });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ ok: false, error: "Gerade ist viel los. Bitte versuch es gleich noch einmal." });
  let hit;
  try {
    const { sh, rows } = rows_();
    hit = rows.find(r => r.code === code && r.token === token);
    if (!hit) return out_({ ok: false, error: "Diese Reservierung wurde nicht gefunden." });
    if (hit.status !== "aktiv") return out_({ ok: true, already: true, code: code });
    sh.getRange(hit.row, 8).setValue("storniert");
    sh.getRange(hit.row, 10).setValue(new Date());
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  const s = settings_();
  try {
    if (s.notify && s.notifyStorno) {
      const n = NIGHTS.find(x => x.id === hit.abend);
      MailApp.sendEmail({ to: s.notify, name: "Reservierungen Mysterious Worlds",
        subject: "Stornierung: " + hit.plaetze + " × " + (n ? n.kurz : hit.abend) + " (" + code + ")",
        body: "Die Reservierung " + code + " (" + hit.plaetze + " Plätze, " + (n ? n.label : hit.abend) + ") wurde vom Gast storniert. Die Plätze sind wieder frei." });
    }
  } catch (err) { console.warn(err); }
  return out_({ ok: true, code: code, plaetze: hit.plaetze });
}
