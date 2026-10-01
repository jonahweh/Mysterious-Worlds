/**
 * Mysterious Worlds – Reservierungen (Google Apps Script) · Version 5 (Warteliste, Einlass, Teil-Storno, Freigeben am Einlass)
 *
 * Einrichtung am Computer ODER iPad (ca. 10 Minuten):
 * 1. Im Browser script.google.com öffnen (iPad: Safari, „Desktop-Website anfordern“)
 *    → Neues Projekt. Den gesamten Inhalt dieser Datei einfügen und speichern.
 *    (Alternativ am Computer: Google-Tabelle → Erweiterungen → Apps Script.)
 * 2. Oben die Funktion "setup" auswählen → Ausführen → Zugriff erlauben.
 *    Das Script legt die Tabelle "Mysterious Worlds – Reservierungen" mit den Blättern
 *    "Reservierungen", "Warteliste" und "Einstellungen" selbst an (Link steht im Ausführungsprotokoll).
 * 3. Bereitstellen → Neue Bereitstellung → Typ "Web-App".
 *    Ausführen als: Ich · Zugriff: Jeder → Bereitstellen.
 *    Die Web-App-URL (endet auf /exec) kommt in die Reservierungsseite.
 *
 * Update von einer älteren Version: Code ersetzen, speichern, "setup" EINMAL ausführen
 * (ergänzt neue Spalten, das Blatt "Warteliste" und die Einlass-PIN, löscht nichts),
 * dann Bereitstellen → Bereitstellungen verwalten → Bearbeiten → Neue Version → Bereitstellen.
 *
 * Im Alltag (geht auch in der Google-Tabellen-App):
 * - Reservierung öffnen/schließen, Plätze, Beginn, Einlass, Ort, Mail-Adresse, Warteliste
 *   und Einlass-PIN im Blatt "Einstellungen" ändern. Wirkt sofort, ohne neue Bereitstellung.
 * - Stornieren von Hand: in "Reservierungen" die Spalte Status auf "storniert" setzen
 *   (oder die Zahl bei Plätze verkleinern). Freie Plätze gehen automatisch an die Warteliste.
 */

const NIGHTS = [
  { id: "2027-03-17", label: "Mittwoch, 17. März 2027", kurz: "Mi 17.03." },
  { id: "2027-03-18", label: "Donnerstag, 18. März 2027", kurz: "Do 18.03." }
];
const SEITE = "https://jonahweh.github.io/Mysterious-Worlds/reservierung/";
const LOGO = "https://jonahweh.github.io/Mysterious-Worlds/mail-header.jpg"; // Kopfbild der Bestätigungsmail
const RES = "Reservierungen", SET = "Einstellungen", WL = "Warteliste";
const COLS = ["Eingang", "Code", "Abend", "Plätze", "Name", "E-Mail", "Hinweis", "Status", "Storno-Token", "Storniert am",
  "Angekommen", "Angekommen um", "Ursprünglich", "Quelle"];
const WCOLS = ["Eingang", "ID", "Abend", "Plätze", "Name", "E-Mail", "Status", "Code", "Token", "Geändert am", "Hinweis"];
const C = { eingang: 1, code: 2, abend: 3, plaetze: 4, name: 5, email: 6, hinweis: 7, status: 8, token: 9, storniert: 10, da: 11, daUm: 12, orig: 13, quelle: 14 };
const W = { eingang: 1, id: 2, abend: 3, plaetze: 4, name: 5, email: 6, status: 7, code: 8, token: 9, geaendert: 10 };

/* ---------- Tabelle finden (an eine Tabelle gebunden oder eigenständig) ---------- */
function book_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("Noch keine Tabelle: bitte zuerst die Funktion setup ausführen.");
  return SpreadsheetApp.openById(id);
}

/* ---------- Einrichtung & Update (darf beliebig oft ausgeführt werden) ---------- */
function setup() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (!ss) {
    const props = PropertiesService.getScriptProperties(), id = props.getProperty("SHEET_ID");
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create("Mysterious Worlds – Reservierungen"); props.setProperty("SHEET_ID", ss.getId()); }
  }
  // Reservierungen
  let r = ss.getSheetByName(RES);
  if (!r) { r = ss.getSheets()[0]; r.setName(RES); }
  ensureHeader_(r, COLS);
  r.setFrozenRows(1);
  r.getRange("A:A").setNumberFormat("dd.mm.yyyy hh:mm");
  r.getRange("C2:C").setNumberFormat("@");
  r.getRange("L:L").setNumberFormat("hh:mm");
  r.getRange("H2:H").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["aktiv", "storniert"], true).build());
  // Warteliste
  let w = ss.getSheetByName(WL);
  if (!w) w = ss.insertSheet(WL);
  ensureHeader_(w, WCOLS);
  w.setFrozenRows(1);
  w.getRange("A:A").setNumberFormat("dd.mm.yyyy hh:mm");
  w.getRange("C2:C").setNumberFormat("@");
  w.getRange("G2:G").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["wartend", "nachgerückt", "ausgetragen", "erledigt"], true).build());
  // Einstellungen
  let s = ss.getSheetByName(SET);
  if (!s) s = ss.insertSheet(SET);
  const me = Session.getEffectiveUser().getEmail();
  ensureSetting_(s, "Reservierung offen", false, "Häkchen setzen, sobald Gäste reservieren dürfen", true);
  ensureSetting_(s, "Plätze " + NIGHTS[0].kurz, 300, "");
  ensureSetting_(s, "Plätze " + NIGHTS[1].kurz, 300, "");
  ensureSetting_(s, "Max. Plätze pro Reservierung", 6, "");
  ensureSetting_(s, "Beginn", "", "z. B. 19:00 Uhr – leer lassen = „folgt“");
  ensureSetting_(s, "Einlass", "", "z. B. ab 18:30 Uhr");
  ensureSetting_(s, "Ort", "", "z. B. Aula, mit Adresse");
  ensureSetting_(s, "Benachrichtigung an", me, "An diese Adresse geht bei jeder Reservierung eine Mail");
  ensureSetting_(s, "Mail bei Stornierung", true, "", true);
  ensureSetting_(s, "Warteliste aktiv", true, "Ist ein Abend voll, können sich Gäste eintragen und rücken automatisch nach", true);
  ensureSetting_(s, "Einlass-PIN", String(100000 + Math.floor(Math.random() * 900000)), "PIN für die Einlass-Seite …/einlass/ – nur an das Einlass-Team geben");
  NIGHTS.forEach(n => ensureSetting_(s, "Reserviert " + n.kurz,
    `=SUMIFS(${RES}!D:D,${RES}!C:C,"${n.id}",${RES}!H:H,"aktiv")`, "wird automatisch berechnet"));
  ensureSetting_(s, "Auf der Warteliste (Plätze)", `=SUMIFS(${WL}!D:D,${WL}!G:G,"wartend")`, "wird automatisch berechnet");
  s.getRange("A:A").setFontWeight("bold");
  s.getRange("C:C").setFontColor("#625b78");
  s.setColumnWidth(1, 230); s.setColumnWidth(2, 260); s.setColumnWidth(3, 420);
  try { ss.setActiveSheet(s); } catch (e) {}
  console.log("Fertig! Deine Tabelle: " + ss.getUrl());
  console.log("Einlass-PIN: " + settings_().pin + " (steht auch im Blatt „Einstellungen“)");
}
function ensureHeader_(sh, cols) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const have = sh.getLastRow() ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  cols.forEach((c, i) => { if (!have[i]) sh.getRange(1, i + 1).setValue(c); });
  sh.getRange(1, 1, 1, cols.length).setFontWeight("bold").setBackground("#ece6fb");
}
function ensureSetting_(s, label, value, note, checkbox) {
  const vals = s.getLastRow() ? s.getRange(1, 1, s.getLastRow(), 1).getValues() : [];
  if (vals.some(r => String(r[0]).trim() === label)) return;
  const row = s.getLastRow() + 1;
  s.getRange(row, 1, 1, 3).setValues([[label, checkbox ? "" : value, note]]);
  if (checkbox) { s.getRange(row, 2).insertCheckboxes(); s.getRange(row, 2).setValue(value); }
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
    notifyStorno: m["Mail bei Stornierung"] !== false,
    warteliste: m["Warteliste aktiv"] !== false,
    pin: String(m["Einlass-PIN"] === undefined ? "" : m["Einlass-PIN"]).trim()
  };
}
const nightId_ = d => (d instanceof Date) ? Utilities.formatDate(d, "Europe/Berlin", "yyyy-MM-dd") : String(d).slice(0, 10);
const today_ = () => nightId_(new Date());
const isPast_ = id => today_() > id;
const nightOf_ = id => NIGHTS.find(n => n.id === id);
const norm_ = s => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
function rows_() {
  const sh = book_().getSheetByName(RES);
  const v = sh.getDataRange().getValues();
  return { sh: sh, rows: v.slice(1).map((r, i) => ({
    row: i + 2, code: String(r[1]), abend: nightId_(r[2]), plaetze: Number(r[3]) || 0, name: String(r[4]),
    email: String(r[5]).toLowerCase().trim(), hinweis: String(r[6] || ""), status: String(r[7]), token: String(r[8]),
    da: Number(r[10]) || 0, orig: Number(r[12]) || 0, quelle: String(r[13] || "")
  })).filter(r => r.code) };
}
function wl_() {
  const sh = book_().getSheetByName(WL);
  const v = sh ? sh.getDataRange().getValues() : [];
  return { sh: sh, rows: v.slice(1).map((r, i) => ({
    row: i + 2, id: String(r[1]), abend: String(r[2]).slice(0, 10) === "beide" ? "beide" : nightId_(r[2]),
    plaetze: Number(r[3]) || 0, name: String(r[4]), email: String(r[5]).toLowerCase().trim(), status: String(r[6]), token: String(r[8]),
    hinweis: String(r[10] || "")
  })).filter(r => r.id) };
}
function booked_(rows) {
  const b = {}; NIGHTS.forEach(n => b[n.id] = 0);
  rows.forEach(r => { if (r.status === "aktiv" && b[r.abend] !== undefined) b[r.abend] += r.plaetze; });
  return b;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function rand_(prefix, n) {
  const a = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let c = prefix;
  for (let i = 0; i < n; i++) c += a.charAt(Math.floor(Math.random() * a.length));
  return c;
}
const token_ = () => Utilities.getUuid().replace(/-/g, "");
const escH_ = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const validEmail_ = e => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e) && e.length <= 120;
function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: "Gerade ist viel los. Bitte versuch es gleich noch einmal." };
  try { return fn(); } finally { lock.releaseLock(); }
}
// Ein Abend nimmt keine Nachrücker mehr auf, sobald der Aufführungstag da ist
const waitlistOpenFor_ = id => today_() < id;

/* ---------- Mails ---------- */
function send_(o) {
  try { MailApp.sendEmail(o); return true; } catch (err) { console.warn("Mail fehlgeschlagen: " + err); return false; }
}
function confirmMail_(s, b) {
  const night = nightOf_(b.abend);
  const storno = SEITE + "?storno=" + encodeURIComponent(b.code) + "&t=" + b.token;
  const wann = night.label + (s.beginn ? ", Beginn " + s.beginn : "") + (s.einlass ? " (Einlass " + s.einlass + ")" : "");
  const intro = b.nachgerueckt
    ? "gute Nachricht: Es sind Plätze frei geworden, und du bist von der Warteliste nachgerückt. Deine Plätze sind fest reserviert, du musst nichts weiter tun."
    : "deine Plätze sind reserviert. Wir freuen uns auf dich!";
  return send_({
    to: b.email, name: "Mysterious Worlds", replyTo: s.notify || undefined,
    subject: (b.nachgerueckt ? "Du bist nachgerückt! Reservierung " : "Deine Reservierung ") + b.code + " – Mysterious Worlds",
    body: "Hallo " + b.name + ",\n\n" + intro + "\n\nCode: " + b.code + "\nAbend: " + wann +
      "\nOrt: " + (s.ort || "folgt") + "\nPlätze: " + b.plaetze + "\n\nDer Eintritt ist frei. Am Einlass reicht dein Name oder der Code." +
      "\n\nDoch keine Zeit, oder kommen weniger Personen? Bitte storniere (auch einzelne Plätze), damit andere nachrücken können:\n" + storno +
      "\n\nFragen? Antworte einfach auf diese Mail.\n\nDein Mysterious-Worlds-Team",
    htmlBody: '<div style="font-family:Arial,sans-serif;max-width:520px;color:#1d1830">' +
      '<img src="' + LOGO + '" width="520" height="165" alt="Mysterious Worlds – Musical, 17. &amp; 18. März 2027" style="display:block;width:100%;max-width:520px;height:auto;border:0;border-radius:8px;margin:0 0 18px;background:#1d1036;color:#ffd97a;font-size:22px;font-weight:bold;text-align:center">' +
      "<p>Hallo " + escH_(b.name) + ",</p><p>" + intro + "</p>" +
      '<table style="border-collapse:collapse;background:#f3f1f7;border-radius:6px;width:100%">' +
      '<tr><td style="padding:10px 14px;color:#625b78">Code</td><td style="padding:10px 14px;font-size:22px;font-weight:bold;letter-spacing:.08em;color:#5b2bb5">' + b.code + "</td></tr>" +
      '<tr><td style="padding:6px 14px;color:#625b78">Abend</td><td style="padding:6px 14px"><b>' + escH_(wann) + "</b></td></tr>" +
      '<tr><td style="padding:6px 14px;color:#625b78">Ort</td><td style="padding:6px 14px">' + escH_(s.ort || "folgt") + "</td></tr>" +
      '<tr><td style="padding:6px 14px 12px;color:#625b78">Plätze</td><td style="padding:6px 14px 12px">' + b.plaetze + "</td></tr></table>" +
      "<p>Der Eintritt ist frei. Am Einlass reicht dein Name oder der Code.</p>" +
      '<p>Doch keine Zeit, oder kommen weniger Personen? Bitte <a href="' + storno + '" style="color:#5b2bb5">storniere hier</a> – auch einzelne Plätze –, damit andere nachrücken können.</p>' +
      '<p style="color:#625b78;font-size:13px">Fragen? Antworte einfach auf diese Mail.</p></div>'
  });
}
function ownerMail_(s, subject, body) {
  if (s.notify) send_({ to: s.notify, name: "Reservierungen Mysterious Worlds", subject: subject, body: body + "\n\nAlle Reservierungen: " + book_().getUrl() });
}

/* ---------- Warteliste abarbeiten (nur innerhalb des Locks aufrufen) ---------- */
function processWaitlist_(s) {
  const done = [];
  if (!s.open || !s.warteliste) return done;
  const Wl = wl_();
  const waiting = Wl.rows.filter(w => w.status === "wartend");
  if (!waiting.length) return done;
  const R = rows_(), b = booked_(R.rows), free = {};
  NIGHTS.forEach(n => free[n.id] = waitlistOpenFor_(n.id) ? s.cap[n.id] - b[n.id] : 0);
  if (!NIGHTS.some(n => free[n.id] > 0)) return done;
  const active = new Set(R.rows.filter(r => r.status === "aktiv" && r.quelle !== "Abendkasse").map(r => r.email));
  const codes = new Set(R.rows.map(r => r.code));
  for (const w of waiting) {
    if (active.has(w.email)) { Wl.sh.getRange(w.row, W.status).setValue("erledigt"); Wl.sh.getRange(w.row, W.geaendert).setValue(new Date()); continue; }
    const cands = w.abend === "beide" ? NIGHTS.map(n => n.id) : [w.abend];
    const night = cands.find(id => free[id] >= w.plaetze);
    if (!night) continue;
    let code; do { code = rand_("MW-", 5); } while (codes.has(code));
    const token = token_();
    R.sh.appendRow([new Date(), code, "'" + night, w.plaetze, w.name, w.email, w.hinweis, "aktiv", token, "", 0, "", "", "Warteliste"]);
    Wl.sh.getRange(w.row, W.status, 1, 4).setValues([["nachgerückt", code, w.token, new Date()]]);
    free[night] -= w.plaetze; active.add(w.email); codes.add(code);
    done.push({ name: w.name, email: w.email, code: code, token: token, abend: night, plaetze: w.plaetze, nachgerueckt: true, free: free[night] });
  }
  if (done.length) SpreadsheetApp.flush();
  return done;
}
function announce_(s, moved) {
  moved.forEach(m => {
    const ok = confirmMail_(s, m);
    ownerMail_(s, "Nachgerückt von der Warteliste: " + m.plaetze + " × " + nightOf_(m.abend).kurz + " – " + m.name,
      "Von der Warteliste nachgerückt: " + m.code + "\n\nName: " + m.name + "\nE-Mail: " + m.email + "\nAbend: " + nightOf_(m.abend).label +
      "\nPlätze: " + m.plaetze + "\nNoch frei an diesem Abend: " + m.free + (ok ? "" : "\n\nACHTUNG: Die Bestätigungsmail an den Gast konnte nicht verschickt werden (Tageslimit?). Bitte den Gast selbst informieren."));
  });
}

/* ---------- Status (GET) ---------- */
function doGet() {
  const s = settings_();
  let moved = [];
  if (s.open && s.warteliste) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) { try { moved = processWaitlist_(s); } finally { lock.releaseLock(); } }
  }
  announce_(s, moved);
  const b = booked_(rows_().rows), wait = {};
  NIGHTS.forEach(n => wait[n.id] = 0);
  wl_().rows.filter(w => w.status === "wartend").forEach(w => (w.abend === "beide" ? NIGHTS.map(n => n.id) : [w.abend]).forEach(id => { if (wait[id] !== undefined) wait[id]++; }));
  return out_({
    ok: true, open: s.open, max: s.max, beginn: s.beginn, einlass: s.einlass, ort: s.ort, warteliste: s.warteliste,
    nights: NIGHTS.map(n => ({ id: n.id, label: n.label, total: s.cap[n.id], booked: b[n.id], free: Math.max(0, s.cap[n.id] - b[n.id]),
      past: isPast_(n.id), waiting: wait[n.id], waitlistOpen: waitlistOpenFor_(n.id) }))
  });
}

/* ---------- POST: Reservieren, Warteliste, Stornieren, Einlass ---------- */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "Ungültige Anfrage." }); }
  if (d.website) return out_({ ok: true, code: "MW-OK", id: "W-OK" }); // Honeypot gegen Spam-Bots
  const a = d.action || "reserve";
  const fn = { reserve: reserve_, waitlist: waitlist_, cancel: cancel_, lookup: lookup_, leave: leave_,
    admin_list: adminList_, checkin: checkin_, walkin: walkin_, release: release_, release_all: releaseAll_, admit: admit_ }[a];
  if (!fn) return out_({ ok: false, error: "Unbekannte Aktion." });
  return out_(fn(d));
}

function readGuest_(d, s) {
  const g = {
    name: String(d.name || "").trim().replace(/\s+/g, " "),
    email: String(d.email || "").trim().toLowerCase(),
    hinweis: String(d.hinweis || "").trim().slice(0, 300),
    plaetze: Math.floor(Number(d.plaetze))
  };
  if (g.name.length < 2 || g.name.length > 80) return { error: "Bitte gib deinen Namen an." };
  if (!validEmail_(g.email)) return { error: "Bitte prüfe deine E-Mail-Adresse." };
  if (!(g.plaetze >= 1 && g.plaetze <= s.max)) return { error: "Pro Reservierung sind 1 bis " + s.max + " Plätze möglich." };
  if (d.datenschutz !== true) return { error: "Bitte stimme dem Datenschutzhinweis zu." };
  return g;
}
const oneNightMsg_ = r => "Mit dieser E-Mail-Adresse gibt es schon eine Reservierung für " + nightOf_(r.abend).label + " (" + r.code +
  "). Damit möglichst viele die Aufführung sehen können, ist pro E-Mail-Adresse nur ein Abend möglich. Für Änderungen antworte einfach auf deine Bestätigungsmail.";

function reserve_(d) {
  const s = settings_();
  if (!s.open) return { ok: false, error: "Die Reservierung ist noch nicht geöffnet." };
  const night = nightOf_(d.abend);
  if (!night) return { ok: false, error: "Bitte wähle einen Abend." };
  if (isPast_(night.id)) return { ok: false, error: "Dieser Abend ist schon vorbei." };
  const g = readGuest_(d, s);
  if (g.error) return { ok: false, error: g.error };
  let moved = [], booking, sameName;
  const res = locked_(() => {
    moved = processWaitlist_(s);
    const R = rows_();
    const mine = R.rows.find(r => r.status === "aktiv" && r.email === g.email && r.quelle !== "Abendkasse");
    if (mine) return { ok: false, oneNight: true, error: oneNightMsg_(mine) };
    const free = s.cap[night.id] - booked_(R.rows)[night.id];
    if (g.plaetze > free) return { ok: false, full: free <= 0, free: Math.max(0, free), waitlist: s.warteliste && waitlistOpenFor_(night.id),
      error: free <= 0 ? "Dieser Abend ist leider ausgebucht." : "Für diesen Abend sind nur noch " + free + " Plätze frei." };
    sameName = R.rows.find(r => r.status === "aktiv" && r.quelle !== "Abendkasse" && norm_(r.name) === norm_(g.name));
    const codes = new Set(R.rows.map(r => r.code));
    let code; do { code = rand_("MW-", 5); } while (codes.has(code));
    booking = { name: g.name, email: g.email, code: code, token: token_(), abend: night.id, plaetze: g.plaetze, free: free - g.plaetze };
    R.sh.appendRow([new Date(), code, "'" + night.id, g.plaetze, g.name, g.email, g.hinweis, "aktiv", booking.token, "", 0, "", "", "Website"]);
    // Falls dieselbe Adresse noch auf der Warteliste steht: erledigt
    const Wl = wl_();
    Wl.rows.filter(w => w.status === "wartend" && w.email === g.email).forEach(w => { Wl.sh.getRange(w.row, W.status).setValue("erledigt"); Wl.sh.getRange(w.row, W.geaendert).setValue(new Date()); });
    SpreadsheetApp.flush();
    return { ok: true };
  });
  announce_(s, moved);
  if (!res.ok) return res;
  const mailed = confirmMail_(s, booking);
  ownerMail_(s, "Neue Reservierung: " + g.plaetze + " × " + night.kurz + " – " + g.name,
    "Neue Reservierung " + booking.code + "\n\nName: " + g.name + "\nE-Mail: " + g.email + "\nAbend: " + night.label + "\nPlätze: " + g.plaetze +
    (g.hinweis ? "\nHinweis: " + g.hinweis : "") + "\n\nNoch frei an diesem Abend: " + booking.free + " von " + s.cap[night.id] +
    (sameName ? "\n\nHINWEIS: Unter dem gleichen Namen gibt es schon eine Reservierung (" + sameName.code + ", " + nightOf_(sameName.abend).kurz + ", " + sameName.plaetze + " Plätze, " + sameName.email + "). Evtl. doppelt mit anderer Mail-Adresse?" : "") +
    (mailed ? "" : "\n\nACHTUNG: Die Bestätigungsmail an den Gast konnte nicht verschickt werden (Tageslimit?). Der Gast hat den Code auf der Seite gesehen."));
  return { ok: true, code: booking.code, abend: night.id, label: night.label, plaetze: g.plaetze, mailed: mailed,
    beginn: s.beginn, einlass: s.einlass, ort: s.ort, storno: SEITE + "?storno=" + encodeURIComponent(booking.code) + "&t=" + booking.token };
}

function waitlist_(d) {
  const s = settings_();
  if (!s.open || !s.warteliste) return { ok: false, error: "Die Warteliste ist gerade nicht geöffnet." };
  const abend = d.abend === "beide" ? "beide" : (nightOf_(d.abend) ? d.abend : null);
  if (!abend) return { ok: false, error: "Bitte wähle einen Abend." };
  const ids = abend === "beide" ? NIGHTS.map(n => n.id) : [abend];
  if (!ids.some(waitlistOpenFor_)) return { ok: false, error: "Für diesen Abend ist die Warteliste geschlossen." };
  const g = readGuest_(d, s);
  if (g.error) return { ok: false, error: g.error };
  let entry, moved = [];
  const res = locked_(() => {
    moved = processWaitlist_(s);
    const R = rows_();
    const mine = R.rows.find(r => r.status === "aktiv" && r.email === g.email && r.quelle !== "Abendkasse");
    if (mine) return { ok: false, oneNight: true, error: oneNightMsg_(mine) };
    const b = booked_(R.rows);
    const fit = ids.find(id => waitlistOpenFor_(id) && s.cap[id] - b[id] >= g.plaetze);
    if (fit) return { ok: false, freeAgain: true, abend: fit, error: "Gute Nachricht: Am " + nightOf_(fit).label + " sind gerade wieder genug Plätze frei. Du kannst direkt reservieren." };
    const Wl = wl_();
    if (Wl.rows.some(w => w.status === "wartend" && w.email === g.email)) return { ok: false, error: "Mit dieser E-Mail-Adresse stehst du schon auf der Warteliste." };
    const idset = new Set(Wl.rows.map(w => w.id));
    let id; do { id = rand_("W-", 5); } while (idset.has(id));
    entry = { id: id, token: token_(), abend: abend, plaetze: g.plaetze, name: g.name, email: g.email };
    entry.position = Wl.rows.filter(w => w.status === "wartend" && (w.abend === "beide" || abend === "beide" || w.abend === abend)).length + 1;
    Wl.sh.appendRow([new Date(), id, abend === "beide" ? "beide" : "'" + abend, g.plaetze, g.name, g.email, "wartend", "", entry.token, "", g.hinweis]);
    SpreadsheetApp.flush();
    return { ok: true };
  });
  announce_(s, moved);
  if (!res.ok) return res;
  const wann = abend === "beide" ? "Mittwoch oder Donnerstag (was zuerst frei wird)" : nightOf_(abend).label;
  const leave = SEITE + "?warteliste=" + encodeURIComponent(entry.id) + "&t=" + entry.token;
  const mailed = send_({
    to: g.email, name: "Mysterious Worlds", replyTo: s.notify || undefined,
    subject: "Du stehst auf der Warteliste – Mysterious Worlds",
    body: "Hallo " + g.name + ",\n\ndu stehst auf der Warteliste für " + wann + " (" + g.plaetze + " " + (g.plaetze === 1 ? "Platz" : "Plätze") + ", Position " + entry.position + ").\n\n" +
      "Sobald genug Plätze frei werden, rückst du automatisch nach und bekommst eine Bestätigung mit deinem Reservierungscode per Mail. Du musst nichts weiter tun.\n\n" +
      "Kein Interesse mehr? Hier kannst du dich austragen:\n" + leave + "\n\nDein Mysterious-Worlds-Team",
    htmlBody: '<div style="font-family:Arial,sans-serif;max-width:520px;color:#1d1830">' +
      '<img src="' + LOGO + '" width="520" height="165" alt="Mysterious Worlds" style="display:block;width:100%;max-width:520px;height:auto;border:0;border-radius:8px;margin:0 0 18px;background:#1d1036;color:#ffd97a;font-size:22px;font-weight:bold;text-align:center">' +
      "<p>Hallo " + escH_(g.name) + ",</p><p>du stehst auf der <b>Warteliste</b> für <b>" + escH_(wann) + "</b> (" + g.plaetze + " " + (g.plaetze === 1 ? "Platz" : "Plätze") + ", Position " + entry.position + ").</p>" +
      "<p>Sobald genug Plätze frei werden, rückst du <b>automatisch nach</b> und bekommst eine Bestätigung mit deinem Reservierungscode per Mail. Du musst nichts weiter tun.</p>" +
      '<p>Kein Interesse mehr? <a href="' + leave + '" style="color:#5b2bb5">Hier austragen</a>.</p></div>'
  });
  return { ok: true, id: entry.id, position: entry.position, abend: abend, label: wann, plaetze: g.plaetze, mailed: mailed };
}

function findBooking_(R, d) {
  const code = String(d.code || "").trim().toUpperCase(), token = String(d.token || "").trim();
  if (!code || !token) return null;
  return R.rows.find(r => r.code === code && r.token === token) || null;
}
function lookup_(d) {
  const hit = findBooking_(rows_(), d);
  if (!hit) return { ok: false, error: "Diese Reservierung wurde nicht gefunden." };
  const n = nightOf_(hit.abend);
  return { ok: true, code: hit.code, name: hit.name, abend: hit.abend, label: n ? n.label : hit.abend, plaetze: hit.plaetze, status: hit.status, past: isPast_(hit.abend) };
}
function cancel_(d) {
  const s = settings_();
  let hit, n, rest, moved = [];
  const res = locked_(() => {
    const R = rows_();
    hit = findBooking_(R, d);
    if (!hit) return { ok: false, error: "Diese Reservierung wurde nicht gefunden." };
    if (hit.status !== "aktiv") return { ok: true, already: true, code: hit.code };
    n = d.anzahl === undefined ? hit.plaetze : Math.floor(Number(d.anzahl));
    if (!(n >= 1 && n <= hit.plaetze)) return { ok: false, error: "Bitte wähle 1 bis " + hit.plaetze + " Plätze." };
    rest = hit.plaetze - n;
    if (rest === 0) {
      R.sh.getRange(hit.row, C.status).setValue("storniert");
      R.sh.getRange(hit.row, C.storniert).setValue(new Date());
    } else {
      R.sh.getRange(hit.row, C.plaetze).setValue(rest);
      if (!hit.orig) R.sh.getRange(hit.row, C.orig).setValue(hit.plaetze);
      if (hit.da > rest) R.sh.getRange(hit.row, C.da).setValue(rest);
    }
    SpreadsheetApp.flush();
    moved = processWaitlist_(s);
    return { ok: true };
  });
  announce_(s, moved);
  if (!res.ok || res.already) return res;
  const night = nightOf_(hit.abend), label = night ? night.label : hit.abend;
  if (s.notifyStorno) ownerMail_(s, (rest ? "Teil-Stornierung: " + n + " von " + hit.plaetze : "Stornierung: " + n) + " × " + (night ? night.kurz : hit.abend) + " (" + hit.code + ")",
    (rest ? "Bei der Reservierung " + hit.code + " (" + hit.name + ", " + label + ") wurden " + n + " von " + hit.plaetze + " Plätzen storniert. Es bleiben " + rest + " Plätze."
      : "Die Reservierung " + hit.code + " (" + hit.name + ", " + n + " Plätze, " + label + ") wurde vom Gast storniert.") +
    (moved.length ? "\n\nDie freien Plätze sind an die Warteliste gegangen (" + moved.length + " Nachrücker)." : "\n\nDie Plätze sind wieder frei."));
  return { ok: true, code: hit.code, cancelled: n, remaining: rest, label: label };
}
function leave_(d) {
  return locked_(() => {
    const Wl = wl_();
    const w = Wl.rows.find(x => x.id === String(d.id || "").trim().toUpperCase() && x.token === String(d.token || "").trim());
    if (!w) return { ok: false, error: "Dieser Wartelisten-Eintrag wurde nicht gefunden." };
    if (w.status === "nachgerückt") return { ok: false, error: "Du bist bereits nachgerückt und hast eine feste Reservierung. Zum Stornieren nutze den Link in deiner Bestätigungsmail." };
    if (w.status !== "wartend") return { ok: true, already: true };
    Wl.sh.getRange(w.row, W.status).setValue("ausgetragen");
    Wl.sh.getRange(w.row, W.geaendert).setValue(new Date());
    return { ok: true };
  });
}

/* ---------- Einlass (mit PIN) ---------- */
function pinOk_(s, pin) {
  if (!s.pin) return { ok: false, error: "Im Blatt „Einstellungen“ ist keine Einlass-PIN gesetzt." };
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("pinfails") || 0);
  if (fails >= 20) return { ok: false, locked: true, error: "Zu viele falsche PINs. Bitte 15 Minuten warten." };
  if (String(pin || "").trim() !== s.pin) {
    cache.put("pinfails", String(fails + 1), 900);
    return { ok: false, badPin: true, error: "Falsche PIN." };
  }
  return { ok: true };
}
function adminData_(s, abend) {
  const R = rows_();
  const list = R.rows.filter(r => r.abend === abend && r.status === "aktiv");
  const guests = list.filter(r => r.quelle !== "Abendkasse");
  const walk = list.filter(r => r.quelle === "Abendkasse").reduce((a, r) => a + r.plaetze, 0);
  const reserved = guests.reduce((a, r) => a + r.plaetze, 0), arrived = guests.reduce((a, r) => a + Math.min(r.da, r.plaetze), 0);
  const waiting = wl_().rows.filter(w => w.status === "wartend" && (w.abend === abend || w.abend === "beide"))
    .map((w, i) => ({ id: w.id, nr: i + 1, name: w.name, plaetze: w.plaetze, hinweis: w.hinweis, beide: w.abend === "beide" }));
  return {
    ok: true, abend: abend, label: nightOf_(abend).label, cap: s.cap[abend], reserved: reserved, arrived: arrived, walkins: walk,
    guests: guests.map(r => ({ code: r.code, name: r.name, plaetze: r.plaetze, da: Math.min(r.da, r.plaetze), orig: r.orig, hinweis: r.hinweis, quelle: r.quelle }))
      .sort((a, b) => a.name.localeCompare(b.name, "de")),
    waiting: waiting
  };
}
// Am Einlass: nicht abgeholte Plätze einer Reservierung freigeben (Plätze = Angekommene; niemand da = storniert)
function releaseRow_(R, r) {
  const da = Math.min(r.da, r.plaetze);
  if (da >= r.plaetze) return 0;
  if (da === 0) {
    R.sh.getRange(r.row, C.status).setValue("storniert");
    R.sh.getRange(r.row, C.storniert).setValue(new Date());
  } else {
    R.sh.getRange(r.row, C.plaetze).setValue(da);
    if (!r.orig) R.sh.getRange(r.row, C.orig).setValue(r.plaetze);
  }
  return r.plaetze - da;
}
function release_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  let freed = 0;
  const res = locked_(() => {
    const R = rows_();
    const r = R.rows.find(x => x.code === String(d.code || "").trim().toUpperCase() && x.status === "aktiv" && x.quelle !== "Abendkasse");
    if (!r) return { ok: false, error: "Reservierung nicht gefunden (oder schon storniert)." };
    freed = releaseRow_(R, r);
    SpreadsheetApp.flush();
    return { ok: true, abend: r.abend };
  });
  return res.ok ? Object.assign(adminData_(s, res.abend), { freed: freed }) : res;
}
function releaseAll_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  const night = nightOf_(d.abend);
  if (!night) return { ok: false, error: "Unbekannter Abend." };
  let freed = 0;
  const res = locked_(() => {
    const R = rows_();
    R.rows.filter(r => r.abend === night.id && r.status === "aktiv" && r.quelle !== "Abendkasse").forEach(r => { freed += releaseRow_(R, r); });
    SpreadsheetApp.flush();
    return { ok: true };
  });
  return res.ok ? Object.assign(adminData_(s, night.id), { freed: freed }) : res;
}
// Am Einlass: wartende Person direkt einlassen (wird als Reservierung mit Quelle „Warteliste“ angelegt, ohne Mail)
function admit_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  const night = nightOf_(d.abend);
  if (!night) return { ok: false, error: "Unbekannter Abend." };
  const res = locked_(() => {
    const Wl = wl_();
    const w = Wl.rows.find(x => x.id === String(d.id || "").trim().toUpperCase());
    if (!w || w.status !== "wartend") return { ok: false, error: "Dieser Eintrag steht nicht mehr auf der Warteliste." };
    if (w.abend !== "beide" && w.abend !== night.id) return { ok: false, error: "Dieser Eintrag gilt für einen anderen Abend." };
    const R = rows_();
    const mine = R.rows.find(r => r.status === "aktiv" && r.email === w.email && r.quelle !== "Abendkasse");
    if (mine) {
      Wl.sh.getRange(w.row, W.status).setValue("erledigt"); Wl.sh.getRange(w.row, W.geaendert).setValue(new Date());
      return { ok: false, error: w.name + " hat schon eine Reservierung (" + mine.code + ", " + nightOf_(mine.abend).kurz + "). Bitte dort abhaken." };
    }
    const n = Math.max(1, Math.min(w.plaetze, Math.floor(Number(d.plaetze) || w.plaetze)));
    const codes = new Set(R.rows.map(r => r.code));
    let code; do { code = rand_("MW-", 5); } while (codes.has(code));
    R.sh.appendRow([new Date(), code, "'" + night.id, n, w.name, w.email, w.hinweis, "aktiv", token_(), "", n, new Date(), "", "Warteliste"]);
    Wl.sh.getRange(w.row, W.status, 1, 4).setValues([["nachgerückt", code, w.token, new Date()]]);
    SpreadsheetApp.flush();
    return { ok: true };
  });
  return res.ok ? adminData_(s, night.id) : res;
}
function adminList_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  const abend = nightOf_(d.abend) ? d.abend : NIGHTS[0].id;
  return Object.assign(adminData_(s, abend), { nights: NIGHTS.map(n => ({ id: n.id, label: n.label, kurz: n.kurz })) });
}
function checkin_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  const res = locked_(() => {
    const R = rows_();
    const r = R.rows.find(x => x.code === String(d.code || "").trim().toUpperCase() && x.status === "aktiv");
    if (!r) return { ok: false, error: "Reservierung nicht gefunden (oder storniert)." };
    const n = Math.max(0, Math.min(r.plaetze, Math.floor(Number(d.da) || 0)));
    R.sh.getRange(r.row, C.da, 1, 2).setValues([[n, n ? new Date() : ""]]);
    return { ok: true, abend: r.abend };
  });
  return res.ok ? adminData_(s, res.abend) : res;
}
function walkin_(d) {
  const s = settings_(), p = pinOk_(s, d.pin);
  if (!p.ok) return p;
  const night = nightOf_(d.abend);
  if (!night) return { ok: false, error: "Unbekannter Abend." };
  const res = locked_(() => {
    const R = rows_();
    const r = R.rows.find(x => x.abend === night.id && x.quelle === "Abendkasse" && x.status === "aktiv");
    const delta = Math.floor(Number(d.delta) || 0);
    if (!r) {
      if (delta > 0) R.sh.appendRow([new Date(), "AK-" + night.id.slice(5).replace("-", ""), "'" + night.id, delta, "Abendkasse (ohne Reservierung)", "", "", "aktiv", "", "", delta, new Date(), "", "Abendkasse"]);
    } else {
      const n = Math.max(0, r.plaetze + delta);
      R.sh.getRange(r.row, C.plaetze).setValue(n);
      R.sh.getRange(r.row, C.da, 1, 2).setValues([[n, new Date()]]);
    }
    return { ok: true };
  });
  return res.ok ? adminData_(s, night.id) : res;
}
