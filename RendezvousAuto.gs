/**
 * PPREN · Rendez-vousen am Pak plangen (eenzel oder automatesch verdeelt)
 * ----------------------------------------------------------------------
 * Router-Zeil an PPREN.gs (doPost), nieft "rendezvousPlangen":
 *   } else if (data.typ === "rendezvousenPlangen") { return jsonResponse(rendezvousenPlangen(data));
 *
 * data = { token, erstalltVum, mailSchecken: true|false,
 *          mailText: "Hallo {virnumm}, … {terminer} … {ziler} … {meilensteng}" (optional),
 *          rendezvousen: [{ schueler, klasse, rvTyp, datum: "dd.MM.yyyy", zaeit: "HH:mm", dauer: 15, notiz }] }
 *
 * - schreift all Rendez-vousen an d'Tab "Rendezvousen" (Spalt K = Dauer an Minutten)
 * - Datum/Zäit ginn als Text gespäichert (soss mécht Sheets en Datum draus)
 * - schéckt all Schüler direkt eng Mail mat sengen Terminer + .ics-Kalennerdatei (mat Erënnerung 1 Dag virdrun)
 * - d'Erënnerung 2 Deeg virdrun (sendeRendezvousErennerungen) leeft weider wéi bis elo
 */

function rendezvousenPlangen(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Rendez-vousen plangen." };
  }
  const liste = (data.rendezvousen || []).filter((r) => r && r.schueler && r.datum);
  if (liste.length === 0) return { ok: false, error: "Keng Rendez-vousen iwwerginn." };

  const sheet = getRendezvousenSheet();
  if (String(sheet.getRange(1, 11).getValue()) !== "Dauer") sheet.getRange(1, 11).setValue("Dauer");

  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const vun = data.erstalltVum || session.numm || "";
  const zeilen = liste.map((r) => {
    r.id = Utilities.getUuid();
    return [r.id, r.schueler, r.klasse || "", r.rvTyp || "", String(r.datum), String(r.zaeit || ""),
      r.notiz || "", vun, "Nee", jetzt, r.dauer ? Number(r.dauer) : ""];
  });
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 5, zeilen.length, 2).setNumberFormat("@");
  sheet.getRange(start, 1, zeilen.length, 11).setValues(zeilen);

  let mailen = 0;
  const ouniMail = [];
  if (data.mailSchecken !== false) {
    const emailMap = {};
    getAktivSchuelerMatEmail().forEach((s) => { emailMap[s.matrikel] = s; });
    const proSchueler = {};
    liste.forEach((r) => { (proSchueler[r.schueler] = proSchueler[r.schueler] || []).push(r); });
    Object.keys(proSchueler).forEach((matrikel) => {
      const info = emailMap[matrikel];
      if (!info || !info.email) { ouniMail.push(matrikel); return; }
      try {
        rvMailSchecken_(info, proSchueler[matrikel], data.mailText);
        mailen++;
      } catch (e) { ouniMail.push(matrikel); }
    });
  }
  return { ok: true, ids: liste.map((r) => r.id), mailen, ouniMail };
}

function rvMailSchecken_(info, terminer, mailText) {
  const sortéiert = terminer.slice().sort((a, b) => rvIso_(a.datum).localeCompare(rvIso_(b.datum)) || String(a.zaeit).localeCompare(String(b.zaeit)));
  const zeilen = sortéiert.map((r) => "  • " + r.rvTyp + ": " + rvDeeg_(r.datum) + " " + r.datum +
    (r.zaeit ? " um " + r.zaeit + (r.dauer ? " (" + r.dauer + " Min.)" : "") : "") +
    (r.notiz ? "\n    Notiz: " + r.notiz : ""));
  const eent = sortéiert.length === 1;
  const betreff = "PPREN: " + (eent ? sortéiert[0].rvTyp + " geplangt fir de " + sortéiert[0].datum + (sortéiert[0].zaeit ? " um " + sortéiert[0].zaeit : "")
    : sortéiert.length + " Rendez-vousen geplangt");
  const text = mailText
    ? String(mailText)
        .replace(/\{virnumm\}/g, info.virnumm || "")
        .replace(/\{terminer\}/g, zeilen.join("\n"))
        .replace(/\{ziler\}/g, () => rvZiler_(info.matrikel))
        .replace(/\{meilensteng\}/g, () => rvMeilensteng_(info.matrikel, rvIso_(sortéiert[0].datum)))
    : "Hallo " + (info.virnumm || "") + ",\n\n" +
    (eent ? "Fir dech ass e Rendez-vous geplangt:" : "Fir dech si Rendez-vousen geplangt:") + "\n\n" +
    zeilen.join("\n") + "\n\n" +
    "Du fënns se och an dengem Suivi: https://pugu-prog.github.io/ppren/suivi.html\n" +
    "Am Unhang ass eng Kalennerdatei (.ics) — opmaachen, da kënnt den Termin mat Erënnerung an däin Handy-Kalenner.\n" +
    "2 Deeg virdrun kriss du nach eng Erënnerung per Mail.\n\nBereet dech w.e.g. gutt vir.";
  const ics = rvIcs_(sortéiert);
  MailApp.sendEmail({
    to: info.email,
    subject: betreff,
    body: text,
    attachments: [Utilities.newBlob(ics, "text/calendar", "PPREN_Rendezvous.ics")],
  });
}

function rvIso_(dmy) {
  const t = String(dmy).split(".");
  return t.length === 3 ? t[2] + "-" + t[1] + "-" + t[0] : String(dmy);
}

function rvDeeg_(dmy) {
  const t = String(dmy).split(".");
  if (t.length !== 3) return "";
  const d = new Date(Number(t[2]), Number(t[1]) - 1, Number(t[0]));
  return ["Sonndeg", "Méindeg", "Dënschdeg", "Mëttwoch", "Donneschdeg", "Freideg", "Samschdeg"][d.getDay()];
}

function rvIcs_(terminer) {
  const p = (n) => String(n).padStart(2, "0");
  const esc = (s) => String(s).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
  const stamp = Utilities.formatDate(new Date(), "UTC", "yyyyMMdd'T'HHmmss'Z'");
  const events = terminer.map((r) => {
    const [t, m, j] = String(r.datum).split(".");
    let start, end;
    if (r.zaeit) {
      const [hh, mi] = String(r.zaeit).split(":").map(Number);
      const ende = hh * 60 + mi + (Number(r.dauer) || 30);
      start = "DTSTART;TZID=Europe/Luxembourg:" + j + m + t + "T" + p(hh) + p(mi) + "00";
      end = "DTEND;TZID=Europe/Luxembourg:" + j + m + t + "T" + p(Math.floor(ende / 60)) + p(ende % 60) + "00";
    } else {
      const d = new Date(Number(j), Number(m) - 1, Number(t) + 1);
      start = "DTSTART;VALUE=DATE:" + j + m + t;
      end = "DTEND;VALUE=DATE:" + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
    }
    return ["BEGIN:VEVENT", "UID:" + (r.id || Utilities.getUuid()) + "@ppren", "DTSTAMP:" + stamp, start, end,
      "SUMMARY:" + esc("PPREN: " + r.rvTyp), r.notiz ? "DESCRIPTION:" + esc(r.notiz) : null,
      "BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:" + esc("PPREN: " + r.rvTyp), "TRIGGER:-P1D", "END:VALARM",
      "BEGIN:VALARM", "ACTION:DISPLAY", "DESCRIPTION:" + esc("PPREN: " + r.rvTyp), "TRIGGER:-PT30M", "END:VALARM",
      "END:VEVENT"].filter(Boolean).join("\r\n");
  });
  return ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PPREN//Rendezvousen//LB", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
    .concat(events, ["END:VCALENDAR"]).join("\r\n");
}

/** Ziler aus dem Projektplang vum Schüler (Übersicht, Spalt J „Projektplan-Details (JSON)“). */
function rvZiler_(matrikel) {
  const werte = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0].getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] !== matrikel) continue;
    let d = {};
    try { d = JSON.parse(werte[i][9] || "{}"); } catch (e) { d = {}; }
    const ziler = String(d.ziele || "").trim();
    if (!ziler) return "  (an dengem Projektplang stinn nach keng Ziler – dat musst du nach nohuelen!)";
    return ziler.split(/\n+/).map((z) => "  " + z.trim()).filter((z) => z.trim()).join("\n");
  }
  return "  (kee Projektplang fonnt)";
}

/**
 * Eegen Meilesteng vum Schüler bis zum Termin: déi lescht ~6 Wochen + al déi nach net erreecht sinn,
 * plus déi nächst 2 duerno. Offiziell Terminer a Vakanz ginn ausgeloos.
 */
function rvMeilensteng_(matrikel, terminIso) {
  const sheet = getUebersichtSS().getSheetByName("Meilensteng");
  if (!sheet) return "  (keng Meilesteng fonnt)";
  const werte = sheet.getDataRange().getValues();
  let liste = [];
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === matrikel) { try { liste = JSON.parse(werte[i][2] || "[]"); } catch (e) { liste = []; } break; }
  }
  liste = liste
    .filter((m) => m && m.titel && m.datum && m.quell !== "Offiziell" && m.quell !== "Vakanz")
    .map((m) => ({ titel: m.titel, iso: String(m.datum).slice(0, 10), erreecht: m.status === "Erreecht" }))
    .sort((a, b) => a.iso.localeCompare(b.iso));
  if (liste.length === 0) return "  (du hues nach keng eege Meilesteng – trag se an PPREN an!)";
  const t = terminIso.split("-").map(Number);
  const grenz = new Date(t[0], t[1] - 1, t[2] - 42);
  const grenzIso = Utilities.formatDate(grenz, "Europe/Luxembourg", "yyyy-MM-dd");
  const faelleg = liste.filter((m) => m.iso <= terminIso && (m.iso > grenzIso || !m.erreecht));
  const duerno = liste.filter((m) => m.iso > terminIso).slice(0, 2);
  const dmy = (iso) => iso.split("-").reverse().join(".");
  const zeil = (m, k) => "  " + k + " " + dmy(m.iso) + " – " + m.titel;
  const deel = [];
  if (faelleg.length) {
    deel.push("Bis zum Termin fälleg (weis, wat s du erreecht hues – a firwat eppes nach net fäerdeg ass):");
    faelleg.forEach((m) => deel.push(zeil(m, m.erreecht ? "✔" : "☐")));
  } else {
    deel.push("Bis zum Termin ass keen eegene Meilesteen fälleg – iwwerpréif deng Planung!");
  }
  if (duerno.length) {
    deel.push("Duerno geplangt (deng nächst Schrëtt):");
    duerno.forEach((m) => deel.push(zeil(m, "→")));
  }
  return deel.join("\n");
}
