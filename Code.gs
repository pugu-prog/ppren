/**
 * PPREN – Google Apps Script
 * ----------------------------------------------------------------
 * Zweck: Nimmt Formulardaten entgegen (Projektplan-Formular und
 * Bewertungsformular) und pflegt folgende Ordnerstruktur in Drive:
 *
 *   Projekte/
 *     2026-27/                      – automatisch berechnetes Schuljahr
 *       <Schüler>/
 *         _quelle_Projektplan_<Schüler>  – internes Google Doc (Arbeitskopie)
 *         Projektplan_<Schüler>.docx     – ECHTE Word-Datei, das ist die Datei,
 *                                          die Schüler/Lehrer öffnen/bearbeiten
 *         _quelle_Suivi_<Schüler>        – internes Google Doc (Arbeitskopie)
 *         Suivi_<Schüler>.docx           – ECHTE Word-Datei (Lehrer-Ansicht)
 *     2027-28/                      – nächstes Schuljahr, automatisch neu angelegt
 *       ...
 *
 * Die "_quelle_"-Dokumente sind interne Arbeitskopien in Google-Docs-
 * Format, die bei jeder Aktualisierung neu befüllt und anschließend als
 * .docx exportiert werden (File.getAs(MimeType.MICROSOFT_WORD)). Nur die
 * .docx-Dateien sind für Schüler/Lehrer gedacht und werden verlinkt.
 *
 * Das Schuljahr wird automatisch aus dem aktuellen Datum berechnet
 * (Beginn: September) – keine manuelle Anpassung pro Jahr nötig.
 * ----------------------------------------------------------------
 */

const FOLDER_ID = "12eeECkttG0U1-zRHM__age2FdZp3MOmy";
const TEMPLATE_DOC_ID = "1kFM0tOdYrtpPRD6pyU7Rqrz0bDp8-1lR";
const OVERVIEW_SHEET_ID = "1eAdAoDkiQMnowCV2sl1mrmTh3GT6W-_946uI06L3JAQ";
const LOGO_FILE_ID = "12x_Q2xM-olpOeF8luCF-9uX2oFiPPBNW";

const LEHRER_EMAILS = {
  "Guy Putz": "guy.putz@lycee.lu",
  "Pol Medernach": "pol.medernach@lycee.lu",
  "Sarah Blum": "sarah.blum@lycee.lu",
  "Tania Ludwig": "tania.ludwig@lycee.lu",
  "Tom Bleyer": "tom.bleyer@lycee.lu",
  "Salman Murad": "salman.murad@lycee.lu",
  "Alex Olinger": "alex.olinger@lycee.lu",
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("PPREN")
    .addItem("Bewertung entsperren…", "bewertungEntsperrenDialog")
    .addToUi();
}

function bewertungEntsperrenDialog() {
  const ui = SpreadsheetApp.getUi();
  const schuelerResp = ui.prompt("Bewertung entsperren", "Numm vum Schüler (genau wéi an der Tabell):", ui.ButtonSet.OK_CANCEL);
  if (schuelerResp.getSelectedButton() !== ui.Button.OK) return;
  const schueler = schuelerResp.getResponseText().trim();
  const periodeResp = ui.prompt("Bewertung entsperren", "Period (z.B. \"Semester 2\" oder \"Trimester 1\"):", ui.ButtonSet.OK_CANCEL);
  if (periodeResp.getSelectedButton() !== ui.Button.OK) return;
  const periode = periodeResp.getResponseText().trim();
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Bewertungen");
  const werte = sheet.getDataRange().getValues();
  const statusSpalte = werte[0].indexOf("Status");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler && werte[i][1] === periode) {
      sheet.getRange(i + 1, statusSpalte + 1).setValue("Entwurf");
      ui.alert("✅ Entspaart: " + schueler + " – " + periode);
      return;
    }
  }
  ui.alert("⚠️ Keng Bewertung fonnt fir: " + schueler + " – " + periode);
}

function doGet(e) {
  if (e.parameter && e.parameter.namen === "1") {
    return jsonResponse({ ok: true, personen: getAktivePersonen() });
  }

  const session = pruefSession(e.parameter && e.parameter.token);
  if (!session.valid) {
    return jsonResponse({ ok: false, error: "Net ugemellt." });
  }
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);

  const uebersichtSheet = ss.getSheets()[0];
  const uebersichtWerte = uebersichtSheet.getDataRange().getValues();
  let projekte = [];
  if (uebersichtWerte.length >= 2) {
    const header = uebersichtWerte[0];
    projekte = uebersichtWerte.slice(1).map((zeile) => {
      const obj = {};
      header.forEach((h, i) => { obj[h] = zeile[i]; });
      return obj;
    });
  }

  let bewertungen = [];
  const bewertungenSheet = ss.getSheetByName("Bewertungen");
  if (bewertungenSheet) {
    const bWerte = bewertungenSheet.getDataRange().getValues();
    if (bWerte.length >= 2) {
      const bHeader = bWerte[0];
      bewertungen = bWerte.slice(1).map((zeile) => {
        const obj = {};
        bHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let meilensteng = [];
  const meilenstengSheet = ss.getSheetByName("Meilensteng");
  if (meilenstengSheet) {
    const mWerte = meilenstengSheet.getDataRange().getValues();
    if (mWerte.length >= 2) {
      const mHeader = mWerte[0];
      meilensteng = mWerte.slice(1).map((zeile) => {
        const obj = {};
        mHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let wochenberichte = [];
  const wbSheet = ss.getSheetByName("Wochenberichte");
  if (wbSheet) {
    const wWerte = wbSheet.getDataRange().getValues();
    if (wWerte.length >= 2) {
      const wHeader = wWerte[0];
      wochenberichte = wWerte.slice(1).map((zeile) => {
        const obj = {};
        wHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let offiziellZaitplang = [];
  const ozSheet = ss.getSheetByName("OffiziellZaitplang");
  if (ozSheet) {
    const ozWerte = ozSheet.getDataRange().getValues();
    if (ozWerte.length >= 2) {
      const ozHeader = ozWerte[0];
      offiziellZaitplang = ozWerte.slice(1).map((zeile) => {
        const obj = {};
        ozHeader.forEach((h, i) => {
          let wert = zeile[i];
          if (h === "Datum" && wert instanceof Date) {
            wert = Utilities.formatDate(wert, "Europe/Luxembourg", "yyyy-MM-dd");
          }
          obj[h] = wert;
        });
        return obj;
      });
    }
  }

  let fachgespraeche = [];
  const fgSheet = ss.getSheetByName("Fachgespraeche");
  if (fgSheet) {
    const fgWerte = fgSheet.getDataRange().getValues();
    if (fgWerte.length >= 2) {
      const fgHeader = fgWerte[0];
      fachgespraeche = fgWerte.slice(1).map((zeile) => {
        const obj = {};
        fgHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let zieluewerpreiwungen = [];
  const zpSheet = ss.getSheetByName("Zieluewerpreiwungen");
  if (zpSheet) {
    const zpWerte = zpSheet.getDataRange().getValues();
    if (zpWerte.length >= 2) {
      const zpHeader = zpWerte[0];
      zieluewerpreiwungen = zpWerte.slice(1).map((zeile) => {
        const obj = {};
        zpHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let rendezvousen = [];
  const rvSheet = ss.getSheetByName("Rendezvousen");
  if (rvSheet) {
    const rvWerte = rvSheet.getDataRange().getValues();
    if (rvWerte.length >= 2) {
      const rvHeader = rvWerte[0];
      rendezvousen = rvWerte.slice(1).map((zeile) => {
        const obj = {};
        rvHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let budget = [];
  const budgetSheet = ss.getSheetByName("Budget");
  if (budgetSheet) {
    const budgetWerte = budgetSheet.getDataRange().getValues();
    if (budgetWerte.length >= 2) {
      const budgetHeader = budgetWerte[0];
      budget = budgetWerte.slice(1).map((zeile) => {
        const obj = {};
        budgetHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let ausgaben = [];
  const ausgabenSheet = ss.getSheetByName("Ausgaben");
  if (ausgabenSheet) {
    const ausgabenWerte = ausgabenSheet.getDataRange().getValues();
    if (ausgabenWerte.length >= 2) {
      const ausgabenHeader = ausgabenWerte[0];
      ausgaben = ausgabenWerte.slice(1).map((zeile) => {
        const obj = {};
        ausgabenHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let bewertungsraster = [];
  const brSheet = ss.getSheetByName("Bewertungsraster");
  if (brSheet) {
    const brWerte = brSheet.getDataRange().getValues();
    if (brWerte.length >= 2) {
      const brHeader = brWerte[0];
      bewertungsraster = brWerte.slice(1).map((zeile) => {
        const obj = {};
        brHeader.forEach((h, i) => { obj[h] = zeile[i]; });
        return obj;
      });
    }
  }

  let personen = getAktivePersonen();

  if (session.rolle === "Schüler") {
    const numm = session.numm;
    projekte = projekte.filter((p) => p.Schüler === numm);
    bewertungen = bewertungen.filter((b) => b.Schüler === numm);
    meilensteng = meilensteng.filter((m) => m.Schüler === numm);
    wochenberichte = wochenberichte.filter((w) => w.Schüler === numm);
    fachgespraeche = fachgespraeche.filter((f) => f.Schüler === numm);
    zieluewerpreiwungen = zieluewerpreiwungen.filter((z) => z.Schüler === numm);
    rendezvousen = rendezvousen.filter((r) => r.Schüler === numm);
    budget = budget.filter((b) => b.Schüler === numm);
    ausgaben = ausgaben.filter((a) => a.Schüler === numm);
  }

  return jsonResponse({ ok: true, projekte, bewertungen, meilensteng, wochenberichte, offiziellZaitplang, fachgespraeche, zieluewerpreiwungen, rendezvousen, budget, ausgaben, personen, bewertungsraster });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let url;
    if (data.typ === "bewertung") {
      url = schreibeBewertung(data);
    } else if (data.typ === "statusAendern") {
      aendereBewertungsStatus(data.schueler, data.periode, data.neuerStatus);
      return jsonResponse({ ok: true });
    } else if (data.typ === "meilensteng") {
      speichereMeilensteng(data.schueler, data.klasse, data.meilensteng);
      return jsonResponse({ ok: true });
    } else if (data.typ === "offiziellZaitplang") {
      speichereOffiziellZaitplang(data.klasse, data.meilensteng);
      return jsonResponse({ ok: true });
    } else if (data.typ === "wochenbericht") {
      const id = speichereWochenbericht(data);
      return jsonResponse({ ok: true, id });
    } else if (data.typ === "wochenberichtBewerten") {
      bewerteWochenbericht(data.id, data.punkte);
      return jsonResponse({ ok: true });
    } else if (data.typ === "wochenberichtEntschellegen") {
      entschellegWochenbericht(data.id, data.grond);
      return jsonResponse({ ok: true });
    } else if (data.typ === "fachgespraech") {
      const id = speichereFachgespraech(data);
      return jsonResponse({ ok: true, id });
    } else if (data.typ === "zieluewerpreiwung") {
      const id = speichereZieluewerpreiwung(data);
      return jsonResponse({ ok: true, id });
    } else if (data.typ === "login") {
      return jsonResponse(login(data.numm, data.pin));
    } else if (data.typ === "sessionPruefen") {
      return jsonResponse(pruefSession(data.token));
    } else if (data.typ === "logout") {
      logout(data.token);
      return jsonResponse({ ok: true });
    } else if (data.typ === "pinZuruecksetzen") {
      return jsonResponse(pinZuruecksetzen(data.numm, data.proffToken, data.neiesPasswuert));
    } else if (data.typ === "dokumentatiounLink") {
      return jsonResponse(speichereDokumentatiounLink(data.schueler, data.klasse, data.link));
    } else if (data.typ === "rendezvousPlangen") {
      return jsonResponse(plangRendezvous(data));
    } else if (data.typ === "rendezvousLaeschen") {
      return jsonResponse(läschRendezvous(data));
    } else if (data.typ === "kostenplanAgereechen") {
      return jsonResponse(kostenplanAgereechen(data));
    } else if (data.typ === "kostenplanGenehmegen") {
      return jsonResponse(kostenplanGenehmegen(data));
    } else if (data.typ === "ausgabSpäicheren") {
      return jsonResponse(speichereAusgab(data));
    } else if (data.typ === "ausgabLaeschen") {
      return jsonResponse(laeschAusgab(data));
    } else if (data.typ === "projektplanWiedereroeffnen") {
      return jsonResponse(projektplanWiedereroeffnen(data));
    } else if (data.typ === "personSpäicheren") {
      return jsonResponse(personSpäicheren(data));
    } else if (data.typ === "personDeaktivéieren") {
      return jsonResponse(personDeaktivéieren(data));
    } else if (data.typ === "personLoeschen") {
      return jsonResponse(personLoeschen(data));
    } else if (data.typ === "personenBulkSpäicheren") {
      return jsonResponse(personenBulkSpäicheren(data));
    } else if (data.typ === "bewertungsrasterSpäicheren") {
      return jsonResponse(bewertungsrasterSpäicheren(data));
    } else {
      url = erstelleOderAktualisiereProjektplan(data);
    }
    return jsonResponse({ ok: true, url });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function aendereBewertungsStatus(schueler, periode, neuerStatus) {
  if (neuerStatus !== "Entwurf" && neuerStatus !== "Finalisiert") {
    throw new Error("Ungültiger Status: " + neuerStatus);
  }
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Bewertungen");
  if (!sheet) throw new Error('Tab "Bewertungen" nicht gefunden.');
  const werte = sheet.getDataRange().getValues();
  const statusSpalte = werte[0].indexOf("Status");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler && werte[i][1] === periode) {
      sheet.getRange(i + 1, statusSpalte + 1).setValue(neuerStatus);
      return;
    }
  }
  throw new Error("Keine Bewertung gefunden für " + schueler + " – " + periode);
}

function speichereMeilensteng(schueler, klasse, meilensteng) {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Meilensteng");
  if (!sheet) {
    sheet = ss.insertSheet("Meilensteng");
    sheet.appendRow(["Schüler", "Klasse", "Meilensteng (JSON)", "Zuletzt aktualisiert"]);
  }
  const werte = sheet.getDataRange().getValues();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[schueler, klasse, JSON.stringify(meilensteng), jetzt]]);
      return;
    }
  }
  sheet.appendRow([schueler, klasse, JSON.stringify(meilensteng), jetzt]);
}

function synchroniséierProjektplangMeilensteng(schueler, klasse, planMeilensteng) {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Meilensteng");
  if (!sheet) {
    sheet = ss.insertSheet("Meilensteng");
    sheet.appendRow(["Schüler", "Klasse", "Meilensteng (JSON)", "Zuletzt aktualisiert"]);
  }
  const werte = sheet.getDataRange().getValues();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");

  let bestehend = [];
  let zeile = -1;
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) {
      zeile = i + 1;
      try { bestehend = JSON.parse(werte[i][2] || "[]"); } catch { bestehend = []; }
      break;
    }
  }

  const ouni_pp = bestehend.filter((m) => m.quell !== "Projektplang");

  const bestehendPpMap = {};
  bestehend.filter((m) => m.quell === "Projektplang").forEach((m) => { bestehendPpMap[m.id] = m; });

  const neiPp = (planMeilensteng || [])
    .filter((m) => m.beschreibung && m.beschreibung.trim())
    .map((m, i) => {
      const id = "pp-" + i;
      const alt = bestehendPpMap[id];
      return {
        id, quell: "Projektplang", kategorie: "Projektplang",
        datum: m.datum || "", titel: m.beschreibung,
        status: alt?.status || "Ausstoend", notiz: alt?.notiz || "",
      };
    });

  const neiKomplett = [...ouni_pp, ...neiPp];

  if (zeile > 0) {
    sheet.getRange(zeile, 1, 1, 4).setValues([[schueler, klasse, JSON.stringify(neiKomplett), jetzt]]);
  } else {
    sheet.appendRow([schueler, klasse, JSON.stringify(neiKomplett), jetzt]);
  }
}

function getOffiziellZaitplangSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("OffiziellZaitplang");
  if (!sheet) {
    sheet = ss.insertSheet("OffiziellZaitplang");
    sheet.appendRow(["Klasse", "Datum", "Titel", "Kategorie"]);
  }
  return sheet;
}

function speichereOffiziellZaitplang(klasse, meilensteng) {
  const sheet = getOffiziellZaitplangSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = werte.length - 1; i >= 1; i--) {
    if (werte[i][0] === klasse) sheet.deleteRow(i + 1);
  }
  meilensteng.forEach((m) => {
    sheet.appendRow([klasse, m.datum, m.titel, m.kategorie || "Event"]);
  });
}

function seedOffiziellZaitplangVunZeitplangHtml() {
  const EIN_1GSE = [
    { datum: "2026-09-07", titel: "Präsentatioun PPREN", kategorie: "Event" },
    { datum: "2026-10-07", titel: "Projektplang erstellen", kategorie: "Event" },
    { datum: "2026-10-27", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-10-28", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-12-08", titel: "Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2026-12-08", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-12-09", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-01-12", titel: "Start 2. Semester", kategorie: "Event" },
    { datum: "2027-01-27", titel: "Fachgespréicher", kategorie: "Event" },
    { datum: "2027-02-02", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-02-03", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-03-03", titel: "Prüfung & Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2027-03-09", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-03-10", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-04-27", titel: "Presentatioun & Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2027-04-28", titel: "Presentatioun & Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2027-05-04", titel: "Fachgespréicher", kategorie: "Event" },
    { datum: "2027-05-05", titel: "Fachgespréicher", kategorie: "Event" },
  ];
  const EIN_2GSE = [
    { datum: "2026-09-07", titel: "Präsentatioun PPREN", kategorie: "Event" },
    { datum: "2026-10-07", titel: "Projektplang erstellen", kategorie: "Event" },
    { datum: "2026-10-27", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-10-28", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-12-08", titel: "Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2026-12-08", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2026-12-09", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-01-05", titel: "Start 2. Trimester", kategorie: "Event" },
    { datum: "2027-01-27", titel: "Fachgespréicher", kategorie: "Event" },
    { datum: "2027-02-02", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-02-03", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-03-03", titel: "Prüfung & Ofgab Dokumentatioun", kategorie: "Ofgab" },
    { datum: "2027-03-09", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-03-10", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-04-13", titel: "Start 3e Trimester", kategorie: "Event" },
    { datum: "2027-05-04", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-05-05", titel: "Ziler iwwerpréiwen", kategorie: "Iwwerpréiwung" },
    { datum: "2027-06-22", titel: "Presentatioun & Ofgab Dokumentatioun (Datum nach net confirméiert — evtl. 29.06)", kategorie: "Ofgab" },
    { datum: "2027-06-29", titel: "Fachgespréicher (Datum nach net confirméiert — evtl. 30.06)", kategorie: "Event" },
  ];
  speichereOffiziellZaitplang("1GSE", EIN_1GSE);
  speichereOffiziellZaitplang("2GSE", EIN_2GSE);
  Logger.log("✅ OffiziellZaitplang gefëllt: " + EIN_1GSE.length + " Zeile(n) fir 1GSE, " + EIN_2GSE.length + " Zeile(n) fir 2GSE.");
}

function getWochenberichteSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Wochenberichte");
  if (!sheet) {
    sheet = ss.insertSheet("Wochenberichte");
    sheet.appendRow([
      "ID", "Schüler", "Klasse", "Periode", "Woche", "Datum Verfassung",
      "Zusammenfassung", "Fortschritt", "Anhänge", "Status",
      "Punkte (JSON)", "Zuletzt aktualisiert", "Betreuer", "Betreuer2", "Entschëllegung-Grond",
      "Meilensteen",
    ]);
  }
  return sheet;
}

function holBetreuerFuerSchueler(schueler) {
  const kombinéiert = getBetreuerFuerSchueler(schueler);
  if (!kombinéiert) return [];
  return kombinéiert.split(",").map((s) => s.trim()).filter(Boolean).slice(0, 2);
}

function notifizéierBetreuerNeierBericht(schueler, betreuerListe, woche) {
  betreuerListe.forEach((betreuer) => {
    const email = LEHRER_EMAILS[betreuer];
    if (!email) return;
    try {
      MailApp.sendEmail({
        to: email,
        subject: "PPREN: Neien Wochenbericht vun " + schueler + " (" + woche + ")",
        body: schueler + " huet e Wochenbericht fir d'Woch " + woche + " ofginn.\n\n" +
          "Hei korrigéieren: https://pugu-prog.github.io/ppren-projektplan/wochenberichte-korrigeieren.html",
      });
    } catch (e) { }
  });
}

function speichereWochenbericht(data) {
  const sheet = getWochenberichteSheet();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const werte = sheet.getDataRange().getValues();

  let anhaengeText = data.anhaenge || "";
  if (data.anhaengeDateien && data.anhaengeDateien.length > 0) {
    const ordner = getStudentFolder(data.schueler);
    const wbOrdnerIter = ordner.getFoldersByName("Wochenberichte_Anhänge");
    const wbOrdner = wbOrdnerIter.hasNext() ? wbOrdnerIter.next() : ordner.createFolder("Wochenberichte_Anhänge");
    const links = [];
    data.anhaengeDateien.forEach((datei) => {
      try {
        const blob = Utilities.newBlob(Utilities.base64Decode(datei.inhaltBase64), datei.mimeType, datei.dateiName);
        const neieDatei = wbOrdner.createFile(blob);
        neieDatei.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        links.push(neieDatei.getUrl());
      } catch (e) {
      }
    });
    if (links.length > 0) {
      anhaengeText = (anhaengeText ? anhaengeText + "\n" : "") + links.join("\n");
    }
  }

  if (data.id) {
    for (let i = 1; i < werte.length; i++) {
      if (werte[i][0] === data.id) {
        if (werte[i][9] === "Bewäert") {
          throw new Error("Dëse Bericht ass schonn bewäert an kann net méi geännert ginn.");
        }
        sheet.getRange(i + 1, 1, 1, 16).setValues([[
          data.id, data.schueler, data.klasse, data.periode, data.woche, data.datumVerfassung,
          data.zusammenfassung, data.fortschritt, anhaengeText, "Agereecht", "", jetzt,
          werte[i][12] || "", werte[i][13] || "", "", data.meilensteen || "",
        ]]);
        return data.id;
      }
    }
  }

  for (let i = 1; i < werte.length; i++) {
    if (werte[i][1] === data.schueler && werte[i][3] === data.periode && werte[i][4] === data.woche) {
      throw new Error("Fir dës Woch (" + data.woche + ") hues du schonn en Bericht ofginn. Änner de bestehende Bericht amplaz en neien unzeleeën.");
    }
  }

  const betreuerListe = holBetreuerFuerSchueler(data.schueler);
  const neiId = Utilities.getUuid();
  sheet.appendRow([
    neiId, data.schueler, data.klasse, data.periode, data.woche, data.datumVerfassung,
    data.zusammenfassung, data.fortschritt, anhaengeText, "Agereecht", "", jetzt,
    betreuerListe[0] || "", betreuerListe[1] || "", "", data.meilensteen || "",
  ]);
  notifizéierBetreuerNeierBericht(data.schueler, betreuerListe, data.woche);
  return neiId;
}

function bewerteWochenbericht(id, punkte) {
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === id) {
      sheet.getRange(i + 1, 10, 1, 3).setValues([["Bewäert", JSON.stringify(punkte), jetzt]]);
      return;
    }
  }
  throw new Error("Wochebericht net fonnt: " + id);
}

function entschellegWochenbericht(id, grond) {
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === id) {
      sheet.getRange(i + 1, 10, 1, 3).setValues([["Entschëllegt", "", jetzt]]);
      sheet.getRange(i + 1, 15).setValue(grond || "");
      return;
    }
  }
  throw new Error("Wochebericht net fonnt: " + id);
}

function berechneWochenberichtSumme(schueler, periode) {
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  let summe = 0;
  let anzahlBewäert = 0;
  let anzahlAgereecht = 0;
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][1] === schueler && werte[i][3] === periode) {
      const status = werte[i][9];
      if (status === "Entschëllegt") continue;
      anzahlAgereecht++;
      if (status === "Bewäert") {
        anzahlBewäert++;
        try {
          const p = JSON.parse(werte[i][10]);
          summe += (p.zusammenfassung || 0) + (p.fortschritt || 0) + (p.anhaenge || 0) + (p.grammatik || 0);
        } catch (e) { }
      } else if (status === "Verpasst") {
        anzahlBewäert++;
      }
    }
  }
  return { summe, max: anzahlBewäert * 5, anzahlBewäert, anzahlAgereecht };
}

function getFachgespraechSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Fachgespraeche");
  if (!sheet) {
    sheet = ss.insertSheet("Fachgespraeche");
    sheet.appendRow([
      "ID", "Schüler", "Klasse", "Periode", "Datum", "Variante",
      "Froen (JSON)", "Fachwissen-Punkte", "Reflexioun-Punkte", "Gesamt",
      "Notiz", "Zuletzt aktualisiert",
    ]);
  }
  return sheet;
}

function speichereFachgespraech(data) {
  const sheet = getFachgespraechSheet();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const fachwissenSumme = (data.froen || []).reduce((s, f) => s + (Number(f.punkte) || 0), 0);
  const reflexionPunkte = data.reflexionAktiv ? (Number(data.reflexionPunkte) || 0) : null;
  const gesamt = fachwissenSumme + (reflexionPunkte || 0);
  const werte = sheet.getDataRange().getValues();

  const zeileWerte = [
    data.id || Utilities.getUuid(), data.schueler, data.klasse, data.periode, data.datum, data.variante,
    JSON.stringify(data.froen || []), fachwissenSumme, reflexionPunkte, gesamt, data.notiz || "", jetzt,
  ];

  if (data.id) {
    for (let i = 1; i < werte.length; i++) {
      if (werte[i][0] === data.id) {
        sheet.getRange(i + 1, 1, 1, zeileWerte.length).setValues([zeileWerte]);
        return data.id;
      }
    }
  }
  sheet.appendRow(zeileWerte);
  return zeileWerte[0];
}

function getZieluewerpreiwungSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Zieluewerpreiwungen");
  if (!sheet) {
    sheet = ss.insertSheet("Zieluewerpreiwungen");
    sheet.appendRow([
      "ID", "Schüler", "Klasse", "Periode", "Termin", "Datum",
      "Punkte", "Max", "Notiz", "Zuletzt aktualisiert",
    ]);
  }
  return sheet;
}

function speichereZieluewerpreiwung(data) {
  const sheet = getZieluewerpreiwungSheet();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const werte = sheet.getDataRange().getValues();

  const zeileWerte = [
    data.id || Utilities.getUuid(), data.schueler, data.klasse, data.periode, data.termin, data.datum,
    Number(data.punkte) || 0, Number(data.max) || 0, data.notiz || "", jetzt,
  ];

  if (data.id) {
    for (let i = 1; i < werte.length; i++) {
      if (werte[i][0] === data.id) {
        sheet.getRange(i + 1, 1, 1, zeileWerte.length).setValues([zeileWerte]);
        return data.id;
      }
    }
  }
  sheet.appendRow(zeileWerte);
  return zeileWerte[0];
}

function getPersonenSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Personen");
  if (!sheet) {
    sheet = ss.insertSheet("Personen");
    sheet.appendRow(["LoginId", "Virnumm", "Numm", "Rolle", "Klasse", "Email", "Aktiv"]);
  }
  return sheet;
}

function generéierMatrikel(klasse) {
  const joer = getSchuljahrLabel().split("-")[0].slice(-2);
  const prefix = joer + "-" + klasse + "-";
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  let max = 0;
  werte.slice(1).forEach((z) => {
    const loginId = String(z[0] || "");
    if (loginId.indexOf(prefix) === 0) {
      const n = parseInt(loginId.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(2, "0");
}

function seedTestSchuelerMatrikel() {
  const TEST_SCHUELER = [
    { virnumm: "Léa", klasse: "1GSE" },
    { virnumm: "Ben", klasse: "2GSE" },
    { virnumm: "Noah", klasse: "1GSE" },
    { virnumm: "Tom", klasse: "2GSE" },
  ];
  const personenSheet = getPersonenSheet();
  const loginSheet = getLoginSheet();

  [personenSheet, loginSheet].forEach((sheet) => {
    const werte = sheet.getDataRange().getValues();
    const rolleSpalt = sheet === personenSheet ? 3 : 1;
    for (let i = werte.length - 1; i >= 1; i--) {
      if (werte[i][rolleSpalt] === "Schüler") sheet.deleteRow(i + 1);
    }
  });

  const resultat = [];
  TEST_SCHUELER.forEach((s) => {
    const matrikel = generéierMatrikel(s.klasse);
    personenSheet.appendRow([matrikel, s.virnumm, "", "Schüler", s.klasse, "", "Jo"]);
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    const salt = Utilities.getUuid();
    loginSheet.appendRow([matrikel, "Schüler", s.klasse, hashPin(pin, salt), salt]);
    resultat.push({ matrikel, virnumm: s.virnumm, klasse: s.klasse, pin });
  });
  Logger.log("✅ " + resultat.length + " Test-Schüler mat neier Matrikel ugeluecht:\n" + JSON.stringify(resultat, null, 2));
  return resultat;
}

function getAktivePersonen() {
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  return werte.slice(1)
    .filter((z) => z[6] !== "Nee")
    .map((z) => ({ numm: z[0], virnumm: z[1] || "", rolle: z[3], klasse: z[4] }));
}

function personSpäicheren(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Persounen verwalten." };
  }
  let loginId = (data.numm || "").trim();
  if (!loginId && data.rolle === "Schüler" && data.klasse) {
    loginId = generéierMatrikel(data.klasse);
  }
  data.numm = loginId;
  if (!loginId || !data.rolle) return { ok: false, error: "LoginId (Matrikel/Numm) a Roll erfuerderlech." };
  if (data.rolle === "Schüler" && !data.klasse) return { ok: false, error: "Klasse erfuerderlech fir Schüler." };

  const virnumm = (data.virnumm || "").trim();
  const vollNumm = data.rolle === "Prof" ? loginId : "";
  const gewenschtPasswuertBestehend = (data.pin || "").trim();
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (String(werte[i][0] || "").trim() === loginId) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([[loginId, virnumm || werte[i][1], vollNumm, data.rolle, data.klasse || "", data.email || werte[i][5], "Jo"]]);
      const loginSheet = getLoginSheet();
      const loginWerte = loginSheet.getDataRange().getValues();
      let loginZeilIndex = -1;
      for (let j = 1; j < loginWerte.length; j++) {
        if (String(loginWerte[j][0] || "").trim() === loginId) { loginZeilIndex = j; break; }
      }
      let neiPin = null;
      if (loginZeilIndex === -1) {
        const pin = gewenschtPasswuertBestehend.length >= 6
          ? gewenschtPasswuertBestehend
          : String(Math.floor(1000 + Math.random() * 9000));
        const salt = Utilities.getUuid();
        loginSheet.appendRow([loginId, data.rolle, data.klasse || "", hashPin(pin, salt), salt]);
        neiPin = pin;
      } else if (gewenschtPasswuertBestehend.length >= 6) {
        const salt = Utilities.getUuid();
        loginSheet.getRange(loginZeilIndex + 1, 4, 1, 2).setValues([[hashPin(gewenschtPasswuertBestehend, salt), salt]]);
        neiPin = gewenschtPasswuertBestehend;
      }
      return { ok: true, neiPin, matrikel: loginId };
    }
  }
  const email = (data.email || "").trim();
  if (data.rolle === "Prof" && !email) return { ok: false, error: "Email ass Pflicht bei enger neier Persoun." };

  sheet.appendRow([loginId, virnumm, vollNumm, data.rolle, data.klasse || "", email, "Jo"]);

  const gewenschtPasswuert = (data.pin || "").trim();
  const pin = gewenschtPasswuert.length >= 6
    ? gewenschtPasswuert
    : String(Math.floor(1000 + Math.random() * 9000));
  const salt = Utilities.getUuid();
  getLoginSheet().appendRow([loginId, data.rolle, data.klasse || "", hashPin(pin, salt), salt]);
  return { ok: true, neiPin: pin, matrikel: loginId };
}

function personenBulkSpäicheren(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Persounen verwalten." };
  }
  const rolle = data.rolle === "Prof" ? "Prof" : "Schüler";
  if (rolle === "Schüler" && !data.klasse) {
    return { ok: false, error: "Klasse erfuerderlech fir Schüler." };
  }
  const zeilen = (data.nimm || [])
    .map((n) => String(n).trim())
    .filter(Boolean);
  if (zeilen.length === 0) return { ok: false, error: "Keng Zeilen ugi." };

  const sheet = getPersonenSheet();
  const loginSheet = getLoginSheet();
  const bestehendNimm = new Set(sheet.getDataRange().getValues().slice(1).map((z) => z[2]));
  const resultater = [];

  zeilen.forEach((zeil) => {
    if (rolle === "Prof") {
      const loginId = zeil;
      if (bestehendNimm.has(loginId)) {
        resultater.push({ numm: loginId, pin: null, status: "scho do" });
        return;
      }
      sheet.appendRow([loginId, "", loginId, "Prof", "", "", "Jo"]);
      const pin = String(Math.floor(1000 + Math.random() * 9000));
      const salt = Utilities.getUuid();
      loginSheet.appendRow([loginId, "Prof", "", hashPin(pin, salt), salt]);
      resultater.push({ numm: loginId, pin, status: "nei" });
      return;
    }
    const virnumm = zeil;
    const matrikel = generéierMatrikel(data.klasse);
    sheet.appendRow([matrikel, virnumm, "", "Schüler", data.klasse, "", "Jo"]);
    const pin = String(Math.floor(1000 + Math.random() * 9000));
    const salt = Utilities.getUuid();
    loginSheet.appendRow([matrikel, "Schüler", data.klasse, hashPin(pin, salt), salt]);
    resultater.push({ numm: matrikel, virnumm, pin, status: "nei" });
  });

  return { ok: true, resultater };
}

function personDeaktivéieren(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Persounen (de)aktivéieren." };
  }
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.numm) {
      sheet.getRange(i + 1, 7).setValue(data.aktiv ? "Jo" : "Nee");
      return { ok: true };
    }
  }
  return { ok: false, error: "Net fonnt." };
}

function personLoeschen(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Persounen permanent läschen." };
  }
  if (!data.numm) return { ok: false, error: "Numm erfuerderlech." };
  if (data.numm === session.numm) {
    return { ok: false, error: "Du kanns däin eegene Kont net läschen (fir de Zougang net aus Versinn ze verléieren)." };
  }

  const personenSheet = getPersonenSheet();
  const personenWerte = personenSheet.getDataRange().getValues();
  let fonnt = false;
  for (let i = personenWerte.length - 1; i >= 1; i--) {
    if (personenWerte[i][0] === data.numm) {
      personenSheet.deleteRow(i + 1);
      fonnt = true;
    }
  }

  const loginSheet = getLoginSheet();
  const loginWerte = loginSheet.getDataRange().getValues();
  for (let i = loginWerte.length - 1; i >= 1; i--) {
    if (loginWerte[i][0] === data.numm) loginSheet.deleteRow(i + 1);
  }

  const sessionsSheet = getSessionsSheet();
  const sessionsWerte = sessionsSheet.getDataRange().getValues();
  for (let i = sessionsWerte.length - 1; i >= 1; i--) {
    if (sessionsWerte[i][1] === data.numm) sessionsSheet.deleteRow(i + 1);
  }

  if (!fonnt) return { ok: false, error: "Persoun net fonnt." };
  return { ok: true };
}

function getBewertungsrasterSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Bewertungsraster");
  if (!sheet) {
    sheet = ss.insertSheet("Bewertungsraster");
    sheet.appendRow([
      "Klasse", "Periode", "GruppeOrder", "GruppeTitel", "GruppeTyp", "ItemOrder",
      "ItemName", "ItemMax", "ItemIndikatoren", "WochenberichteMax", "WochenberichteLabel", "OffizielleSkala",
    ]);
  }
  return sheet;
}

const BEWERTUNGSRASTER_STANDARD = {
  "2GSE": {
    "Trimester 1": {
      offizielleSkala: 270, wochenberichteMax: 60, wochenberichteLabel: "Wochenberichte (12 × 5P)",
      gruppen: [
        { titel: "Planung / Themenfindung", typ: "planung", items: [
          { name: "Projektplan", max: 30, indikatoren: "Beschreibung des Projektes (15P); Motivation/Bezug zu Umweltwissenschaften (3P); Ziele (6P); Aufgabenteilung bei Gruppenarbeit (6P); Kostenplan falls erforderlich" },
        ] },
        { titel: "Umsetzung des Projektes und Dokumentation", typ: "umsetzung", items: [
          { name: "Bewältigung des Themas", max: 30, indikatoren: "Arbeit an vereinbarten Zielsetzungen; Zielüberprüfung an 2 Terminen (je 11P); Autonomie (8P)" },
          { name: "Dokumentation", max: 30, indikatoren: "Einleitung, Vorwort, Titelseite, Inhaltsverzeichnis, Verzeichnisse; Recherche/Theorie (20P)" },
        ] },
        { titel: "Methodologie 1 — Bewerbungsmappe & Praktikumssuche", typ: "methodologie", items: [
          { name: "Bewerbungsmappe (FR/DE/EN)", max: 30, indikatoren: "Alle Dokumente vorhanden; Richtlinien für Bewerbungsschreiben/Lebenslauf beachtet" },
          { name: "Überprüfung der erledigten Arbeiten", max: 10, indikatoren: "Methodische, gezielte Kontrolle; Nutzung von Korrekturressourcen; Fehler erkannt und korrigiert" },
          { name: "Praktikumsplatzsuche", max: 20, indikatoren: "Alle Suchmöglichkeiten ausgeschöpft; Termine eingehalten" },
        ] },
        { titel: "Methodologie 2 — Kommunikation & Recherche", typ: "methodologie", items: [
          { name: "Schriftliche Kommunikation per E-Mail", max: 40, indikatoren: "Standardsätze/Höflichkeitsformeln bekannt; passender Wortschatz; korrekte Struktur" },
          { name: "Elektronische Kommunikationsmittel", max: 10, indikatoren: "Modalitäten bekannt und korrekt angewendet; Anweisungen korrekt umgesetzt" },
          { name: "Effiziente Nutzung verschiedener Quellen", max: 10, indikatoren: "Informationsquellen unterschieden und projektrelevant ausgewählt" },
        ] },
      ],
    },
    "Trimester 2": {
      offizielleSkala: 190, wochenberichteMax: 40, wochenberichteLabel: "Meilensteine / Wochenberichte (8 × 5P)",
      gruppen: [
        { titel: "Umsetzung des Projektes und Dokumentation", typ: "umsetzung", items: [
          { name: "Bewältigung des Themas", max: 30, indikatoren: "Arbeit an Zielsetzungen; 1–2 Prüftermine; Autonomie" },
          { name: "Dokumentation", max: 30, indikatoren: "Theoretische Grundlagen (25P); Quellenverzeichnis (5P)" },
        ] },
        { titel: "Fachgespräch", typ: "fachgespraech", items: [
          { name: "Fachwissen", max: 30, indikatoren: "4 Fragen mit Skizze oder 5 Fragen kompetent & fachlich korrekt beantwortet, Fachbegriffe benutzt" },
        ] },
        { titel: "Schriftliche Prüfung", typ: "dokumentation", items: [
          { name: "Schriftliche Prüfung", max: 30, indikatoren: "Vorbereitung auf das schriftliche Examen; Analyse und Reflexion des Projekts" },
        ] },
        { titel: "Methodologie", typ: "methodologie", items: [
          { name: "Digitale Kompetenzen", max: 15, indikatoren: "Beherrschung fortgeschrittener Funktionen der verwendeten Software (MS Office)" },
          { name: "Recherche, Analyse und Dokumentation", max: 15, indikatoren: "Theoretische Grundlagen & Quellenverzeichnis; Vielfalt & Relevanz der Quellen" },
        ] },
      ],
    },
    "Trimester 3": {
      offizielleSkala: 300, wochenberichteMax: 60, wochenberichteLabel: "Meilensteine / Wochenberichte (8–12 × 5P)",
      gruppen: [
        { titel: "Vorgehensweise bei der Umsetzung des Projektes", typ: "umsetzung", items: [
          { name: "Bewältigung des Themas", max: 30, indikatoren: "Arbeit an Zielsetzungen; 1–2 Prüftermine; Autonomie (8P)" },
          { name: "Sachliche Qualität", max: 60, indikatoren: "Komponenten/Produkt funktionieren wie geplant (30P); Qualität inkl. Film überzeugt (30P)" },
        ] },
        { titel: "Schriftliche Arbeit / Dokumentation", typ: "dokumentation", items: [
          { name: "Darstellung", max: 8, indikatoren: "Übersichtlich strukturiert & gegliedert; sorgfältiges, kompaktes Layout" },
          { name: "Sprache", max: 7, indikatoren: "Verständlich, flüssig, prägnant; korrekte Grammatik/Satzbau/Rechtschreibung/Zeichensetzung" },
          { name: "Zitate, Quellen, Verzeichnisse", max: 5, indikatoren: "Vollständige Quellenangaben & Verzeichnisse" },
          { name: "Sachliche Qualität", max: 40, indikatoren: "Vollständig, fachlich begründet und korrekt" },
        ] },
        { titel: "Präsentation", typ: "praesentation", items: [
          { name: "Struktur", max: 6, indikatoren: "Durchdacht und ansprechend gestaltet" },
          { name: "Inhalt", max: 12, indikatoren: "Repräsentativ, mit Einblick in Ergebnisse und Produkt" },
          { name: "Sprache", max: 3, indikatoren: "Flüssig, verständlich, korrekt" },
          { name: "Interaktion", max: 3, indikatoren: "Interesse geweckt, Publikum angemessen einbezogen" },
          { name: "Fragen", max: 6, indikatoren: "Kompetent beantwortet" },
        ] },
        { titel: "Fachgespräch", typ: "fachgespraech", items: [
          { name: "Fachwissen", max: 20, indikatoren: "4/5 Fragen kompetent beantwortet, Fachbegriffe benutzt" },
          { name: "Eigenbewertung / Reflexion", max: 10, indikatoren: "Analyse und Reflexion des Projekts" },
        ] },
        { titel: "Methodologie", typ: "methodologie", items: [
          { name: "Vorbereitung und Verwaltung des Projekts", max: 20, indikatoren: "Checkliste vor Abgabe genutzt; Zwischentermine eingehalten" },
          { name: "Teamarbeit (PowerPoint-Präsentation)", max: 10, indikatoren: "Organisation der Notizen; individuelle Beiträge; Koordination/Kommunikation" },
        ] },
      ],
    },
  },
  "1GSE": {
    "Semester 1": {
      offizielleSkala: 210, wochenberichteMax: 60, wochenberichteLabel: "Wochenberichte (12 × 5P)",
      gruppen: [
        { titel: "Planung / Themenfindung", typ: "planung", items: [
          { name: "Projektplan", max: 30, indikatoren: "Beschreibung des Projektes (15P); Motivation/Bezug zu Umweltwissenschaften (3P); Ziele (6P); Meilensteine bis Ende 1. Semester; Aufgabenteilung bei Gruppenarbeit (6P); Kostenplan falls erforderlich" },
        ] },
        { titel: "Methodologie", typ: "methodologie", items: [
          { name: "Reflexion", max: 15, indikatoren: "Zeitplanung; Kostenplanung; Eigene Entwicklung (Initiative, Autonomie, Gelerntes); Qualität des Projektergebnisses; Teamwork" },
          { name: "Zusammenfassen", max: 15, indikatoren: "Bedeutung des Themas; Stand der Wissenschaft/Technik; Fragestellung; Methodisches Vorgehen; wichtigste Resultate/Erkenntnisse/Argumente" },
          { name: "Zitieren", max: 15, indikatoren: "Zitate und Quellenangabe nach APA (Bücher, Artikel, PDF-Dokumente, Internetseiten, Video)" },
          { name: "Überprüfung der erledigten Arbeiten", max: 15, indikatoren: "Methodische, gezielte Kontrolle; Nutzung von Korrekturressourcen; Fehler erkannt und korrigiert" },
        ] },
        { titel: "Umsetzung des Projektes und Dokumentation", typ: "umsetzung", items: [
          { name: "Bewältigung des Themas", max: 30, indikatoren: "Arbeit an vereinbarten Zielsetzungen; Zielüberprüfung an 2 Terminen (je 11P); Autonomie (8P)" },
          { name: "Dokumentation", max: 30, indikatoren: "Einleitung, Vorwort, Titelseite, Inhaltsverzeichnis, Verzeichnisse; Recherche/Theorie (20P)" },
        ] },
      ],
    },
    "Semester 2": {
      offizielleSkala: 330, wochenberichteMax: 60, wochenberichteLabel: "Meilensteine / Wochenberichte (12 × 5P)",
      gruppen: [
        { titel: "Fachgespräch", typ: "fachgespraech", items: [
          { name: "Fachwissen", max: 30, indikatoren: "4 Fragen mit Skizze oder 5 Fragen kompetent & fachlich korrekt beantwortet, Fachbegriffe benutzt" },
        ] },
        { titel: "Schriftliche Prüfung", typ: "dokumentation", items: [
          { name: "Schriftliche Prüfung", max: 30, indikatoren: "Vorbereitung auf das schriftliche Examen; Analyse und Reflexion des Projekts" },
        ] },
        { titel: "Vorgehensweise bei der Umsetzung des Projektes", typ: "umsetzung", items: [
          { name: "Bewältigung des Themas", max: 30, indikatoren: "Zielsetzungen (10P); 1–2 Prüftermine (2×6P); Autonomie (8P)" },
          { name: "Sachliche Qualität", max: 60, indikatoren: "Komponenten/Produkt funktionieren wie geplant (30P); Qualität inkl. Film überzeugt (30P)" },
        ] },
        { titel: "Schriftliche Arbeit / Dokumentation", typ: "dokumentation", items: [
          { name: "Darstellung", max: 8, indikatoren: "Übersichtlich strukturiert & gegliedert; sorgfältiges, kompaktes Layout" },
          { name: "Sprache", max: 7, indikatoren: "Verständlich, flüssig, prägnant; korrekte Grammatik/Satzbau/Rechtschreibung/Zeichensetzung" },
          { name: "Zitate, Quellen, Verzeichnisse", max: 5, indikatoren: "Vollständige Quellenangaben & Verzeichnisse" },
          { name: "Sachliche Qualität", max: 40, indikatoren: "Vollständig, fachlich begründet und korrekt" },
        ] },
        { titel: "Präsentation", typ: "praesentation", items: [
          { name: "Struktur", max: 6, indikatoren: "Durchdacht und ansprechend gestaltet" },
          { name: "Inhalt", max: 12, indikatoren: "Repräsentativ, mit Einblick in Ergebnisse und Produkt" },
          { name: "Sprache", max: 3, indikatoren: "Flüssig, verständlich, korrekt" },
          { name: "Interaktion", max: 3, indikatoren: "Interesse geweckt, Publikum angemessen einbezogen" },
          { name: "Fragen", max: 6, indikatoren: "Kompetent beantwortet" },
        ] },
        { titel: "Methodologie", typ: "methodologie", items: [
          { name: "Umfragen, Interview", max: 15, indikatoren: "Planung, Durchführung und Auswertung" },
          { name: "Film", max: 15, indikatoren: "Erstellen eines Storyboards (Drehbuch)" },
        ] },
      ],
    },
  },
};

function schreiwBewertungsrasterZeilen(sheet, klasse, periode, def) {
  def.gruppen.forEach((g, gi) => {
    g.items.forEach((item, ii) => {
      sheet.appendRow([
        klasse, periode, gi, g.titel, g.typ, ii, item.name, item.max, item.indikatoren || "",
        def.wochenberichteMax, def.wochenberichteLabel, def.offizielleSkala,
      ]);
    });
  });
}

function seedBewertungsrasterVunHardcoded() {
  const sheet = getBewertungsrasterSheet();
  const werte = sheet.getDataRange().getValues();
  const bestehendKombis = new Set(werte.slice(1).map((z) => z[0] + "||" + z[1]));
  let neiZuel = 0;
  Object.keys(BEWERTUNGSRASTER_STANDARD).forEach((klasse) => {
    Object.keys(BEWERTUNGSRASTER_STANDARD[klasse]).forEach((periode) => {
      if (bestehendKombis.has(klasse + "||" + periode)) return;
      schreiwBewertungsrasterZeilen(sheet, klasse, periode, BEWERTUNGSRASTER_STANDARD[klasse][periode]);
      neiZuel++;
    });
  });
  Logger.log("✅ Bewertungsraster gesät fir " + neiZuel + " Klass/Period-Kombinatioun(en).");
}

function bewertungsrasterSpäicheren(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen d'Bewertungsraster änneren." };
  }
  if (!data.klasse || !data.periode) return { ok: false, error: "Klasse a Period erfuerderlech." };
  if (!data.gruppen || data.gruppen.length === 0) return { ok: false, error: "Op d'mannst eng Grupp erfuerderlech." };

  const sheet = getBewertungsrasterSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = werte.length - 1; i >= 1; i--) {
    if (werte[i][0] === data.klasse && werte[i][1] === data.periode) sheet.deleteRow(i + 1);
  }
  schreiwBewertungsrasterZeilen(sheet, data.klasse, data.periode, {
    gruppen: data.gruppen,
    wochenberichteMax: Number(data.wochenberichteMax) || 0,
    wochenberichteLabel: data.wochenberichteLabel || "",
    offizielleSkala: Number(data.offizielleSkala) || 0,
  });
  return { ok: true };
}

function getRendezvousenSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Rendezvousen");
  if (!sheet) {
    sheet = ss.insertSheet("Rendezvousen");
    sheet.appendRow(["ID", "Schüler", "Klasse", "Typ", "Datum", "Zäit", "Notiz", "Erstallt vum", "Erënnert", "Zuletzt aktualisiert"]);
  }
  return sheet;
}

function plangRendezvous(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Rendez-vousen plangen." };
  }
  const sheet = getRendezvousenSheet();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const neiId = Utilities.getUuid();
  sheet.appendRow([
    neiId, data.schueler, data.klasse, data.rvTyp, data.datum, data.zaeit || "",
    data.notiz || "", data.erstalltVum || "", "Nee", jetzt,
  ]);
  return { ok: true, id: neiId };
}

function läschRendezvous(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen Rendez-vousen läschen." };
  }
  const sheet = getRendezvousenSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.id) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, error: "Net fonnt." };
}

const BUDGET_STANDARD_MAX = 500;

function getBudgetSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Budget");
  if (!sheet) {
    sheet = ss.insertSheet("Budget");
    sheet.appendRow(["Schüler", "Klasse", "Status", "Max", "GenehmegtVum", "GenehmegtUm"]);
  }
  return sheet;
}

function getAusgabenSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Ausgaben");
  if (!sheet) {
    sheet = ss.insertSheet("Ausgaben");
    sheet.appendRow(["ID", "Schüler", "Klasse", "Bezeechnung", "Betrag", "Datum", "Kategorie", "Beleg-Link", "Erstallt"]);
  }
  return sheet;
}

function kostenplanAgereechen(data) {
  const session = pruefSession(data.token);
  if (!session.valid) return { ok: false, error: "Net ugemellt." };
  const sheet = getBudgetSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.schueler) {
      return { ok: true, status: werte[i][2] };
    }
  }
  sheet.appendRow([data.schueler, data.klasse || "", "Ageraecht", BUDGET_STANDARD_MAX, "", ""]);
  return { ok: true, status: "Ageraecht" };
}

function kostenplanGenehmegen(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen e Kostenplang genehmegen." };
  }
  const sheet = getBudgetSheet();
  const werte = sheet.getDataRange().getValues();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.schueler) {
      sheet.getRange(i + 1, 3, 1, 4).setValues([["Genehmegt", werte[i][3] || BUDGET_STANDARD_MAX, session.numm, jetzt]]);
      return { ok: true };
    }
  }
  sheet.appendRow([data.schueler, data.klasse || "", "Genehmegt", BUDGET_STANDARD_MAX, session.numm, jetzt]);
  return { ok: true };
}

function speichereAusgab(data) {
  const session = pruefSession(data.token);
  if (!session.valid) return { ok: false, error: "Net ugemellt." };

  const budgetSheet = getBudgetSheet();
  const budgetWerte = budgetSheet.getDataRange().getValues();
  const budgetZeil = budgetWerte.find((z) => z[0] === data.schueler);
  if (!budgetZeil || budgetZeil[2] !== "Genehmegt") {
    return { ok: false, error: "De Kostenplang ass nach net genehmegt — Ausgaben nach net méiglech." };
  }

  let belegLink = "";
  if (data.belegBase64 && data.belegDateiName) {
    try {
      const ordner = getStudentFolder(data.schueler);
      const belegOrdnerIter = ordner.getFoldersByName("Ausgaben_Belege");
      const belegOrdner = belegOrdnerIter.hasNext() ? belegOrdnerIter.next() : ordner.createFolder("Ausgaben_Belege");
      const blob = Utilities.newBlob(Utilities.base64Decode(data.belegBase64), data.belegMimeType || "application/octet-stream", data.belegDateiName);
      const neieDatei = belegOrdner.createFile(blob);
      neieDatei.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      belegLink = neieDatei.getUrl();
    } catch (e) { }
  }

  const sheet = getAusgabenSheet();
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const neiId = Utilities.getUuid();
  sheet.appendRow([neiId, data.schueler, data.klasse || "", data.bezeechnung, Number(data.betrag) || 0, data.datum || "", data.kategorie || "", belegLink, jetzt]);
  return { ok: true, id: neiId };
}

function laeschAusgab(data) {
  const session = pruefSession(data.token);
  if (!session.valid) return { ok: false, error: "Net ugemellt." };
  const sheet = getAusgabenSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.id) {
      if (session.rolle !== "Prof" && werte[i][1] !== session.numm) {
        return { ok: false, error: "Just deng eege Ausgaben." };
      }
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: "Net fonnt." };
}

function sendeRendezvousErennerungen() {
  const sheet = getRendezvousenSheet();
  const werte = sheet.getDataRange().getValues();
  const haut = new Date();
  haut.setHours(0, 0, 0, 0);
  const zielDatum = new Date(haut);
  zielDatum.setDate(haut.getDate() + 2);
  const zielDatumStr = Utilities.formatDate(zielDatum, "Europe/Luxembourg", "dd.MM.yyyy");
  const emailMap = {};
  getAktivSchuelerMatEmail().forEach((s) => { emailMap[s.matrikel] = { email: s.email, virnumm: s.virnumm }; });

  for (let i = 1; i < werte.length; i++) {
    const [id, schueler, , typ, datum, zaeit, notiz, , erënnert] = werte[i];
    if (erënnert === "Jo") continue;
    if (datum !== zielDatumStr) continue;

    const info = emailMap[schueler];
    if (info && info.email) {
      try {
        MailApp.sendEmail({
          to: info.email,
          subject: "PPREN: " + typ + " an 2 Deeg (" + datum + (zaeit ? ", " + zaeit : "") + ")",
          body: "Hallo " + (info.virnumm || "") + ",\n\nDenk drun: en/eng " + typ + " ass geplangt fir de " + datum +
            (zaeit ? " um " + zaeit : "") + ".\n" + (notiz ? "\nNotiz vum Prof: " + notiz + "\n" : "") +
            "\nBereet dech w.e.g. gutt vir.",
        });
      } catch (e) { }
    }
    sheet.getRange(i + 1, 9).setValue("Jo");
  }
}

function installRendezvousTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "sendeRendezvousErennerungen") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendeRendezvousErennerungen").timeBased().everyDays(1).atHour(7).create();
  Logger.log("✅ Trigger installéiert: Rendez-vous-Erënnerung all Dag ~7.00h.");
}

const WOCHENBERICHT_START_ISO = "2026-09-21";

function aktuellWocheLabel() {
  const heute = new Date();
  const tag = heute.getDay();
  const diffZuMontag = tag === 0 ? -6 : 1 - tag;
  const montag = new Date(heute);
  montag.setDate(heute.getDate() + diffZuMontag);
  const sonndeg = new Date(montag);
  sonndeg.setDate(montag.getDate() + 6);
  const fmt = (d) => Utilities.formatDate(d, "Europe/Luxembourg", "dd.MM.yyyy");
  return fmt(montag) + " – " + fmt(sonndeg);
}

function nachVirWochenberichtStart() {
  const heute = new Date();
  const tag = heute.getDay();
  const diffZuMontag = tag === 0 ? -6 : 1 - tag;
  const montag = new Date(heute);
  montag.setDate(heute.getDate() + diffZuMontag);
  const montagIso = Utilities.formatDate(montag, "Europe/Luxembourg", "yyyy-MM-dd");
  return montagIso < WOCHENBERICHT_START_ISO;
}

function getAktivSchuelerMatEmail() {
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  return werte.slice(1)
    .filter((z) => z[6] !== "Nee" && z[3] === "Schüler")
    .map((z) => ({ matrikel: z[0], virnumm: z[1] || "", klasse: z[4], email: z[5] || "" }));
}

function sendeErennerungen() {
  if (nachVirWochenberichtStart()) return;
  const woche = aktuellWocheLabel();
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  const ofginn = new Set(werte.slice(1).filter((z) => z[4] === woche).map((z) => z[1]));

  getAktivSchuelerMatEmail().forEach((s) => {
    if (ofginn.has(s.matrikel)) return;
    if (!s.email) return;
    try {
      MailApp.sendEmail({
        to: s.email,
        subject: "PPREN: Wochenbericht net vergiessen (haut Owend 22.00h zou)",
        body: "Hallo " + (s.virnumm || "") + ",\n\nDu hues fir dës Woch (" + woche + ") nach kee Wochenbericht ofginn. " +
          "D'Ofgab ass haut Owend um 22.00h — duerno gëtt automatesch 0 Punkte gesat.\n\n" +
          "Hei ofginn: https://pugu-prog.github.io/ppren-projektplan/wochenbericht.html",
      });
    } catch (e) { }
  });
}

function schliesseWochenberichterAb() {
  if (nachVirWochenberichtStart()) return;
  const woche = aktuellWocheLabel();
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  const ofginn = new Set(werte.slice(1).filter((z) => z[4] === woche).map((z) => z[1]));
  const jetzt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");

  getAktivSchuelerMatEmail().forEach((s) => {
    if (ofginn.has(s.matrikel)) return;
    const eegen = werte.slice(1).filter((z) => z[1] === s.matrikel);
    const letztPeriode = eegen.length > 0 ? eegen[eegen.length - 1][3] : (s.klasse === "2GSE" ? "Trimester 1" : "Semester 1");
    const betreuerListe = holBetreuerFuerSchueler(s.matrikel);
    const neiId = Utilities.getUuid();
    sheet.appendRow([
      neiId, s.matrikel, s.klasse, letztPeriode, woche, "",
      "(Net ofginn — automatesch op 0 Punkte gesat)", "", "", "Verpasst",
      JSON.stringify({ zusammenfassung: 0, fortschritt: 0, anhaenge: 0, grammatik: 0 }), jetzt,
      betreuerListe[0] || "", betreuerListe[1] || "", "",
    ]);
  });
}

function raeumWochenberichterVirStart() {
  const sheet = getWochenberichteSheet();
  const werte = sheet.getDataRange().getValues();
  let geläscht = 0;
  for (let i = werte.length - 1; i >= 1; i--) {
    const wocheLabel = werte[i][4];
    const m = String(wocheLabel || "").match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (!m) continue;
    const isoDatum = `${m[3]}-${m[2]}-${m[1]}`;
    if (isoDatum < WOCHENBERICHT_START_ISO) {
      sheet.deleteRow(i + 1);
      geläscht++;
    }
  }
  Logger.log("✅ " + geläscht + " Wochenberichter-Zeile(n) virun der Startwoch (" + WOCHENBERICHT_START_ISO + ") geläscht.");
}

function installWochenberichtTriggers() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "sendeErennerungen" || t.getHandlerFunction() === "schliesseWochenberichterAb") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("sendeErennerungen").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(10).create();
  ScriptApp.newTrigger("schliesseWochenberichterAb").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
  Logger.log("✅ Trigger installéiert: Erënnerung Sonndes ~10.00h, Ofschloss Sonndes ~23.00h.");
}

function getSchuljahrLabel() {
  const jetzt = new Date();
  const jahr = jetzt.getFullYear();
  const monat = jetzt.getMonth() + 1;
  if (monat >= 9) {
    return `${jahr}-${String(jahr + 1).slice(-2)}`;
  }
  return `${jahr - 1}-${String(jahr).slice(-2)}`;
}

function getSchuljahrOrdner() {
  const hauptordner = DriveApp.getFolderById(FOLDER_ID);
  const label = getSchuljahrLabel();
  const bestehende = hauptordner.getFoldersByName(label);
  if (bestehende.hasNext()) return bestehende.next();
  return hauptordner.createFolder(label);
}

/**
 * Fir sichtbaren Text an Dokumenter (Titelblat, Iwwerschrëften): bei
 * Schüler gëtt "Virnumm (Matrikel)" ugewisen (z.B. "Ben (26-1GSE-01)"),
 * bei Proffen bleift de ganze Numm onverännert. Sichtbar am Dokument ≠
 * Ordner-/Dateinumm (déi bleiwen bewosst nëmmen d'Matrikel, fir d'Drive-
 * URLen/Pad unanonym ze halen).
 */
function anzeigeNummFirSchueler(matrikelOderNumm) {
  const sheet = getPersonenSheet();
  const werte = sheet.getDataRange().getValues();
  const gesicht = String(matrikelOderNumm || "").trim();
  const zeil = werte.slice(1).find((z) => String(z[0] || "").trim() === gesicht);
  if (!zeil) return matrikelOderNumm || "";
  if (zeil[3] === "Schüler") {
    return zeil[1] ? `${zeil[1]} (${zeil[0]})` : zeil[0];
  }
  return matrikelOderNumm;
}

function getStudentFolder(schueler) {
  const jahrOrdner = getSchuljahrOrdner();
  const bestehende = jahrOrdner.getFoldersByName(schueler);
  if (bestehende.hasNext()) return bestehende.next();
  return jahrOrdner.createFolder(schueler);
}

function getOrCreateDoc(ordner, dateiname, ausVorlage) {
  const bestehende = ordner.getFilesByName(dateiname);
  if (bestehende.hasNext()) {
    return { doc: DocumentApp.openById(bestehende.next().getId()), neu: false };
  }
  if (ausVorlage) {
    const vorlage = DriveApp.getFileById(TEMPLATE_DOC_ID);
    const kopie = vorlage.makeCopy(dateiname, ordner);
    return { doc: DocumentApp.openById(kopie.getId()), neu: true };
  }
  const neuesDoc = DocumentApp.create(dateiname);
  DriveApp.getFileById(neuesDoc.getId()).moveTo(ordner);
  return { doc: neuesDoc, neu: true };
}

function exportiereAlsWord(doc, ordner, sichtbarerName) {
  doc.saveAndClose();
  const exportUrl = `https://docs.google.com/document/d/${doc.getId()}/export?format=docx`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  const wordBlob = response.getBlob().setName(sichtbarerName + ".docx");
  const alte = ordner.getFilesByName(sichtbarerName + ".docx");
  while (alte.hasNext()) alte.next().setTrashed(true);
  const wordDatei = ordner.createFile(wordBlob);
  return wordDatei.getUrl();
}

function exportiereAlsPdf(doc, ordner, sichtbarerName) {
  const exportUrl = `https://docs.google.com/document/d/${doc.getId()}/export?format=pdf`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  const pdfBlob = response.getBlob().setName(sichtbarerName + ".pdf");
  const alte = ordner.getFilesByName(sichtbarerName + ".pdf");
  while (alte.hasNext()) alte.next().setTrashed(true);
  const pdfDatei = ordner.createFile(pdfBlob);
  pdfDatei.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return pdfDatei.getUrl();
}

function erstelleOderAktualisiereProjektplan(data) {
  const bisherigerStatus = getAktuellenStatus(data.schueler);
  if (bisherigerStatus && bisherigerStatus.startsWith("Frei") && data.status !== "Entwurf") {
    throw new Error("Der Projektplan ist bereits freigegeben und gesperrt (Stand Ende Oktober). Nur die Betreuer/-in kann ihn zur Bearbeitung wieder öffnen.");
  }
  if (bisherigerStatus && bisherigerStatus !== "Entwurf" && data.status === "Entwurf") {
    throw new Error("Nëmmen de Betreier kann en areechte Projektplang zréck op Entwurf setzen. Kontaktéier w.e.g. däi Betreier.");
  }

  const ordner = getStudentFolder(data.schueler);
  const { doc } = getOrCreateDoc(ordner, `_quelle_Projektplan_${data.schueler}`, true);
  const body = doc.getBody();
  body.clear();

  fuelleProjektplanDokument(doc, data);
  const projektplanUrl = exportiereAlsWord(doc, ordner, `Projektplan_${data.schueler}`);

  if (data.betreuerEmail) {
    DriveApp.getFileById(doc.getId()).addEditor(data.betreuerEmail);
  }

  aktualisiereUebersicht(data, projektplanUrl, null);
  synchroniséierProjektplangMeilensteng(data.schueler, data.klasse, data.meilensteine || []);

  if ((data.status || "").startsWith("Frei")) {
    const suiviUrl = uebernehmeProjektplanInSuivi(data);
    aktualisiereUebersicht(data, projektplanUrl, suiviUrl);
  }

  return projektplanUrl;
}

function projektplanWiedereroeffnen(data) {
  const session = pruefSession(data.proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen e Projektplang zréck op Entwurf setzen." };
  }
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const werte = sheet.getDataRange().getValues();
  const header = werte[0] || [];
  const statusSpalte = header.indexOf("Status");
  if (statusSpalte === -1) return { ok: false, error: 'Spalt "Status" net fonnt.' };
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.schueler) {
      sheet.getRange(i + 1, statusSpalte + 1).setValue("Entwurf");
      return { ok: true };
    }
  }
  return { ok: false, error: "Schüler net fonnt." };
}

function getAktuellenStatus(schueler) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) return werte[i][4];
  }
  return null;
}

function fuegeFarbBalkenEin(body, farbe, hoehePt, breitePt) {
  const tabelle = body.appendTable([[""]]);
  tabelle.setBorderWidth(0);
  if (breitePt) {
    tabelle.setColumnWidth(0, breitePt);
  }
  const zeile = tabelle.getRow(0);
  zeile.setMinimumHeight(hoehePt || 4);
  const zelle = zeile.getCell(0);
  zelle.setBackgroundColor(farbe);
  zelle.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
}

const LTETT_LOGO_URL = "https://raw.githubusercontent.com/pugu-prog/ppren-projektplan/main/LTEtt_Logo.png";
const SE_LOGO_URL = "https://raw.githubusercontent.com/pugu-prog/ppren-projektplan/main/LogoSE_weisHintergrund.png";

function erstelleDeckblattBildViaSlides(dokumentTyp, data, behalten) {
  const PAGE_W = 595, PAGE_H = 842;
  const seBlob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
  const ltettBlob = UrlFetchApp.fetch(LTETT_LOGO_URL).getBlob();

  const erstellteDatei = Slides.Presentations.create({
    title: "_temp_deckblatt_" + Utilities.getUuid(),
    pageSize: {
      width: { magnitude: PAGE_W, unit: "PT" },
      height: { magnitude: PAGE_H, unit: "PT" },
    },
  });
  const praesentation = SlidesApp.openById(erstellteDatei.presentationId);
  praesentation.getSlides().forEach((s, i) => { if (i > 0) s.remove(); });
  const slide = praesentation.getSlides()[0];
  slide.getShapes().forEach((sh) => sh.remove());

  const GREEN = "#4CAF50", BLACK = "#1c2621", ORANGE = "#E8952E",
        BLUE = "#3A6EA5", GRAY = "#55625a", LIGHTGRAY = "#93a098", TEAL = "#3A9BB5";
  const margin = 57;

  const textBox = (text, left, top, width, height, size, color, bold) => {
    const box = slide.insertTextBox(text, left, top, width, height);
    const stil = box.getText().getTextStyle();
    stil.setFontSize(size).setForegroundColor(color).setBold(!!bold).setFontFamily("Arial");
    return box;
  };
  const farbBalken = (left, top, width, height, farbe) => {
    const balken = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, height);
    balken.getFill().setSolidFill(farbe);
    balken.getBorder().setTransparent();
  };

  const kleinBreite = 210;
  const kleinBild = slide.insertImage(seBlob.copyBlob());
  const seVerhaeltnis = kleinBild.getHeight() / kleinBild.getWidth();
  kleinBild.setLeft(PAGE_W - 150).setTop(-100).setWidth(kleinBreite).setHeight(kleinBreite * seVerhaeltnis);

  textBox("SCIENCES", margin, 70, 320, 42, 28, GREEN, true);
  textBox("ENVIRONNEMENTALES", margin, 120, 430, 42, 28, GREEN, true);
  textBox("ËMWELTWËSSENSCHAFTEN", margin, 175, 430, 30, 18, BLACK, true);
  farbBalken(margin, 215, 210, 8, ORANGE);

  textBox("PROJET PERSONNEL ENCADRE (PPREN)", margin, 243, 430, 26, 13, BLUE, true);
  textBox(dokumentTyp, margin, 277, 300, 32, 20, BLUE, true);
  textBox(`${data.klasse || ""} – Schuljahr ${getSchuljahrLabel()} – ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "MMMM yyyy")}`, margin, 313, 430, 26, 13, BLUE, true);

  const betreierListe = [data.betreuer, data.betreuer2].filter((b) => b && b.trim() !== "");
  const betreierText = betreierListe.length > 0 ? betreierListe.join(", ") : "–";
  const betreierLabel = betreierListe.length > 1 ? "Betreier/-innen" : "Betreuer/-in";
  textBox(`${anzeigeNummFirSchueler(data.schueler)}  ·  ${betreierLabel}: ${betreierText}`, margin, 353, 430, 24, 13, GRAY, false);
  textBox(`Stand: ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm")}`, margin, 375, 300, 18, 9, LIGHTGRAY, false);

  const grossBreite = 260;
  const grossBild = slide.insertImage(seBlob.copyBlob());
  grossBild.setLeft(-50).setTop(580).setWidth(grossBreite).setHeight(grossBreite * seVerhaeltnis);

  const ltettBreite = 170;
  const ltettBild = slide.insertImage(ltettBlob.copyBlob());
  const ltettVerhaeltnis = ltettBild.getHeight() / ltettBild.getWidth();
  const ltettHoehe = ltettBreite * ltettVerhaeltnis;
  const ltettLeft = 210 + 30;
  const ltettTop = 580 + (grossBreite * seVerhaeltnis - ltettHoehe) / 2;
  ltettBild.setLeft(ltettLeft).setTop(ltettTop).setWidth(ltettBreite).setHeight(ltettHoehe);

  farbBalken(ltettLeft, ltettTop + ltettHoehe + 15, ltettBreite, 8, TEAL);

  SpreadsheetApp.flush();
  Utilities.sleep(4000);

  const presentationId = praesentation.getId();
  const slideId = slide.getObjectId();
  Logger.log("Slide zum Ankucken: https://docs.google.com/presentation/d/" + presentationId + "/edit");

  const exportUrl = `https://docs.google.com/presentation/d/${presentationId}/export/png?id=${presentationId}&pageid=${slideId}`;
  const antwort = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  Logger.log("Export-HTTP-Status: " + antwort.getResponseCode() + ", Bildgröße: " + antwort.getBlob().getBytes().length + " Bytes");

  if (antwort.getResponseCode() !== 200) {
    if (!behalten) DriveApp.getFileById(presentationId).setTrashed(true);
    throw new Error("Slide-Export fehlgeschlagen (HTTP " + antwort.getResponseCode() + ")");
  }
  const bildBlob = antwort.getBlob().setName("Deckblatt.png");
  if (!behalten) DriveApp.getFileById(presentationId).setTrashed(true);
  return bildBlob;
}

function baueDeckblattHtml(dokumentTyp, data) {
  const seBlob = UrlFetchApp.fetch(SE_LOGO_URL).getBlob();
  const ltettBlob = UrlFetchApp.fetch(LTETT_LOGO_URL).getBlob();
  const seB64 = Utilities.base64Encode(seBlob.getBytes());
  const ltettB64 = Utilities.base64Encode(ltettBlob.getBytes());

  const betreierListe = [data.betreuer, data.betreuer2].filter((b) => b && b.trim() !== "");
  const betreierText = betreierListe.length > 0 ? betreierListe.join(", ") : "–";
  const betreierLabel = betreierListe.length > 1 ? "Betreier/-innen" : "Betreuer/-in";
  const klasseZeile = `${data.klasse || ""} – Schuljahr ${getSchuljahrLabel()} – ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "MMMM yyyy")}`;

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
  *{margin:0;padding:0;box-sizing:border-box;}
  html{width:1508.03px;height:2132.79px;overflow:hidden;}
  body{width:1508.03px;height:2132.79px;position:relative;font-family:Arial,Helvetica,sans-serif;overflow:hidden;background:#ffffff;}
  .page-frame{position:absolute;top:0;left:0;width:793.7px;height:1122.52px;overflow:hidden;background:#ffffff;transform:scale(1.9);transform-origin:top left;}
  .el{position:absolute;}
  .top-logo{left:359.06px;top:-332.6px;width:755.91px;}
  .big-logo{left:-207.87px;top:626.35px;width:755.91px;}
  .ltett-logo{left:506.46px;top:933.05px;width:228.66px;}
  .txt{left:88.82px;font-weight:700;white-space:nowrap;}
  .txt-sciences{top:232.44px;font-size:45.33px;color:#00A94F;transform:translateY(-50%);}
  .txt-environnementales{top:311.81px;font-size:45.33px;color:#00A94F;transform:translateY(-50%);}
  .txt-emwelt{top:394.96px;font-size:37.33px;color:#000000;font-weight:400;transform:translateY(-50%);}
  .bar-orange{left:88.82px;top:436.54px;width:277.8px;height:8.8px;background:#F67E1B;}
  .txt-ppren{top:498.9px;font-size:24.67px;color:#4576B5;transform:translateY(-50%);}
  .txt-typ{top:540.47px;font-size:22.67px;color:#4576B5;transform:translateY(-50%);}
  .txt-klasse{top:582.05px;font-size:22.67px;color:#4576B5;transform:translateY(-50%);}
  .txt-schueler{top:618.5px;font-size:16px;font-weight:400;color:#55625a;}
  .bar-teal{left:500.78px;top:1070.36px;width:277.8px;height:8.8px;background:#48B5CA;}
  </style></head><body>
  <div class="page-frame">
  <img class="el top-logo" src="data:image/png;base64,${seB64}">
  <img class="el big-logo" src="data:image/png;base64,${seB64}">
  <img class="el ltett-logo" src="data:image/png;base64,${ltettB64}">
  <div class="el txt txt-sciences">SCIENCES</div>
  <div class="el txt txt-environnementales">ENVIRONNEMENTALES</div>
  <div class="el txt txt-emwelt">ËMWELTWËSSENSCHAFTEN</div>
  <div class="el bar-orange"></div>
  <div class="el txt txt-ppren">PROJET PERSONNEL ENCADRÉ (PPREN)</div>
  <div class="el txt txt-typ">${dokumentTyp}</div>
  <div class="el txt txt-klasse">${klasseZeile}</div>
  <div class="el txt txt-schueler">${anzeigeNummFirSchueler(data.schueler)}  ·  ${betreierLabel}: ${betreierText}</div>
  <div class="el bar-teal"></div>
  </div>
  </body></html>`;
}

function erstelleDeckblattBildViaPDFShift(dokumentTyp, data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("PDFSHIFT_API_KEY");
  if (!apiKey) {
    throw new Error("PDFSHIFT_API_KEY fehlt in den Script Properties.");
  }
  const html = baueDeckblattHtml(dokumentTyp, data);

  const antwort = UrlFetchApp.fetch("https://api.pdfshift.io/v3/convert/png", {
    method: "post",
    headers: { "X-API-Key": apiKey, "Content-Type": "application/json" },
    payload: JSON.stringify({ source: html, fullpage: true, viewport: "1508x2133" }),
    muteHttpExceptions: true,
  });

  if (antwort.getResponseCode() !== 200) {
    throw new Error("PDFShift-Fehler (HTTP " + antwort.getResponseCode() + "): " + antwort.getContentText().substring(0, 300));
  }
  return antwort.getBlob().setName("Deckblatt.png");
}

function fuegeDeckblattEin(body, dokumentTyp, data) {
  try {
    const bildBlob = erstelleDeckblattBildViaPDFShift(dokumentTyp, data);
    const bild = body.appendImage(bildBlob);
    bild.setWidth(468).setHeight(468 * 1123 / 794);
    body.appendPageBreak();
    return;
  } catch (e1) {
    body.appendParagraph("[PDFShift-Deckblatt fehlgeschlagen (" + e1.message + "), versuche Slides-Methode]")
      .setForegroundColor("#c0392b").setFontSize(8);
  }

  try {
    const bildBlob = erstelleDeckblattBildViaSlides(dokumentTyp, data);
    const bild = body.appendImage(bildBlob);
    bild.setWidth(468).setHeight(468 * 842 / 595);
    body.appendPageBreak();
    return;
  } catch (e2) {
    body.appendParagraph("[Slides-Deckblatt fehlgeschlagen (" + e2.message + "), einfache Version wird verwendet]")
      .setForegroundColor("#c0392b").setFontSize(8);
  }

  fuegeDeckblattEinFallback(body, dokumentTyp, data);
}

function fuegeDeckblattEinFallback(body, dokumentTyp, data) {
  let seBlob = null;
  try {
    seBlob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
  } catch (e) {
    body.appendParagraph("[SE-Logo konnte nicht geladen werden: " + e.message + "]")
      .setForegroundColor("#c0392b").setFontSize(9);
  }
  let ltettBlob = null;
  try {
    ltettBlob = UrlFetchApp.fetch(LTETT_LOGO_URL).getBlob();
  } catch (e) {
    body.appendParagraph("[LTEtt-Logo konnte nicht geladen werden: " + e.message + "]")
      .setForegroundColor("#c0392b").setFontSize(9);
  }

  if (seBlob) {
    const kopfAbsatz = body.appendParagraph("").setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    const kleinesBild = kopfAbsatz.appendInlineImage(seBlob.copyBlob());
    const kleinBreite = 130;
    const kleinVerhaeltnis = kleinesBild.getHeight() / kleinesBild.getWidth();
    kleinesBild.setWidth(kleinBreite).setHeight(Math.round(kleinBreite * kleinVerhaeltnis));
  }

  body.appendParagraph("");
  body.appendParagraph("SCIENCES")
    .setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setBold(true).setFontSize(26).setForegroundColor("#4CAF50").setLineSpacing(1.5);
  body.appendParagraph("ENVIRONNEMENTALES")
    .setBold(true).setFontSize(26).setForegroundColor("#4CAF50").setLineSpacing(1.5);
  body.appendParagraph("ËMWELTWËSSENSCHAFTEN")
    .setBold(true).setFontSize(17).setForegroundColor("#1c2621").setLineSpacing(1.5);
  body.appendParagraph("");
  fuegeFarbBalkenEin(body, "#E8952E", 8, 230);
  body.appendParagraph("");
  body.appendParagraph("PROJET PERSONNEL ENCADRE (PPREN)")
    .setBold(true).setFontSize(13).setForegroundColor("#3A6EA5").setSpacingAfter(6);
  body.appendParagraph(dokumentTyp)
    .setBold(true).setFontSize(20).setForegroundColor("#3A6EA5").setSpacingAfter(6);
  body.appendParagraph(`${data.klasse || ""} – Schuljahr ${getSchuljahrLabel()} – ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "MMMM yyyy")}`)
    .setBold(true).setFontSize(13).setForegroundColor("#3A6EA5");
  body.appendParagraph("");
  const betreierListe = [data.betreuer, data.betreuer2].filter((b) => b && b.trim() !== "");
  const betreierText = betreierListe.length > 0 ? betreierListe.join(", ") : "–";
  const betreierLabel = betreierListe.length > 1 ? "Betreier/-innen" : "Betreuer/-in";
  body.appendParagraph(`${anzeigeNummFirSchueler(data.schueler)}  ·  ${betreierLabel}: ${betreierText}`)
    .setFontSize(13).setForegroundColor("#55625a");
  body.appendParagraph(`Stand: ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm")}`)
    .setFontSize(9).setForegroundColor("#93a098");
  body.appendParagraph("");
  body.appendParagraph("");

  if (seBlob || ltettBlob) {
    const logoTabelle = body.appendTable([["", ""]]);
    logoTabelle.setBorderWidth(0);
    const logoZeile = logoTabelle.getRow(0);
    const linkeZelle = logoZeile.getCell(0);
    const rechteZelle = logoZeile.getCell(1);
    linkeZelle.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
    rechteZelle.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(24).setPaddingRight(6);

    if (seBlob) {
      const grossesBild = linkeZelle.insertImage(0, seBlob.copyBlob());
      const grossBreite = 200;
      const grossVerhaeltnis = grossesBild.getHeight() / grossesBild.getWidth();
      grossesBild.setWidth(grossBreite).setHeight(Math.round(grossBreite * grossVerhaeltnis));
    }
    if (ltettBlob) {
      const ltettBild = rechteZelle.insertImage(0, ltettBlob.copyBlob());
      const ltettBreite = 160;
      const ltettVerhaeltnis = ltettBild.getHeight() / ltettBild.getWidth();
      ltettBild.setWidth(ltettBreite).setHeight(Math.round(ltettBreite * ltettVerhaeltnis));
    }
  }

  body.appendParagraph("").setSpacingBefore(12);
  fuegeFarbBalkenEin(body, "#3A9BB5", 8, 180);
  body.appendPageBreak();
}

function fuelleProjektplanDokument(doc, data) {
  const body = doc.getBody();

  const meilensteineText = (data.meilensteine || [])
    .map((m) => `• ${m.datum || "(kein Datum)"} — ${m.beschreibung}`)
    .join("\n");

  const kostenText = data.kostenplanNoetig
    ? (data.kosten || []).map((k) => `• ${k.bezeichnung}: ${k.betrag} €`).join("\n")
    : "nicht erforderlich";

  const werte = {
    "{{SCHUELER}}": data.schueler || "",
    "{{KLASSE}}": data.klasse || "",
    "{{BETREUER}}": data.betreuer || "",
    "{{BETREUER2}}": data.betreuer2 || "",
    "{{TITEL}}": data.titel || "",
    "{{INHALT}}": data.inhalt || "",
    "{{MOTIVATION}}": data.motivation || "",
    "{{ZIELE}}": data.ziele || "",
    "{{MEILENSTEINE}}": meilensteineText,
    "{{GRUPPENARBEIT}}": data.gruppenarbeit ? `Ja, mit ${data.partner || "–"}` : "Nein",
    "{{AUFGABENTEILUNG}}": data.aufgabenteilung || "–",
    "{{KOSTENPLAN}}": kostenText,
    "{{STATUS}}": data.status || "Entwurf",
    "{{DATUM}}": Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm"),
  };

  Object.keys(werte).forEach((platzhalter) => {
    body.replaceText(platzhalter.replace(/[{}]/g, "\\$&"), werte[platzhalter]);
  });

  if (body.getText().trim() === "") {
    fuegeDeckblattEin(body, "Projektplan", data);
    body.appendParagraph(`Projektplan — ${anzeigeNummFirSchueler(werte["{{SCHUELER}}"])} (${werte["{{KLASSE}}"]})`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Status: ${werte["{{STATUS}}"]}  ·  Stand: ${werte["{{DATUM}}"]}`);
    const betreierAnzeige = [werte["{{BETREUER}}"], werte["{{BETREUER2}}"]].filter((b) => b && b.trim() !== "").join(", ");
    body.appendParagraph(`Betreuer/-in: ${betreierAnzeige || "–"}`);
    body.appendParagraph(`Titel: ${werte["{{TITEL}}"]}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Inhalt / Beschreibung").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{INHALT}}"]);
    body.appendParagraph("Motivation").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{MOTIVATION}}"]);
    body.appendParagraph("Ziele").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{ZIELE}}"]);
    body.appendParagraph("Meilensteine").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{MEILENSTEINE}}"]);
    body.appendParagraph("Gruppenarbeit / Aufgabenteilung").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{GRUPPENARBEIT}}"]);
    body.appendParagraph(werte["{{AUFGABENTEILUNG}}"]);
    body.appendParagraph("Kostenplan (max. 500 €)").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(werte["{{KOSTENPLAN}}"]);
  }
}

function initialisiereSuiviDoc(body, data) {
  fuegeDeckblattEin(body, "Suivi", data);
  body.appendParagraph(`Suivi — ${anzeigeNummFirSchueler(data.schueler)} (${data.klasse || ""})`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Lehrer-Verlaufs- und Bewertungsdokument. Wird von der Betreuer/-in geführt.");
  body.appendParagraph("Kapitel 1: Projektplan").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("(Noch nicht freigegeben.)");
}

function uebernehmeProjektplanInSuivi(data) {
  const ordner = getStudentFolder(data.schueler);
  const { doc, neu } = getOrCreateDoc(ordner, `_quelle_Suivi_${data.schueler}`, false);
  const body = doc.getBody();

  if (neu) {
    initialisiereSuiviDoc(body, data);
  }

  const headingText = "Kapitel 1: Projektplan";
  body.appendParagraph("");
  const startIndex = entferneVorherigenAbschnitt(body, headingText);

  const meilensteineText = (data.meilensteine || [])
    .map((m) => `• ${m.datum || "(kein Datum)"} — ${m.beschreibung}`)
    .join("\n");
  const kostenText = data.kostenplanNoetig
    ? (data.kosten || []).map((k) => `• ${k.bezeichnung}: ${k.betrag} €`).join("\n")
    : "nicht erforderlich";

  let pos = startIndex >= 0 ? startIndex : naechsteEinfuegePosition(body, headingText);

  const einfuegen = (text, heading, bold) => {
    const p = body.insertParagraph(pos, text);
    if (heading) p.setHeading(heading);
    if (bold) p.setBold(true);
    pos++;
  };

  einfuegen(headingText, DocumentApp.ParagraphHeading.HEADING1, false);
  einfuegen(`Freigegeben am: ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm")}`, null, false);
  einfuegen(`Betreuer/-in: ${[data.betreuer, data.betreuer2].filter((b) => b && b.trim() !== "").join(", ") || "–"}`, null, false);
  einfuegen(`Titel: ${data.titel || ""}`, DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen("Inhalt / Beschreibung", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(data.inhalt || "–", null, false);
  einfuegen("Motivation", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(data.motivation || "–", null, false);
  einfuegen("Ziele", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(data.ziele || "–", null, false);
  einfuegen("Meilensteine", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(meilensteineText || "–", null, false);
  einfuegen("Gruppenarbeit / Aufgabenteilung", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(data.gruppenarbeit ? `Ja, mit ${data.partner || "–"}` : "Nein", null, false);
  if (data.gruppenarbeit) einfuegen(data.aufgabenteilung || "–", null, false);
  einfuegen("Kostenplan (max. 500 €)", DocumentApp.ParagraphHeading.HEADING3, false);
  einfuegen(kostenText, null, false);

  if (pos < body.getNumChildren() && body.getChild(pos).getType() !== DocumentApp.ElementType.PAGE_BREAK) {
    body.insertPageBreak(pos);
  }

  return exportiereAlsWord(doc, ordner, `Suivi_${data.schueler}`);
}

const ABSCHNITTS_REIHENFOLGE = [
  "Kapitel 1: Projektplan",
  "Semester 1 – Bewertung & Kommentar",
  "Semester 2 – Bewertung & Kommentar",
  "Trimester 1 – Bewertung & Kommentar",
  "Trimester 2 – Bewertung & Kommentar",
  "Trimester 3 – Bewertung & Kommentar",
];

function naechsteEinfuegePosition(body, headingText) {
  const meinIndex = ABSCHNITTS_REIHENFOLGE.indexOf(headingText);
  const anzahl = body.getNumChildren();
  if (meinIndex !== -1) {
    for (let i = 0; i < anzahl; i++) {
      const el = body.getChild(i);
      if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const p = el.asParagraph();
        if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
          const idx = ABSCHNITTS_REIHENFOLGE.indexOf(p.getText());
          if (idx !== -1 && idx > meinIndex) {
            if (i > 0 && body.getChild(i - 1).getType() === DocumentApp.ElementType.PAGE_BREAK) {
              return i - 1;
            }
            return i;
          }
        }
      }
    }
  }
  return Math.max(0, anzahl - 1);
}

function getAktuellenBewertungsStatus(schueler, periode) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Bewertungen");
  if (!sheet) return null;
  const werte = sheet.getDataRange().getValues();
  const header = werte[0] || [];
  const statusSpalte = header.indexOf("Status");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler && werte[i][1] === periode) {
      return statusSpalte >= 0 ? werte[i][statusSpalte] : null;
    }
  }
  return null;
}

function bewertungEntsperren() {
  const SCHUELER = "Léa Muller";
  const PERIODE = "Semester 2";

  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Bewertungen");
  if (!sheet) {
    Logger.log('Tab "Bewertungen" nicht gefunden.');
    return;
  }
  const werte = sheet.getDataRange().getValues();
  const header = werte[0];
  const statusSpalte = header.indexOf("Status");
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === SCHUELER && werte[i][1] === PERIODE) {
      sheet.getRange(i + 1, statusSpalte + 1).setValue("Entwurf");
      Logger.log("✅ Entspaart: " + SCHUELER + " – " + PERIODE);
      return;
    }
  }
  Logger.log("⚠️ Keng Bewertung fonnt fir: " + SCHUELER + " – " + PERIODE);
}

function getBetreuerFuerSchueler(schueler) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Übersicht");
  if (!sheet) return "";
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) return werte[i][2] || "";
  }
  return "";
}

function schreibeBewertung(data) {
  const bisherigerStatus = getAktuellenBewertungsStatus(data.schueler, data.periode);
  if (bisherigerStatus === "Finalisiert") {
    throw new Error("Diese Bewertung (" + data.periode + ") ist bereits finalisiert und gesperrt.");
  }
  if (!data.betreuer) {
    data.betreuer = getBetreuerFuerSchueler(data.schueler);
  }

  const gesamt = (data.kategorien || []).reduce((s, k) => s + (Number(k.punkte) || 0), 0)
    + (Number(data.wochenberichtePunkte) || 0);
  const gesamtMax = (data.kategorien || []).reduce((s, k) => s + (Number(k.max) || 0), 0)
    + (Number(data.wochenberichteMax) || 0);

  if (!data.finalisiert) {
    aktualisiereBewertungsSheet(data, gesamt, gesamtMax, null, null);
    return null;
  }

  const ordner = getStudentFolder(data.schueler);
  const { doc, neu } = getOrCreateDoc(ordner, `_quelle_Suivi_${data.schueler}`, false);
  const body = doc.getBody();

  if (neu) {
    initialisiereSuiviDoc(body, data);
  }

  const headingText = `${data.periode} – Bewertung & Kommentar`;
  body.appendParagraph("");
  entferneVorherigenAbschnitt(body, headingText);

  let pos = naechsteEinfuegePosition(body, headingText);
  const einfuegenAbsatz = (text) => { const p = body.insertParagraph(pos, text); pos++; return p; };
  const einfuegenTabelle = (daten) => { const t = body.insertTable(pos, daten); pos++; return t; };
  const einfuegenSeitenumbruch = () => { body.insertPageBreak(pos); pos++; };

  einfuegenSeitenumbruch();
  einfuegenAbsatz(headingText).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  einfuegenAbsatz(`Stand: ${Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm")}`);
  einfuegenAbsatz("✅ Validéiert / Finalisiert").setBold(true).setForegroundColor("#1F5C3B");

  const itemsProGruppe = {};
  (data.kategorien || []).forEach((k) => {
    const teile = k.name.split(" — ");
    const gruppe = teile[0];
    const position = teile.slice(1).join(" — ") || k.name;
    if (!itemsProGruppe[gruppe]) itemsProGruppe[gruppe] = [];
    itemsProGruppe[gruppe].push({ name: position, punkte: k.punkte, max: k.max });
  });

  const gruppenReihenfolge = (data.gruppenBemerkungen && data.gruppenBemerkungen.length > 0)
    ? data.gruppenBemerkungen.map((g) => g.titel)
    : Object.keys(itemsProGruppe);

  gruppenReihenfolge.forEach((titel) => {
    const items = itemsProGruppe[titel] || [];
    if (items.length === 0) return;
    const teilSumme = items.reduce((s, i) => s + (Number(i.punkte) || 0), 0);
    const teilMax = items.reduce((s, i) => s + (Number(i.max) || 0), 0);

    einfuegenAbsatz(`${titel}  (Teilnote: ${teilSumme} / ${teilMax})`).setHeading(DocumentApp.ParagraphHeading.HEADING3);

    const tabellenDaten = [["Position", "Punkte"]].concat(
      items.map((i) => [i.name, `${i.punkte} / ${i.max}`])
    );
    const tabelle = einfuegenTabelle(tabellenDaten);
    styleKopfzeile(tabelle);
    styleTabellenzellen(tabelle);

    const gBem = (data.gruppenBemerkungen || []).find((g) => g.titel === titel);
    if (gBem) {
      (gBem.positiv || []).forEach((r) => {
        einfuegenAbsatz("✓ " + r).setForegroundColor("#1F5C3B").setFontSize(10);
      });
      (gBem.negativ || []).forEach((r) => {
        einfuegenAbsatz("✗ " + r).setForegroundColor("#99392F").setFontSize(10);
      });
      if (gBem.kommentar) {
        einfuegenAbsatz(gBem.kommentar).setItalic(true).setFontSize(10);
      }
    }
    einfuegenAbsatz("").setFontSize(4);
  });

  if (data.wochenberichtePunkte !== undefined) {
    einfuegenAbsatz(data.wochenberichteLabel || "Wochenberichte").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const wbTabelle = einfuegenTabelle([
      ["Position", "Punkte"],
      [data.wochenberichteLabel || "Wochenberichte", `${data.wochenberichtePunkte} / ${data.wochenberichteMax}`],
    ]);
    styleKopfzeile(wbTabelle);
    styleTabellenzellen(wbTabelle);
    einfuegenAbsatz("").setFontSize(4);
  }

  const noteAnzeige = data.noteBerechnet !== undefined ? data.noteBerechnet : null;

  einfuegenAbsatz(`Gesamt: ${gesamt} / ${gesamtMax} Punkte${noteAnzeige !== null ? `   ·   Note: ${noteAnzeige} / 60` : ""}`)
    .setBold(true).setFontSize(13);

  einfuegenAbsatz("Kommentar der Lehrperson (Zusammenfassung)").setHeading(DocumentApp.ParagraphHeading.HEADING3);
  (data.allgemeinPositiv || []).forEach((r) => {
    einfuegenAbsatz("✓ " + r).setForegroundColor("#1F5C3B").setFontSize(10);
  });
  (data.allgemeinNegativ || []).forEach((r) => {
    einfuegenAbsatz("✗ " + r).setForegroundColor("#99392F").setFontSize(10);
  });
  einfuegenAbsatz(data.kommentar || "–");

  const suiviUrl = exportiereAlsWord(doc, ordner, `Suivi_${data.schueler}`);
  const suiviPdfUrl = exportiereAlsPdf(doc, ordner, `Suivi_${data.schueler}`);

  aktualisiereBewertungsSheet(data, gesamt, gesamtMax, suiviUrl, suiviPdfUrl);
  aktualisiereUebersicht({ schueler: data.schueler }, null, suiviUrl, suiviPdfUrl);

  return suiviUrl;
}

function styleKopfzeile(tabelle) {
  const kopf = tabelle.getRow(0);
  for (let c = 0; c < kopf.getNumCells(); c++) {
    const zelle = kopf.getCell(c);
    zelle.setBackgroundColor("#EAF2EC");
    zelle.editAsText().setBold(true).setFontSize(10);
  }
}

function styleTabellenzellen(tabelle) {
  for (let r = 1; r < tabelle.getNumRows(); r++) {
    const zeile = tabelle.getRow(r);
    for (let c = 0; c < zeile.getNumCells(); c++) {
      zeile.getCell(c).editAsText().setFontSize(10).setBold(false);
    }
  }
}

function entferneVorherigenAbschnitt(body, headingText) {
  const anzahl = body.getNumChildren();
  let startIndex = -1;
  let endIndex = -1;

  for (let i = 0; i < anzahl; i++) {
    const el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = el.asParagraph();
      if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING1 && p.getText() === headingText) {
        startIndex = i;
        continue;
      }
      if (startIndex > -1 && p.getHeading() === DocumentApp.ParagraphHeading.HEADING1 && p.getText() !== headingText) {
        endIndex = i;
        break;
      }
    }
  }
  if (startIndex === -1) return -1;
  if (endIndex === -1) endIndex = anzahl - 1;

  let entfernterPageBreak = false;
  if (startIndex > 0 && body.getChild(startIndex - 1).getType() === DocumentApp.ElementType.PAGE_BREAK) {
    entfernterPageBreak = true;
  }

  for (let i = endIndex - 1; i >= startIndex; i--) {
    body.removeChild(body.getChild(i));
  }
  if (entfernterPageBreak) {
    body.removeChild(body.getChild(startIndex - 1));
    return startIndex - 1;
  }
  return startIndex;
}

function aktualisiereUebersicht(data, projektplanUrl, suiviUrl, suiviPdfUrl) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const sollHeader = ["Schüler", "Klasse", "Betreuer", "Titel", "Status", "Zuletzt aktualisiert", "Projektplan-Link", "Suivi-Link", "Suivi-PDF-Link (Schüler)", "Projektplan-Details (JSON)", "Dokumentatioun-Link"];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(sollHeader);
  } else {
    const istHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let abweichend = istHeader.length < sollHeader.length;
    if (!abweichend) {
      for (let i = 0; i < sollHeader.length; i++) {
        if (istHeader[i] !== sollHeader[i]) { abweichend = true; break; }
      }
    }
    if (abweichend) {
      sheet.getRange(1, 1, 1, sollHeader.length).setValues([sollHeader]);
    }
  }

  const werte = sheet.getDataRange().getValues();
  let zeile = -1;
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.schueler) { zeile = i + 1; break; }
  }

  const bestehend = zeile > 0 ? werte[zeile - 1] : null;

  const projektplanDetails = data.inhalt !== undefined
    ? JSON.stringify({
        inhalt: data.inhalt || "",
        motivation: data.motivation || "",
        ziele: data.ziele || "",
        meilensteine: data.meilensteine || [],
        gruppenarbeit: !!data.gruppenarbeit,
        partner: data.partner || "",
        aufgabenteilung: data.aufgabenteilung || "",
        kostenplanNoetig: !!data.kostenplanNoetig,
        kosten: data.kosten || [],
      })
    : (bestehend ? bestehend[9] : "");

  const betreuerKombiniert = (data.betreuer || data.betreuer2)
    ? [data.betreuer, data.betreuer2].filter((b) => b && b.trim() !== "").join(", ")
    : (bestehend ? bestehend[2] : "");

  const neueZeile = [
    data.schueler || "",
    data.klasse || (bestehend ? bestehend[1] : ""),
    betreuerKombiniert,
    data.titel || (bestehend ? bestehend[3] : ""),
    data.status || (bestehend ? bestehend[4] : "Entwurf"),
    Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm"),
    projektplanUrl || (bestehend ? bestehend[6] : ""),
    suiviUrl || (bestehend ? bestehend[7] : ""),
    suiviPdfUrl || (bestehend ? bestehend[8] : ""),
    projektplanDetails,
    bestehend ? (bestehend[10] || "") : "",
  ];

  if (zeile > 0) {
    sheet.getRange(zeile, 1, 1, neueZeile.length).setValues([neueZeile]);
  } else {
    sheet.appendRow(neueZeile);
  }
}

function speichereDokumentatiounLink(schueler, klasse, link) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const sollHeader = ["Schüler", "Klasse", "Betreuer", "Titel", "Status", "Zuletzt aktualisiert", "Projektplan-Link", "Suivi-Link", "Suivi-PDF-Link (Schüler)", "Projektplan-Details (JSON)", "Dokumentatioun-Link"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(sollHeader);
  }
  const werte = sheet.getDataRange().getValues();
  let zeile = -1;
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) { zeile = i + 1; break; }
  }
  if (zeile > 0) {
    sheet.getRange(zeile, 11).setValue(link);
  } else {
    const neiZeil = [schueler, klasse || "", "", "", "Entwurf", Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm"), "", "", "", "", link];
    sheet.appendRow(neiZeil);
  }
  return { ok: true };
}

function aktualisiereBewertungsSheet(data, gesamt, gesamtMax, docUrl, pdfUrl) {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Bewertungen");
  const sollHeader = ["Schüler", "Periode", "Punkte", "Max", "Note", "Kommentar", "Zuletzt aktualisiert", "Suivi-Link", "Suivi-PDF-Link (Schüler)", "Kategorien (JSON)", "Gruppenbemerkungen (JSON)", "Allgemeine Remarquen (JSON)", "Status"];
  if (!sheet) {
    sheet = ss.insertSheet("Bewertungen");
    sheet.appendRow(sollHeader);
  } else {
    const istHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let abweichend = istHeader.length < sollHeader.length;
    if (!abweichend) {
      for (let i = 0; i < sollHeader.length; i++) {
        if (istHeader[i] !== sollHeader[i]) { abweichend = true; break; }
      }
    }
    if (abweichend) {
      sheet.getRange(1, 1, 1, sollHeader.length).setValues([sollHeader]);
    }
  }

  const werte = sheet.getDataRange().getValues();
  let zeile = -1;
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === data.schueler && werte[i][1] === data.periode) { zeile = i + 1; break; }
  }
  const bestehend = zeile > 0 ? werte[zeile - 1] : null;

  const finalDocUrl = (docUrl !== null && docUrl !== undefined) ? docUrl : (bestehend ? bestehend[7] : "");
  const finalPdfUrl = (pdfUrl !== null && pdfUrl !== undefined) ? pdfUrl : (bestehend ? bestehend[8] : "");

  const note = data.noteBerechnet !== undefined
    ? data.noteBerechnet
    : (gesamtMax > 0 ? Math.round((gesamt / gesamtMax) * 4 * 10) / 10 : "");

  const neueZeile = [
    data.schueler || "",
    data.periode || "",
    gesamt,
    gesamtMax,
    note,
    data.kommentar || "",
    Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm"),
    finalDocUrl,
    finalPdfUrl,
    JSON.stringify(data.kategorien || []),
    JSON.stringify(data.gruppenBemerkungen || []),
    JSON.stringify({ positiv: data.allgemeinPositiv || [], negativ: data.allgemeinNegativ || [] }),
    data.finalisiert ? "Finalisiert" : "Entwurf",
  ];

  if (zeile > 0) {
    sheet.getRange(zeile, 1, 1, neueZeile.length).setValues([neueZeile]);
  } else {
    sheet.appendRow(neueZeile);
  }

  setzeStatusDropdown(sheet, sollHeader.indexOf("Status") + 1);
}

function setzeStatusDropdown(sheet, statusSpaltenNr) {
  const regel = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Entwurf", "Finalisiert"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusSpaltenNr, 500, 1).setDataValidation(regel);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function alleFormulardatenLoeschen() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  const sheets = [ss.getSheets()[0], ss.getSheetByName("Bewertungen")];
  sheets.forEach((sheet) => {
    if (!sheet) return;
    const letzteZeile = sheet.getLastRow();
    if (letzteZeile > 1) {
      sheet.deleteRows(2, letzteZeile - 1);
      Logger.log(`Tab "${sheet.getName()}": ${letzteZeile - 1} Zeile(n) gelöscht.`);
    } else {
      Logger.log(`Tab "${sheet.getName()}": bereits leer.`);
    }
  });
  Logger.log("✅ Alle Formulardaten gelöscht.");
}

function raeumeVerwaisteEintraegeAuf() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let gesamtGeloescht = 0;

  const sheets = [ss.getSheets()[0], ss.getSheetByName("Bewertungen")];
  sheets.forEach((sheet) => {
    if (!sheet) {
      Logger.log("Tab nicht gefunden, übersprungen.");
      return;
    }
    const tabName = sheet.getName();
    const werte = sheet.getDataRange().getValues();
    if (werte.length <= 1) {
      Logger.log(`Tab "${tabName}": keine Datenzeilen vorhanden.`);
      return;
    }
    let geloeschtHier = 0;
    for (let i = werte.length - 1; i >= 1; i--) {
      const schueler = werte[i][0];
      if (!schueler) continue;
      const existiertNoch = pruefeObSchuelerOrdnerExistiert(schueler);
      if (!existiertNoch) {
        sheet.deleteRow(i + 1);
        geloeschtHier++;
      }
    }
    Logger.log(`Tab "${tabName}": ${geloeschtHier} verwaiste Zeile(n) entfernt.`);
    gesamtGeloescht += geloeschtHier;
  });

  Logger.log(`✅ Fertig. Insgesamt ${gesamtGeloescht} Zeile(n) entfernt.`);
}

function pruefeObSchuelerOrdnerExistiert(schuelerName) {
  const hauptordner = DriveApp.getFolderById(FOLDER_ID);
  const ordner = hauptordner.getFoldersByName(schuelerName);
  return ordner.hasNext();
}

function statusDropdownEinmaligEinrichten() {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheetByName("Bewertungen");
  if (!sheet) {
    Logger.log('Tab "Bewertungen" nicht gefunden.');
    return;
  }
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusSpalte = header.indexOf("Status") + 1;
  if (statusSpalte === 0) {
    Logger.log('Spalte "Status" nicht gefunden.');
    return;
  }
  setzeStatusDropdown(sheet, statusSpalte);
  Logger.log("✅ Dropdown eingerichtet.");
}

function testDeckblatt() {
  const blob = erstelleDeckblattBildViaPDFShift("Suivi", {
    schueler: "Léa Muller",
    klasse: "2GSE",
    betreuer: "Prof. Kayser",
    betreuer2: "Prof. Reuter",
  });
  const datei = DriveApp.getFolderById(FOLDER_ID).createFile(blob);
  Logger.log("Testbild erstellt: " + datei.getUrl());
}

function testLauf() {
  const beispiel = {
    schueler: "Testschüler XY",
    klasse: "1GSE",
    betreuer: "Prof. Kayser",
    titel: "Wasserqualität der Attert",
    inhalt: "Untersuchung der Wasserqualität an drei Messpunkten entlang der Attert.",
    motivation: "Interesse an Gewässerökologie und praktischer Feldarbeit.",
    ziele: "Bis Ende Oktober: Messpunkte definiert, erste Referenzmessung durchgeführt.",
    meilensteine: [
      { datum: "2026-10-07", beschreibung: "Projektplan abgegeben" },
      { datum: "2026-11-15", beschreibung: "Erste Messreihe abgeschlossen" },
    ],
    gruppenarbeit: false,
    kostenplanNoetig: true,
    kosten: [{ bezeichnung: "pH-Messgerät", betrag: 45 }],
    status: "Freigegeben & digital unterschrieben",
  };
  Logger.log(erstelleOderAktualisiereProjektplan(beispiel));
}

const SESSION_GUELTEGKEET_STONNEN = 18;

function getLoginSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Login");
  if (!sheet) {
    sheet = ss.insertSheet("Login");
    sheet.appendRow(["LoginId", "Rolle", "Klasse", "PIN-Hash", "Salt"]);
  }
  return sheet;
}

function getSessionsSheet() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("Sessions");
  if (!sheet) {
    sheet = ss.insertSheet("Sessions");
    sheet.appendRow(["Token", "LoginId", "Rolle", "Klasse", "Erstallt"]);
  }
  return sheet;
}

function hashPin(pin, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin + ":" + salt);
  return digest.map((b) => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

const FEST_PINS = {
  "Guy Putz": "2803",
  "Pol Medernach": "4127",
  "Sarah Blum": "5936",
  "Tania Ludwig": "7451",
  "Tom Bleyer": "3298",
  "Salman Murad": "6614",
  "Alex Olinger": "1007",
};

function initialiséierLoginPins() {
  const sheet = getLoginSheet();
  const werte = sheet.getDataRange().getValues();
  const bestehendNimm = new Set(werte.slice(1).map((z) => z[0]));
  const nei = [];

  Object.keys(typeof LEHRER_EMAILS !== "undefined" ? LEHRER_EMAILS : {}).forEach((numm) => {
    if (bestehendNimm.has(numm)) return;
    const pin = FEST_PINS[numm] || String(Math.floor(1000 + Math.random() * 9000));
    const salt = Utilities.getUuid();
    sheet.appendRow([numm, "Prof", "", hashPin(pin, salt), salt]);
    nei.push({ numm, rolle: "Prof", pin });
  });
  Logger.log(JSON.stringify(nei, null, 2));
  return nei;
}

function login(loginIdEingabe, pin) {
  if (!loginIdEingabe || !pin) return { ok: false, error: "Umeldung a Passwuert erfuerderlech." };
  const eingabe = String(loginIdEingabe).trim();
  const eingabeLower = eingabe.toLowerCase();

  const personenSheet = getPersonenSheet();
  const pWerte = personenSheet.getDataRange().getValues();
  const persoonZeilIndex = pWerte.slice(1).findIndex((z) => String(z[0] || "").trim().toLowerCase() === eingabeLower);
  const persoonZeil = persoonZeilIndex !== -1 ? pWerte[persoonZeilIndex + 1] : null;
  if (persoonZeil && persoonZeil[6] === "Nee") {
    return { ok: false, error: "Dëse Zougang ass deaktivéiert. Frot de Prof." };
  }

  const sheet = getLoginSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    const [loginId, rolle, klasse, pinHash, salt] = werte[i];
    if (String(loginId || "").trim().toLowerCase() === eingabeLower) {
      if (hashPin(String(pin), salt) !== pinHash) {
        return { ok: false, error: "Falscht Passwuert." };
      }
      const token = Utilities.getUuid();
      const jetzt = new Date().toISOString();
      getSessionsSheet().appendRow([token, loginId, rolle, klasse, jetzt]);
      if (persoonZeilIndex !== -1) {
        const jetztLiesbar = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
        personenSheet.getRange(persoonZeilIndex + 2, 8).setValue(jetztLiesbar);
      }
      return { ok: true, token, numm: loginId, rolle, klasse };
    }
  }
  return { ok: false, error: "Onbekannt Umeldung. Frot de Prof no Ärer Matrikelnummer/Ärem Zougang." };
}

function pruefSession(token) {
  if (!token) return { valid: false };
  const sheet = getSessionsSheet();
  const werte = sheet.getDataRange().getValues();
  const jetzt = new Date();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === token) {
      const erstallt = new Date(werte[i][4]);
      const stonnen = (jetzt - erstallt) / (1000 * 60 * 60);
      if (stonnen > SESSION_GUELTEGKEET_STONNEN) {
        sheet.deleteRow(i + 1);
        return { valid: false, error: "Session ofgelaf, weg nei umellen." };
      }
      return { valid: true, numm: werte[i][1], rolle: werte[i][2], klasse: werte[i][3] };
    }
  }
  return { valid: false };
}

function logout(token) {
  if (!token) return;
  const sheet = getSessionsSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === token) { sheet.deleteRow(i + 1); return; }
  }
}

function pinZuruecksetzen(numm, proffToken, neiesPasswuert) {
  const session = pruefSession(proffToken);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen PINs/Passwierder änneren." };
  }
  const gesichtNumm = String(numm || "").trim();
  const sheet = getLoginSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (String(werte[i][0] || "").trim() === gesichtNumm) {
      const gewenschtPasswuert = (neiesPasswuert || "").trim();
      const pin = gewenschtPasswuert.length >= 6
        ? gewenschtPasswuert
        : String(Math.floor(1000 + Math.random() * 9000));
      const salt = Utilities.getUuid();
      sheet.getRange(i + 1, 4, 1, 2).setValues([[hashPin(pin, salt), salt]]);
      return { ok: true, numm: gesichtNumm, pin };
    }
  }
  return { ok: false, error: "Numm net fonnt an der Login-Tabell. Probéiert d'Persoun iwwer 'Nei Persoun' nach eng Kéier ze späicheren (dat erstellt eng Login-Zeil, falls se feelt)." };
}

function seedRichtegSchuelerSept2026() {
  const SCHUELER = [
    { virnumm: "Ben", klasse: "1GSE", pin: "4172" },
    { virnumm: "Test", klasse: "1GSE", pin: "1508" },
    { virnumm: "Max", klasse: "2GSE", pin: "8035" },
    { virnumm: "Ivan", klasse: "2GSE", pin: "2947" },
    { virnumm: "Mergon", klasse: "2GSE", pin: "6183" },
    { virnumm: "Edson", klasse: "2GSE", pin: "5290" },
    { virnumm: "Guillaume", klasse: "2GSE", pin: "7364" },
  ];

  const personenSheet = getPersonenSheet();
  const loginSheet = getLoginSheet();

  [personenSheet, loginSheet].forEach((sheet) => {
    const werte = sheet.getDataRange().getValues();
    const rolleSpalt = sheet === personenSheet ? 3 : 1;
    for (let i = werte.length - 1; i >= 1; i--) {
      if (werte[i][rolleSpalt] === "Schüler") sheet.deleteRow(i + 1);
    }
  });

  const resultat = [];
  SCHUELER.forEach((s) => {
    const matrikel = generéierMatrikel(s.klasse);
    personenSheet.appendRow([matrikel, s.virnumm, "", "Schüler", s.klasse, "", "Jo"]);
    const salt = Utilities.getUuid();
    loginSheet.appendRow([matrikel, "Schüler", s.klasse, hashPin(s.pin, salt), salt]);
    resultat.push({ matrikel, virnumm: s.virnumm, klasse: s.klasse, pin: s.pin });
  });
  Logger.log("✅ " + resultat.length + " Schüler ugeluecht:\n" + JSON.stringify(resultat, null, 2));
  return resultat;
}

/**
 * Eemoleg Opraum-Funktioun: virun der Matrikel-Ëmstellung hat d'Personen-
 * Sheet e falscht Format (6 Kolonnen: Numm, Rolle, Klasse, Email, Aktiv,
 * Untis-Code). Duerno gouf op 7 Kolonnen ëmgestallt (LoginId, Virnumm,
 * Numm, Rolle, Klasse, Email, Aktiv), mä al Zeilen goufen NET ëmgeschriwwen.
 * Dës Funktioun fënnt esou al Zeilen (erkannt drun, datt Kolonn D net
 * "Schüler" oder "Prof" ass) a behandelt se sou:
 *  - Al Prof-Zeil (Roll stoung a Kolonn B): gëtt an d'nei Format
 *    ëmgeschriwwen, LoginId/Numm/Email bleiwen erhalen (Login-Zougang
 *    onberuehrt).
 *  - Al Test-Schüler-Zeil: gëtt komplett geläscht (Numm ass duerch déi
 *    richteg Matrikel-Schüler ersat), zesumme mat der entspriechender
 *    Login-Zeil.
 * Just eemol lafen, duerno kann dës Funktioun ignoréiert ginn.
 */
function raeumAlFormatPersonenOp() {
  const personenSheet = getPersonenSheet();
  const loginSheet = getLoginSheet();
  const werte = personenSheet.getDataRange().getValues();
  const geläschtNimm = [];
  const ëmgeschriwwen = [];

  for (let i = werte.length - 1; i >= 1; i--) {
    const z = werte[i];
    const roleSpaltD = z[3]; // nei Positioun fir "Rolle"
    if (roleSpaltD === "Schüler" || roleSpaltD === "Prof") continue; // scho nei Format

    // Al Format: [Numm, Rolle, Klasse, Email, Aktiv, UntisCode]
    const alNumm = z[0];
    const alRolle = z[1];
    const alEmail = z[3];
    const alAktiv = z[4];

    if (alRolle === "Prof") {
      personenSheet.getRange(i + 1, 1, 1, 7).setValues([[alNumm, "", alNumm, "Prof", "", alEmail || "", alAktiv === "Nee" ? "Nee" : "Jo"]]);
      ëmgeschriwwen.push(alNumm);
    } else if (alRolle === "Schüler") {
      personenSheet.deleteRow(i + 1);
      geläschtNimm.push(alNumm);
    } else {
      Logger.log("⚠️ Onbekannt al Zeil (net ëmgeschriwwen, net geläscht): " + JSON.stringify(z));
    }
  }

  if (geläschtNimm.length > 0) {
    const loginWerte = loginSheet.getDataRange().getValues();
    for (let i = loginWerte.length - 1; i >= 1; i--) {
      if (geläschtNimm.includes(loginWerte[i][0])) loginSheet.deleteRow(i + 1);
    }
  }

  Logger.log("✅ Opraum fäerdeg. Ëmgeschriwwen (Prof): " + JSON.stringify(ëmgeschriwwen) + ". Geläscht (al Test-Schüler): " + JSON.stringify(geläschtNimm));
  return { ëmgeschriwwen, geläschtNimm };
}
