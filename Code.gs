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
 *         Projektplan_<Schüler>.pdf      – PDF-Version (Deckblat + Inhalt)
 *         _quelle_Suivi_<Schüler>        – internes Google Doc (Arbeitskopie)
 *         Suivi_<Schüler>.docx           – ECHTE Word-Datei (Lehrer-Ansicht)
 *         Suivi_<Schüler>.pdf            – PDF-Version (Deckblat + Inhalt)
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
  // Ouni Login ofrufbar (fir de Login-Formulaire selwer, ier eng Session
  // besteet) — gëtt bewosst NËMME Proffen-Nimm eraus, keng Klassen/Matrikelen/
  // Schüler. Prof-Nimm sinn net esou sensibel wéi déi voll Schüler-Lëscht
  // (déi ?namen=1 hei drënner gëtt, déi Login erfuerdert). Zweck: en
  // zouverléissegen Dropdown beim Umellen, fir Tippfehler/Schreifweis-
  // Ënnerscheeder (z.B. "Guy Putz" vs. "Guy PUTZ") auszeschléissen.
  if (e.parameter && e.parameter.proffen === "1") {
    const proffen = getAktivePersonen()
      .filter((p) => p.rolle === "Prof")
      .map((p) => p.numm);
    return jsonResponse({ ok: true, proffen });
  }

  const session = pruefSession(e.parameter && e.parameter.token);

  if (e.parameter && e.parameter.namen === "1") {
    if (!session.valid) {
      return jsonResponse({ ok: false, error: "Net ugemellt." });
    }
    return jsonResponse({ ok: true, personen: getAktivePersonen() });
  }

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
  let data;
  try {
    data = JSON.parse(e.postData.contents);
    let url;
    let pdfUrl;
    if (data.typ === "bewertung") {
      url = schreibeBewertung(data);
    } else if (data.typ === "statusAendern") {
      aendereBewertungsStatus(data);
      return jsonResponse({ ok: true });
    } else if (data.typ === "meilensteng") {
      speichereMeilensteng(data);
      return jsonResponse({ ok: true });
    } else if (data.typ === "offiziellZaitplang") {
      speichereOffiziellZaitplang(data);
      return jsonResponse({ ok: true });
    } else if (data.typ === "wochenbericht") {
      const id = speichereWochenbericht(data);
      return jsonResponse({ ok: true, id });
    } else if (data.typ === "wochenberichtBewerten") {
      bewerteWochenbericht(data);
      return jsonResponse({ ok: true });
    } else if (data.typ === "wochenberichtEntschellegen") {
      entschellegWochenbericht(data);
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
      return jsonResponse(speichereDokumentatiounLink(data));
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
      const resultat = erstelleOderAktualisiereProjektplan(data);
      url = resultat.url;
      pdfUrl = resultat.pdfUrl;
    }
    return jsonResponse({ ok: true, url, pdfUrl });
  } catch (err) {
    Logger.log("doPost Feeler: " + err.message + "\nStack: " + err.stack);
    try {
      const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
      let debugSheet = ss.getSheetByName("Debug");
      if (!debugSheet) {
        debugSheet = ss.insertSheet("Debug");
        debugSheet.appendRow(["Zäit", "Typ", "Feeler", "Stack"]);
      }
      debugSheet.appendRow([
        Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm:ss"),
        (data && data.typ) || "",
        err.message,
        err.stack || "",
      ]);
    } catch (debugErr) { }
    return jsonResponse({ ok: false, error: err.message });
  }
}

function aendereBewertungsStatus(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    throw new Error("Nëmme Proffen dierfen de Bewertungsstatus änneren.");
  }
  const { schueler, periode, neuerStatus } = data;
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

function speichereMeilensteng(data) {
  const session = pruefSession(data.token);
  if (!session.valid || (session.rolle !== "Prof" && session.numm !== data.schueler)) {
    throw new Error("Net erlaabt.");
  }
  const { schueler, klasse, meilensteng } = data;
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

function speichereOffiziellZaitplangIntern(klasse, meilensteng) {
  const sheet = getOffiziellZaitplangSheet();
  const werte = sheet.getDataRange().getValues();
  for (let i = werte.length - 1; i >= 1; i--) {
    if (werte[i][0] === klasse) sheet.deleteRow(i + 1);
  }
  meilensteng.forEach((m) => {
    sheet.appendRow([klasse, m.datum, m.titel, m.kategorie || "Event"]);
  });
}

/** Ëffentlech (HTTP-)Variant mat Session/Roll-Kontroll — nëmmen Proffen. */
function speichereOffiziellZaitplang(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    throw new Error("Nëmme Proffen dierfen den offizielle Zäitplang änneren.");
  }
  speichereOffiziellZaitplangIntern(data.klasse, data.meilensteng);
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
  speichereOffiziellZaitplangIntern("1GSE", EIN_1GSE);
  speichereOffiziellZaitplangIntern("2GSE", EIN_2GSE);
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
          "Hei korrigéieren: https://pugu-prog.github.io/ppren/wochenberichte-korrigeieren.html",
      });
    } catch (e) { }
  });
}

function speichereWochenbericht(data) {
  const session = pruefSession(data.token);
  if (!session.valid || (session.rolle !== "Prof" && session.numm !== data.schueler)) {
    throw new Error("Net erlaabt.");
  }
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

function bewerteWochenbericht(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    throw new Error("Nëmme Proffen dierfen Wochenberichter bewäerten.");
  }
  const { id, punkte } = data;
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

function entschellegWochenbericht(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    throw new Error("Nëmme Proffen dierfen e Wochenbericht entschëllegen.");
  }
  const { id, grond } = data;
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
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    throw new Error("Nëmme Proffen dierfen e Fachgespréich bewäerten.");
  }
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
      "ID", "Schüler", "Klasse", "Termin", "Datum",
      "Punkte", "Max", "Notiz", "Zuletzt aktualisiert",
    ]);
  }
  return sheet;
}
