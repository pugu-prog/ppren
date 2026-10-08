// =====================================================================
// PlangKiFeedback.gs — KI-Entworf fir de Projektplang-Feedback (Prof-Usicht)
//
// Installatioun am Apps-Script-Projet "PPREN web":
//   1. Nei Datei "PlangKiFeedback.gs" uleeën an dësen Inhalt drakopéieren.
//   2. Am doPost-Router vu Code.gs eng Zeil derbäisetzen:
//        } else if (data.typ === "plangKiFeedback") {
//          return jsonResponse(plangKiFeedback(data));
//   3. Script Property ANTHROPIC_API_KEY muss gesat sinn (gëtt och vum Feedback-Check benotzt).
//   4. Deploy → Manage deployments → ✏️ Edit → New version.
//
// De Plang gëtt aus der Iwwersiicht-Sheet gelies (net vum Browser geschéckt).
// All KI-Entworf gëtt am Sheet-Tab "PlangKiFeedback" gespäichert (mat Datum,
// Prof a Plang-Stand) — sou bleift d'Entwécklung vum Schüler nokucken, an en
// zweeten Opruff kascht näischt. Mat data.nurLaden = true gëtt just d'Historique
// zréckginn (keen API-Opruff). Wat un de Schüler geet, wielt de Prof am Dashboard.
// =====================================================================

const KI_FEEDBACK_MODELL = "claude-opus-5-5";

function plangKiFeedback(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen de KI-Feedback benotzen." };
  }
  if (data.nurLaden) {
    return { ok: true, historique: kiFeedbackHistorique_(data.schueler) };
  }
  if (data.pdf) {
    return feedbackPdfErstellen_(data);
  }
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    return { ok: false, error: "KI nach net ageriicht (ANTHROPIC_API_KEY feelt an de Script Properties)." };
  }

  const plang = kiPlangLiesen_(data.schueler);
  if (!plang) return { ok: false, error: "Kee Projektplang fonnt fir " + data.schueler + "." };

  const ass1GSE = /1GSE/i.test(plang.klasse);
  const payload = {
    model: KI_FEEDBACK_MODELL,
    max_tokens: 16000,
    thinking: { type: "adaptive" },
    output_config: {
      effort: "medium",
      format: { type: "json_schema", schema: KI_FEEDBACK_SCHEMA },
    },
    fallbacks: "default",
    system: kiSystemPrompt_(ass1GSE),
    messages: [{ role: "user", content: kiPlangAlsText_(plang) }],
  };

  const res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "server-side-fallback-2026-07-01",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  let antwort;
  try {
    antwort = JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, error: "KI-Äntwert konnt net gelies ginn (HTTP " + code + ")." };
  }
  if (code !== 200) {
    const msg = antwort && antwort.error ? antwort.error.message : "HTTP " + code;
    return { ok: false, error: "KI-Feeler: " + msg };
  }
  if (antwort.stop_reason === "refusal") {
    return { ok: false, error: "D'KI huet d'Ufro ofgeleent." };
  }
  if (antwort.stop_reason === "max_tokens") {
    return { ok: false, error: "D'KI-Äntwert gouf ofgeschnidden. Nach eng Kéier probéieren." };
  }
  const textBlock = (antwort.content || []).find((b) => b.type === "text");
  if (!textBlock) return { ok: false, error: "D'KI huet keen Text geliwwert." };

  let resultat;
  try {
    resultat = JSON.parse(textBlock.text);
  } catch (e) {
    return { ok: false, error: "D'KI-Äntwert war kee gültegen JSON." };
  }
  const erstallt = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy HH:mm");
  const id = Utilities.getUuid();
  kiFeedbackSheet_().appendRow([id, plang.schueler, plang.klasse, erstallt, session.numm || "", plang.stand, antwort.model || KI_FEEDBACK_MODELL, JSON.stringify(resultat)]);
  return { ok: true, resultat, modell: antwort.model, eintrag: { id, erstallt, vunProf: session.numm || "", planStand: plang.stand, resultat } };
}

function kiFeedbackSheet_() {
  const ss = SpreadsheetApp.openById(OVERVIEW_SHEET_ID);
  let sheet = ss.getSheetByName("PlangKiFeedback");
  if (!sheet) {
    sheet = ss.insertSheet("PlangKiFeedback");
    sheet.appendRow(["ID", "Schüler", "Klasse", "Erstallt", "VunProf", "Plang-Stand", "Modell", "Resultat (JSON)"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// All gespäichert KI-Entwërf vun engem Schüler, nei → al.
function kiFeedbackHistorique_(schueler) {
  const werte = kiFeedbackSheet_().getDataRange().getValues();
  const lescht = [];
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][1] !== schueler) continue;
    let resultat = null;
    try { resultat = JSON.parse(werte[i][7] || "null"); } catch (e) { resultat = null; }
    if (!resultat) continue;
    const zeit = (v) => (v instanceof Date ? Utilities.formatDate(v, "Europe/Luxembourg", "dd.MM.yyyy HH:mm") : String(v || ""));
    lescht.push({ id: werte[i][0], erstallt: zeit(werte[i][3]), vunProf: werte[i][4], planStand: zeit(werte[i][5]), resultat });
  }
  return lescht.reverse();
}

function kiPlangLiesen_(schueler) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) {
      let details = {};
      try { details = JSON.parse(werte[i][9] || "{}"); } catch (e) { details = {}; }
      const stand = werte[i][5] instanceof Date
        ? Utilities.formatDate(werte[i][5], "Europe/Luxembourg", "dd.MM.yyyy HH:mm")
        : String(werte[i][5] || "");
      return {
        schueler: werte[i][0],
        klasse: String(werte[i][1] || ""),
        titel: String(werte[i][3] || ""),
        status: String(werte[i][4] || ""),
        stand,
        details,
      };
    }
  }
  return null;
}

function kiPlangAlsText_(plang) {
  const d = plang.details || {};
  const ms = (d.meilensteine || [])
    .map((m) => "- " + (m.datum || "(ouni Datum)") + " — " + (m.beschreibung || m.titel || ""))
    .join("\n") || "(keng)";
  const kosten = d.kostenplanNoetig
    ? ((d.kosten || []).map((k) => "- " + (k.bezeichnung || "") + ": " + (k.betrag || "0") + " €").join("\n") || "(keng Positiounen)")
    : "nicht erforderlich";
  const heit = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy");
  return [
    "Heutiges Datum: " + heit,
    "Klasse: " + plang.klasse,
    "Status: " + plang.status,
    "",
    "<projektplan>",
    "Titel: " + plang.titel,
    "GSE-Themen (vom Schüler gewählt): " + ((d.gseThemen || []).join(", ") || "(keine)"),
    "",
    "Inhalt / Beschreibung:\n" + (d.inhalt || "(leer)"),
    "",
    "Motivation:\n" + (d.motivation || "(leer)"),
    "",
    "Ziele:\n" + (d.ziele || "(leer)"),
    "",
    "Meilensteine:\n" + ms,
    "",
    "Gruppenarbeit: " + (d.gruppenarbeit ? "Ja, mit " + (d.partner || "?") : "Nein"),
    "Aufgaben / Aufgabenteilung:\n" + (d.aufgabenteilung || "(leer)"),
    "",
    "Kostenplan (max. 500 €):\n" + kosten,
    "</projektplan>",
  ].join("\n");
}

function kiSystemPrompt_(ass1GSE) {
  const raster = ass1GSE
    ? "Klasse 1GSE (Semester-System). Projektplan /30: Beschreibung 15 P, Motivation mit Bezug zu den Umweltwissenschaften 3 P, Ziele 6 P, Aufgaben 6 P, Kostenplan falls erforderlich. Zusätzlich: Meilensteine bis Ende 1. Semester mit klar formuliertem Soll-Ist-Datum. Im 1. Semester wird Zitieren nach APA bewertet, im 2. Semester ein Interview/eine Umfrage und ein Film-Storyboard."
    : "Klasse 2GSE (Trimester-System). Projektplan /30: Beschreibung 15 P, Motivation mit Bezug zu den Umweltwissenschaften 3 P, Ziele 6 P, Aufgaben 6 P, Kostenplan falls erforderlich.";
  return [
    "Du unterstützt die Lehrer der Sektion Sciences Environnementales (GSE) am Lycée Technique d'Ettelbruck (Luxemburg) beim Feedback zu Projektplänen im PPREN (Projet personnel encadré). Die Schüler sind 16–19 Jahre alt. Die Lehrer entscheiden, welche deiner Vorschläge der Schüler als Kommentar bekommt.",
    "",
    "Bewertungsraster: " + raster,
    "",
    "Prüfe den Projektplan in dieser Reihenfolge:",
    "1. Felder: leere Felder; „Ziele“ enthält nur Termine oder eine Aufgabenliste statt Zielen; Meilensteine fehlen, sind vage, nicht chronologisch, nur Theorie oder ohne praktischen Teil; „Kostenplan nicht erforderlich“, obwohl Material gebraucht wird; Positionen mit 0 €. Das Feld „Aufgaben“ ist fast immer leer und bringt 6 Punkte.",
    "2. Bezug zu den sechs GSE-Themen (Energie, Luft, Wasser, Boden, Abfall, Mobilität) – er muss in Motivation UND Zielen stehen. Fehlt er, schlage passende Bezüge für genau dieses Projekt vor.",
    "3. Übergeordnetes Ziel: Der Bau eines Geräts ist das Mittel, nicht der Zweck. Wo wird das Ergebnis eingesetzt, wer (z. B. Kinder, Mitschüler) lernt was daraus, wie wird es vermittelt?",
    "4. Fachliche Fehler und Machbarkeit: falsche Physik/Chemie/Biologie, zu großer Umfang, unrealistische Organismen oder Messungen, technische Stolpersteine. Wenn eine kurze Überschlagsrechnung dem Schüler die Größenordnung zeigt, formuliere sie als Aufgabe mit Richtwerten.",
    "5. Sprache: konkrete Fehler mit Verbesserung (Rechtschreibung, Grammatik, Fachbegriffe).",
    "6. KI-Indizien: sehr glatte, fehlerfreie, rhetorische Prosa neben fehlerhaften Titeln/Meilensteinen, Füllbeispiele ohne Bezug, viel Ankündigung ohne konkreten Versuch. Das sind Indizien, kein Beweis – nur für die Lehrer.",
    "",
    "Regeln für die Kommentare an den Schüler:",
    "- Schreibe in der Sprache des Projektplans (meist Deutsch, manchmal Englisch oder Französisch), duze den Schüler, sei klar, konkret und ermutigend, ohne Noten oder Punkte.",
    "- Ein Kommentar = ein Punkt, den der Schüler im Plan ändern kann. Nenne genau, was er ergänzen oder ändern soll. Keine Allgemeinplätze.",
    "- Beginne mit dem, was schon gut ist (Feld „Allgemein“, Priorität „klein“), aber nur wenn es wirklich zutrifft.",
    "- Aufgaben-Vorschlag: konkrete Arbeitsschritte für genau dieses Projekt, jeweils mit Theorie (für die Dokumentation) und Praxis (Bau, Messung, Versuch).",
    "- Meilenstein-Vorschlag: nur bis Ende des ersten " + (ass1GSE ? "Semesters" : "Trimesters") + ", 4–6 Meilensteine mit Datum (ab dem heutigen Datum) und überprüfbarem Ergebnis, nicht nur einem Thema.",
    "- Wenn der Plan gut ist, mach wenige Kommentare. Erfinde keine Mängel.",
    "",
    "Alles im Feld „fuer_lehrer“ sehen nur die Lehrer. Schreibe es auf Deutsch, knapp.",
  ].join("\n");
}

const KI_FEEDBACK_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["kommentare", "aufgaben_vorschlag", "meilenstein_vorschlag", "fuer_lehrer"],
  properties: {
    kommentare: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["feld", "prioritaet", "text"],
        properties: {
          feld: { type: "string", enum: ["Allgemein", "Titel", "Beschreibung", "Motivation", "Ziele", "Meilensteine", "Aufgaben", "Kostenplan", "Umweltbezug", "Sprache", "Sicherheit"] },
          prioritaet: { type: "string", enum: ["wichtig", "mittel", "klein"] },
          text: { type: "string" },
        },
      },
    },
    aufgaben_vorschlag: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["schritt", "theorie", "praxis"],
        properties: {
          schritt: { type: "string" },
          theorie: { type: "string" },
          praxis: { type: "string" },
        },
      },
    },
    meilenstein_vorschlag: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["datum", "ergebnis"],
        properties: {
          datum: { type: "string", description: "Format TT.MM.JJJJ" },
          ergebnis: { type: "string" },
        },
      },
    },
    fuer_lehrer: {
      type: "object",
      additionalProperties: false,
      required: ["zusammenfassung", "raster_schaetzung", "ki_indizien", "gespraechsfragen"],
      properties: {
        zusammenfassung: { type: "string" },
        raster_schaetzung: { type: "string", description: "z. B. Beschreibung 8/15, Motivation 2/3, Ziele 3/6, Aufgaben 1/6 – mit je einem Halbsatz Begründung" },
        ki_indizien: { type: "string", description: "leer lassen, wenn es keine Hinweise gibt" },
        gespraechsfragen: { type: "array", items: { type: "string" } },
      },
    },
  },
};


// =====================================================================
// PDF-Export (LaTeX, selwecht Deckblat/Layout wéi de Projektplang-PDF)
// data.pdf = { fuerProffen: bool, punkte: [{feld, text, quell, aktiv}], aktivId }
// Gëtt d'PDF als base64 zréck (Download am Browser), späichert näischt am Drive.
// =====================================================================
function feedbackPdfErstellen_(data) {
  const plang = kiPlangLiesen_(data.schueler);
  if (!plang) return { ok: false, error: "Kee Projektplang fonnt fir " + data.schueler + "." };
  const fuerProffen = !!data.pdf.fuerProffen;
  const historique = kiFeedbackHistorique_(data.schueler);
  const aktiv = historique.find((h) => String(h.id) === String(data.pdf.aktivId)) || historique[0] || null;

  const tex = feedbackPdfLatex_(plang, data.pdf.punkte || [], fuerProffen, aktiv, historique);
  const seVectorBlob = DriveApp.getFileById(LOGO_VECTOR_PDF_ID).getBlob();
  const ltettBlob = UrlFetchApp.fetch(LTETT_LOGO_URL).getBlob();
  const payload = {
    compiler: "pdflatex",
    resources: [
      { main: true, content: tex },
      { path: "se_logo_vector.pdf", file: Utilities.base64Encode(seVectorBlob.getBytes()) },
      { path: "ltett_logo.png", file: Utilities.base64Encode(ltettBlob.getBytes()) },
    ],
  };
  let antwort, feeler;
  for (let versuch = 0; versuch < 2; versuch++) {
    try {
      antwort = UrlFetchApp.fetch("https://latex.ytotech.com/builds/sync", {
        method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true,
      });
      if (antwort.getResponseCode() === 201) break;
      feeler = "HTTP " + antwort.getResponseCode() + ": " + String(antwort.getContentText()).substring(0, 300);
    } catch (e) {
      feeler = e.message;
    }
    if (versuch === 0) Utilities.sleep(2000);
  }
  if (!antwort || antwort.getResponseCode() !== 201) {
    return { ok: false, error: "LaTeX-Kompiléierung feelgeschloen: " + (feeler || "onbekannte Feeler") };
  }
  const nummDeel = String(anzeigeNummFirSchueler(plang.schueler) || plang.schueler).replace(/[^\wÀ-ÿ.-]+/g, "_");
  return {
    ok: true,
    dateiname: "Feedback_Projektplan_" + nummDeel + (fuerProffen ? "_LEHRER" : "") + ".pdf",
    pdfBase64: Utilities.base64Encode(antwort.getBlob().getBytes()),
  };
}

function feedbackPdfLatex_(plang, punkte, fuerProffen, aktiv, historique) {
  const E = latexEscape;
  const EA = latexEscapeAbsaetz;
  const betreuer = String(kiBetreuerVun_(plang.schueler) || "");
  const deckblat = baueDeckblattFragmentLatex(fuerProffen ? "Feedback – Lehrerversion" : "Feedback zum Projektplan", {
    schueler: plang.schueler, klasse: plang.klasse, betreuer, betreuer2: "",
  });

  const reiefolleg = ["Allgemein", "Titel", "Beschreibung", "Motivation", "Umweltbezug", "Ziele", "Aufgaben", "Meilensteine", "Kostenplan", "Sicherheit", "Sprache"];
  const gewielt = punkte.filter((k) => String(k.text || "").trim() && (fuerProffen || k.aktiv));
  const felder = [];
  gewielt.forEach((k) => { if (felder.indexOf(k.feld) < 0) felder.push(k.feld); });
  felder.sort((a, b) => ((reiefolleg.indexOf(a) + 99) % 99) - ((reiefolleg.indexOf(b) + 99) % 99));

  const punkteTex = felder.map((f) => {
    const items = gewielt.filter((k) => k.feld === f).map((k) => {
      const meta = fuerProffen ? ` {\\color{selightgray}\\footnotesize (${E(k.quell || "")}${k.aktiv ? "" : ", nicht gesendet"})}` : "";
      const text = EA(String(k.text).trim()).replace(/\n\n/g, "\\\\\n");
      return `\\item ${k.aktiv || !fuerProffen ? "" : "{\\color{segray}"}${text}${k.aktiv || !fuerProffen ? "" : "}"}${meta}`;
    }).join("\n");
    return `{\\color{seblue}\\large\\bfseries ${E(f)}}\\par\\vspace{-2mm}\n\\begin{itemize}[leftmargin=6mm, itemsep=1.5mm, topsep=1mm]\n${items}\n\\end{itemize}\n\\vspace{2mm}`;
  }).join("\n\n");

  let lehrerTex = "";
  if (fuerProffen && aktiv && aktiv.resultat && aktiv.resultat.fuer_lehrer) {
    const l = aktiv.resultat.fuer_lehrer;
    const fragen = (l.gespraechsfragen || []).map((f) => `\\item ${E(f)}`).join("\n");
    lehrerTex = `\\noindent\\fcolorbox{seorange}{selehrer}{\\begin{minipage}{\\dimexpr\\linewidth-2\\fboxsep-2\\fboxrule}
{\\color{seorange}\\bfseries\\small NUR FÜR DAS LEHRPERSONAL ~~\$\\cdot\$~~ KI-Entwurf vom ${E(aktiv.erstallt)}${aktiv.planStand ? ", Plan-Stand " + E(aktiv.planStand) : ""}}\\\\[1mm]
${l.zusammenfassung ? EA(l.zusammenfassung) + "\\\\[1mm]" : ""}
${l.raster_schaetzung ? "\\textbf{Raster (KI-Einschätzung):} " + E(l.raster_schaetzung) + "\\\\[1mm]" : ""}
${l.ki_indizien ? "\\textbf{KI-Indizien:} " + E(l.ki_indizien) + "\\\\[1mm]" : ""}
${fragen ? "\\textbf{Fragen für das Gespräch:}\n\\begin{itemize}[leftmargin=6mm, itemsep=0.5mm, topsep=1mm]\n" + fragen + "\n\\end{itemize}" : ""}
\\end{minipage}}\\par\n\\vspace{6mm}\n`;
  }

  let entwTex = "";
  if (fuerProffen && historique.length > 1) {
    const zeilen = historique.slice().reverse().map((h) =>
      `${E(h.erstallt)} & ${E(h.planStand || "--")} & ${E((h.resultat && h.resultat.fuer_lehrer && h.resultat.fuer_lehrer.raster_schaetzung) || "--")} \\\\\n\\hline`).join("\n");
    entwTex = `{\\color{seblue}\\large\\bfseries Entwicklung}\\par\\vspace{1mm}
\\begin{tabular}{|p{30mm}|p{30mm}|p{95mm}|}
\\hline
\\textbf{KI-Entwurf} & \\textbf{Plan-Stand} & \\textbf{Raster (KI-Einschätzung)} \\\\
\\hline
${zeilen}
\\end{tabular}\\par
\\vspace{6mm}\n`;
  }

  const datum = Utilities.formatDate(new Date(), "Europe/Luxembourg", "dd.MM.yyyy");
  const fuss = fuerProffen ? "" : `\\vspace{4mm}\n{\\color{segray}\\small Arbeite diese Punkte in deinen Projektplan in PPREN ein und gib ihn erneut ab. Halte Gespräche mit deinen Betreuern selbst schriftlich fest -- gleich danach, gerne auch mit KI (aber kontrolliere das Ergebnis).}`;

  return `\\documentclass[a4paper,11pt]{article}
\\usepackage[T1]{fontenc}
\\usepackage[utf8]{inputenc}
\\usepackage[ngerman]{babel}
\\usepackage{geometry}
\\geometry{margin=0mm}
\\usepackage{textpos}
\\usepackage{xcolor}
\\usepackage{graphicx}
\\usepackage[scaled]{helvet}
\\usepackage{enumitem}
\\usepackage{parskip}
\\renewcommand{\\familydefault}{\\sfdefault}
\\setlength{\\parindent}{0mm}
\\setlength{\\TPHorizModule}{1mm}
\\setlength{\\TPVertModule}{1mm}
\\textblockorigin{0mm}{0mm}
\\definecolor{segreen}{HTML}{4CAF50}
\\definecolor{seblack}{HTML}{1c2621}
\\definecolor{seorange}{HTML}{E8952E}
\\definecolor{seblue}{HTML}{3A6EA5}
\\definecolor{segray}{HTML}{55625A}
\\definecolor{selightgray}{HTML}{93A098}
\\definecolor{seteal}{HTML}{3A9BB5}
\\definecolor{selehrer}{HTML}{FBF3E2}
\\newcommand{\\selogo}[1]{\\includegraphics[trim=162pt 271pt 160pt 297pt, clip, width=#1]{se_logo_vector.pdf}}
\\begin{document}
\\pagestyle{empty}
${deckblat}
\\clearpage
\\newgeometry{margin=25mm, top=20mm}
\\pagestyle{plain}
\\pagenumbering{arabic}
{\\color{segreen}\\LARGE\\bfseries ${fuerProffen ? "Feedback (Lehrerversion)" : "Feedback zu deinem Projektplan"}}\\\\[2mm]
{\\color{segray}\\normalsize ${E(anzeigeNummFirSchueler(plang.schueler))} ~~\$\\cdot\$~~ ${E(plang.klasse)} ~~\$\\cdot\$~~ ${E(plang.titel || "")}}\\\\[1mm]
{\\color{segray}\\small Feedback vom ${E(datum)}${plang.stand ? " ~~\$\\cdot\$~~ Projektplan vom " + E(plang.stand) : ""}}
\\vspace{6mm}

${lehrerTex}${entwTex}${punkteTex || "Keine Punkte ausgewählt."}
${fuss}
\\end{document}
`;
}

function kiBetreuerVun_(schueler) {
  const werte = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0].getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) if (werte[i][0] === schueler) return werte[i][2];
  return "";
}
