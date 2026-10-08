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
// D'Resultat gëtt NET gespäichert: de Prof wielt am Dashboard aus, wat als
// Kommentar un de Schüler geet.
// =====================================================================

const KI_FEEDBACK_MODELL = "claude-opus-5-5";

function plangKiFeedback(data) {
  const session = pruefSession(data.token);
  if (!session.valid || session.rolle !== "Prof") {
    return { ok: false, error: "Nëmme Proffen dierfen de KI-Feedback benotzen." };
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
  return { ok: true, resultat, modell: antwort.model };
}

function kiPlangLiesen_(schueler) {
  const sheet = SpreadsheetApp.openById(OVERVIEW_SHEET_ID).getSheets()[0];
  const werte = sheet.getDataRange().getValues();
  for (let i = 1; i < werte.length; i++) {
    if (werte[i][0] === schueler) {
      let details = {};
      try { details = JSON.parse(werte[i][9] || "{}"); } catch (e) { details = {}; }
      return {
        schueler: werte[i][0],
        klasse: String(werte[i][1] || ""),
        titel: String(werte[i][3] || ""),
        status: String(werte[i][4] || ""),
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
