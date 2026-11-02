// Gemeinsam Hëllef fir wochenberichte-korrigeieren.html a bewertung.html:
// rechent aus, fir wéi eng ofgeschloss Wochen (Frist Sonndes 22:00 laanscht,
// keng Schoulvakanz) e Schüler GUER keen Eintrag am Wochenberichte-Blat huet.
// Esou ass e feelende Bericht direkt ze gesinn, och wann den automateschen
// "Verpasst"-Eintrag (Trigger Sonndes ~23h) net gelaf ass.
window.PPREN_WB = (function () {
  // Muss mat WOCHENBERICHT_START_ISO am PPREN.gs iwwereneestëmmen.
  const START_ISO = "2026-09-21";

  // Lëtzebuerger Schoulvakanzen 2026/2027 (wéi an zeitplang-verwalten.html, Quell: men.public.lu).
  const VAKANZEN = [
    { start: "2026-10-31", end: "2026-11-08" },
    { start: "2026-12-19", end: "2027-01-03" },
    { start: "2027-02-06", end: "2027-02-14" },
    { start: "2027-03-27", end: "2027-04-11" },
    { start: "2027-05-29", end: "2027-06-06" },
    { start: "2027-07-16", end: "2027-09-14" },
  ];

  const zwee = (n) => String(n).padStart(2, "0");
  const iso = (d) => `${d.getFullYear()}-${zwee(d.getMonth() + 1)}-${zwee(d.getDate())}`;
  const dmy = (d) => `${zwee(d.getDate())}.${zwee(d.getMonth() + 1)}.${d.getFullYear()}`;
  const plusDeeg = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };

  // Eng Woch ass Vakanz, wann hire Méindeg bis Freideg ganz an enger Vakanz leien.
  function istVakanzWoch(montag) {
    const mo = iso(montag), fr = iso(plusDeeg(montag, 4));
    return VAKANZEN.some((v) => mo >= v.start && fr <= v.end);
  }

  function labelVunMontag(montag) {
    return `${dmy(montag)} – ${dmy(plusDeeg(montag, 6))}`;
  }

  // "21.09.2026 – 27.09.2026" (oder en Datum-Objet) → Méindeg als Date.
  function montagVunLabel(woche) {
    if (woche instanceof Date) return woche;
    // Huet Google Sheets d'Zell als Datum gespäichert, kënnt se als ISO-Text un.
    if (/^\d{4}-\d{2}-\d{2}T/.test(String(woche))) {
      const d = new Date(woche);
      return new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((d.getDay() + 6) % 7));
    }
    const m = String(woche || "").match(/(\d{2})\.(\d{2})\.(\d{4})/);
    return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
  }

  function istVakanzLabel(woche) {
    const mo = montagVunLabel(woche);
    return !!mo && istVakanzWoch(mo);
  }

  // All Wochen vum Start bis zur leschter Woch, där hir Frist (Sonndes 22:00) laanscht ass, ouni Vakanzen.
  function ofgeschlossWochen(jetzt) {
    const now = jetzt || new Date();
    const [j, m, t] = START_ISO.split("-").map(Number);
    let montag = new Date(j, m - 1, t);
    const resultat = [];
    while (true) {
      const frist = plusDeeg(montag, 6);
      frist.setHours(22, 0, 0, 0);
      if (frist > now) break;
      if (!istVakanzWoch(montag)) resultat.push(labelVunMontag(montag));
      montag = plusDeeg(montag, 7);
    }
    return resultat;
  }

  function standardPeriode(klasse) {
    return klasse === "2GSE" ? "Trimester 1" : "Semester 1";
  }

  // personen: [{ numm, rolle, klasse, virnumm }] (aktiv, vu ?namen=1)
  // wochenberichte: Zeilen vum Blat. Periode vun enger Lück = lescht Period
  // vum Schüler (selwecht Regel wéi schliesseWocheAbIntern am Backend).
  function luecken(personen, wochenberichte, jetzt) {
    const wochen = ofgeschlossWochen(jetzt);
    const resultat = [];
    (personen || []).filter((p) => p.rolle === "Schüler").forEach((p) => {
      const eegen = (wochenberichte || []).filter((w) => w.Schüler === p.numm);
      const mat = new Set(eegen.map((w) => {
        const mo = montagVunLabel(w.Woche);
        return mo ? labelVunMontag(mo) : String(w.Woche);
      }));
      const periode = eegen.length > 0 && eegen[eegen.length - 1].Periode ? eegen[eegen.length - 1].Periode : standardPeriode(p.klasse);
      wochen.forEach((woche) => {
        if (!mat.has(woche)) resultat.push({ Schüler: p.numm, Klasse: p.klasse, Periode: periode, Woche: woche });
      });
    });
    return resultat;
  }

  // Wochenberichter-Punkte fir eng Period: Bewäert = Punkte, Verpasst an
  // feelend (keen Eintrag) = 0 vu 5, Entschëllegt a Vakanz-Wochen zielen net.
  // "Agereecht" (nach net korrigéiert) ziele nach net mat.
  function punkteFirPeriod(schueler, periode, personen, wochenberichte, jetzt) {
    let summe = 0, bewaert = 0, verpasst = 0, feelend = 0;
    (wochenberichte || []).forEach((w) => {
      if (w.Schüler !== schueler || w.Periode !== periode) return;
      if (w.Status === "Bewäert") {
        try {
          const pk = JSON.parse(w["Punkte (JSON)"]);
          summe += (pk.zusammenfassung || 0) + (pk.fortschritt || 0) + (pk.anhaenge || 0) + (pk.grammatik || 0);
          bewaert++;
        } catch (e) { }
      } else if (w.Status === "Verpasst" && !istVakanzLabel(w.Woche)) {
        verpasst++;
      }
    });
    luecken((personen || []).filter((p) => p.numm === schueler), wochenberichte, jetzt)
      .forEach((l) => { if (l.Periode === periode) feelend++; });
    const anzahl = bewaert + verpasst + feelend;
    return { summe, max: anzahl * 5, anzahl, bewaert, verpasst, feelend };
  }

  return { ofgeschlossWochen, luecken, punkteFirPeriod, istVakanzLabel, START_ISO };
})();
