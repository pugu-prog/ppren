/**
 * PPREN Auth-Guard
 * ----------------
 * Op JIDDER geschützter Säit als éischt Script agebonnen (virun de React-Scripten).
 * Préift ob eng gëlteg Session existéiert; wa net, gëtt direkt op login.html
 * geleet (mat "zréck"-Parameter, fir nom Umellen hei zréckzekommen).
 *
 * Setzt global "PPREN_SESSION" = { token, numm, rolle, klasse } zur Verfügung,
 * deen d'React-Apps benotze kënnen ({window.PPREN_SESSION}).
 *
 * Notzung: <script src="auth-guard.js"></script>  (virun de React-Scripten)
 * Fir Proffen-only Säiten: <script>PPREN_ERFUERDERT_ROLLE = "Prof";</script>
 * virum auth-guard.js Script-Tag setzen.
 *
 * WICHTEG: Proffen dierfen op JIDD Säit dran, och wa se eng aner Roll
 * (z.B. "Schüler") verlaangt — sou kann e Prof engem Schüler eng Säit
 * weisen/erklären oder am Numm vun engem Schüler eppes nokucken. Just
 * ëmgedréint (Schüler op enger "Prof"-Säit) bleift gespaart.
 */
(function () {
  const token = localStorage.getItem("ppren_token");
  const numm = localStorage.getItem("ppren_numm");
  const rolle = localStorage.getItem("ppren_rolle");
  const klasse = localStorage.getItem("ppren_klasse");

  function laafBassNoLogin() {
    const zréck = encodeURIComponent(window.location.pathname.split("/").pop() + window.location.search);
    window.location.href = "login.html?zréck=" + zréck;
  }

  if (!token || !numm || !rolle) {
    laafBassNoLogin();
    return;
  }

  const erfuerdertRolle = window.PPREN_ERFUERDERT_ROLLE;
  if (erfuerdertRolle && rolle !== erfuerdertRolle && rolle !== "Prof") {
    // document.open() bricht d'Parsen vum ursprénglechen Dokument definitiv of
    // (net just document.write() eleng — dat géif nëmmen den Text asetzen an
    // duerno de Rescht vun der Säit, inkl. spéider Script-Tags, TROTZDEM weider
    // lueden/ausféieren, well de Wurf hei drënner just dëse Script-Block ofbrécht,
    // net d'ganzt Dokument).
    document.open();
    document.write(
      '<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><title>Kee Zougrëff</title></head><body>' +
      '<div style="padding:60px 20px;text-align:center;font-family:-apple-system,\'Segoe UI\',Inter,Arial,sans-serif;color:#8a2b1f;">' +
      "<h2>Kee Zougrëff</h2><p>Dës Säit ass just fir " + erfuerdertRolle + ".</p>" +
      '<a href="links.html" style="color:#00A94F;font-weight:700;">← Zréck</a></div>' +
      '</body></html>'
    );
    document.close();
    throw new Error("PPREN Auth: Roll net erlaabt");
  }

  window.PPREN_SESSION = { token, numm, rolle, klasse };

  /** Mellt of: läscht d'Session um Server, réimt localStorage of, leet op login.html. */
  window.ppRenOfmellen = function () {
    const WEB_APP_URL_LOGOUT = "https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec";
    fetch(WEB_APP_URL_LOGOUT, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ typ: "logout", token }),
    }).finally(() => {
      localStorage.removeItem("ppren_token");
      localStorage.removeItem("ppren_numm");
      localStorage.removeItem("ppren_rolle");
      localStorage.removeItem("ppren_klasse");
      window.location.href = "login.html";
    });
  };

  // Klengt Ofmell-Knäppchen a "Zréck zur Haaptsäit"-Link um Rand vun der Säit
  // derbäisetzen, sou datt béid automatesch iwwerall do sinn, ouni all Säit
  // enzeel upassen ze mussen. "links.html" routt sech selwer weider op
  // proffen.html/schueler.html, jee no Roll.
  function fügOfmellKnäppchenBäi() {
    const startLink = document.createElement("a");
    startLink.textContent = "🏠 Haaptsäit";
    startLink.href = "links.html";
    startLink.style.cssText = [
      "position:fixed", "bottom:12px", "left:12px", "z-index:9999",
      "background:#00A94F", "color:#fff", "border:none", "border-radius:20px",
      "padding:8px 14px", "font-size:11.5px", "font-weight:700",
      "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif",
      "text-decoration:none", "cursor:pointer", "box-shadow:0 2px 8px rgba(0,0,0,0.25)", "opacity:0.85",
    ].join(";");
    startLink.onmouseenter = () => { startLink.style.opacity = "1"; };
    startLink.onmouseleave = () => { startLink.style.opacity = "0.85"; };

    const knäppchen = document.createElement("button");
    knäppchen.textContent = "👤 " + numm + " · Ofmellen";
    knäppchen.setAttribute("onclick", "window.ppRenOfmellen()");
    knäppchen.style.cssText = [
      "position:fixed", "bottom:12px", "right:12px", "z-index:9999",
      "background:#1c2621", "color:#fff", "border:none", "border-radius:20px",
      "padding:8px 14px", "font-size:11.5px", "font-weight:700",
      "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif",
      "cursor:pointer", "box-shadow:0 2px 8px rgba(0,0,0,0.25)", "opacity:0.85",
    ].join(";");
    knäppchen.onmouseenter = () => { knäppchen.style.opacity = "1"; };
    knäppchen.onmouseleave = () => { knäppchen.style.opacity = "0.85"; };

    const bäisetzen = () => {
      document.body.appendChild(startLink);
      document.body.appendChild(knäppchen);
    };
    document.addEventListener("DOMContentLoaded", bäisetzen);
    if (document.readyState === "complete" || document.readyState === "interactive") {
      setTimeout(() => document.body && bäisetzen(), 0);
    }
  }
  fügOfmellKnäppchenBäi();

  // Glack 🔔 (NËMMEN Proffen): zeigt Meldungen aus dem Backend (z.B. geplangt Rendez-vousen,
  // Erënnerungen) mat Zuel vun den neien. Hänkt sech selwer op all Prof-Säit, ouni dass
  // eng Säit enzel ugepasst muss ginn. Fehler hei blockéieren ni d'Säit.
  function fügGlackBäi() {
    if (rolle !== "Prof") return;
    const API = "https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec";
    const FONT = "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif";
    let meldungen = [];
    let panelOffen = false;

    const knapp = document.createElement("button");
    knapp.setAttribute("type", "button");
    knapp.title = "Meldungen";
    knapp.style.cssText = [
      "position:fixed", "bottom:52px", "right:12px", "z-index:9999",
      "background:#fff", "color:#1c2621", "border:1px solid #c9d5cd", "border-radius:20px",
      "padding:7px 13px", "font-size:15px", FONT, "cursor:pointer",
      "box-shadow:0 2px 8px rgba(0,0,0,0.25)",
    ].join(";");
    const zeechen = document.createElement("span");
    zeechen.style.cssText = "display:none;position:absolute;top:-7px;right:-7px;min-width:19px;height:19px;padding:0 5px;border-radius:10px;background:#c0392b;color:#fff;font-size:11px;font-weight:800;line-height:19px;text-align:center;" + FONT;
    const knappText = document.createElement("span");
    knappText.textContent = "🔔";
    knapp.appendChild(knappText);
    knapp.appendChild(zeechen);

    const panel = document.createElement("div");
    panel.style.cssText = [
      "position:fixed", "bottom:96px", "right:12px", "z-index:9999", "display:none",
      "width:min(440px,92vw)", "max-height:62vh", "overflow-y:auto", "background:#fff",
      "border:1px solid #c9d5cd", "border-radius:12px", "padding:12px 14px",
      "box-shadow:0 8px 30px rgba(0,0,0,0.3)", FONT, "font-size:13px", "color:#1c2621",
    ].join(";");

    const ungelies = () => meldungen.filter((m) => !m.gelies).length;
    const el = (tag, css, text) => {
      const x = document.createElement(tag);
      if (css) x.style.cssText = css;
      if (text !== undefined) x.textContent = text;
      return x;
    };

    function zeechnen() {
      const n = ungelies();
      zeechen.textContent = String(n);
      zeechen.style.display = n > 0 ? "block" : "none";
      panel.textContent = "";
      const kapp = el("div", "display:flex;justify-content:space-between;align-items:center;margin-bottom:8px");
      kapp.appendChild(el("strong", "font-size:14px", "🔔 Meldungen" + (n > 0 ? " (" + n + " nei)" : "")));
      if (n > 0) {
        const alle = el("button", "border:1px solid #c9d5cd;background:#eef3ef;color:#00A94F;font-size:11px;font-weight:700;padding:4px 9px;border-radius:7px;cursor:pointer", "✓ Alles gelies");
        alle.setAttribute("type", "button");
        alle.onclick = () => markGelies(null);
        kapp.appendChild(alle);
      }
      panel.appendChild(kapp);
      if (meldungen.length === 0) {
        panel.appendChild(el("div", "color:#6d7d74;font-size:12.5px;padding:8px 0", "Nach keng Meldungen."));
        return;
      }
      meldungen.forEach((m) => {
        const box = el("div", "border:1px solid #e2e8e3;border-left:4px solid " + (m.gelies ? "#c9d5cd" : "#c0392b") + ";border-radius:8px;padding:8px 10px;margin-bottom:8px;background:" + (m.gelies ? "#fff" : "#fffaf0"));
        const zeil = el("div", "display:flex;justify-content:space-between;gap:8px;align-items:baseline");
        zeil.appendChild(el("strong", "font-size:12.5px", m.Titel));
        zeil.appendChild(el("span", "font-size:10.5px;color:#6d7d74;white-space:nowrap", m.Zait));
        box.appendChild(zeil);
        box.appendChild(el("div", "font-size:12px;line-height:1.5;white-space:pre-wrap;margin:5px 0;color:#1c2621", m.Text));
        const akt = el("div", "display:flex;gap:12px;align-items:center");
        if (m.Link) {
          const a = el("a", "color:#00A94F;font-weight:700;text-decoration:none;font-size:12px", "Opmaachen ↗");
          a.href = m.Link;
          akt.appendChild(a);
        }
        if (!m.gelies) {
          const g = el("button", "border:1px solid #c9d5cd;background:#eef3ef;color:#00A94F;font-size:11px;font-weight:700;padding:3px 9px;border-radius:7px;cursor:pointer", "✓ Gelies");
          g.setAttribute("type", "button");
          g.onclick = () => markGelies(m.ID);
          akt.appendChild(g);
        }
        box.appendChild(akt);
        panel.appendChild(box);
      });
    }

    function markGelies(id) {
      meldungen = meldungen.map((m) => (id === null || m.ID === id ? Object.assign({}, m, { gelies: true }) : m));
      zeechnen();
      fetch(API, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(Object.assign({ typ: "notifikatiounGelies", token }, id === null ? { alle: true } : { id })),
      }).catch(() => {});
    }

    function lueden() {
      fetch(API + "?notifikatiounen=1&token=" + encodeURIComponent(token))
        .then((res) => res.json())
        .then((data) => {
          if (data && data.ok) { meldungen = data.notifikatiounen || []; zeechnen(); }
        })
        .catch(() => {});
    }

    knapp.onclick = () => {
      panelOffen = !panelOffen;
      panel.style.display = panelOffen ? "block" : "none";
      if (panelOffen) lueden();
    };

    let bäigesat = false;
    const bäisetzen = () => {
      if (bäigesat || !document.body) return;
      bäigesat = true;
      document.body.appendChild(panel);
      document.body.appendChild(knapp);
      zeechnen();
      lueden();
      setInterval(lueden, 5 * 60 * 1000);
    };
    document.addEventListener("DOMContentLoaded", bäisetzen);
    if (document.readyState === "complete" || document.readyState === "interactive") {
      setTimeout(bäisetzen, 0);
    }
  }
  try { fügGlackBäi(); } catch (e) { /* d'Glack däerf ni d'Säit oder d'Session-Kontroll blockéieren */ }

  // Anonym-Modus 🕶 (NËMMEN Proffen): fir d'Tool virzeweisen (Fortbildung, Beamer),
  // ouni datt Schülernimm, Matrikelen oder E-Mailen ze gesi sinn. Ersetzt se op ALL
  // Säiten duerch "Schüler A", "Schüler B" … (nëmmen an der Uschtung, näischt gëtt
  // geännert oder gespäichert). Un/aus mam Knäppchen ënnen oder Ctrl+Shift+A; de
  // Zoustand bleift am Browser (localStorage "ppren_anonym") bis en ausgeschalt gëtt.
  function anonymModus() {
    if (rolle !== "Prof") return;
    const API = "https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec";
    const SCHLESSEL = "ppren_anonym";
    const un = localStorage.getItem(SCHLESSEL) === "1";

    const knapp = document.createElement("button");
    knapp.textContent = un ? "🕶 Anonym: UN" : "🕶 Anonym";
    knapp.title = "Schülernimm verstoppen (Ctrl+Shift+A)";
    knapp.style.cssText = [
      "position:fixed", "bottom:12px", "left:132px", "z-index:9999",
      "background:" + (un ? "#c0392b" : "#5a6b63"), "color:#fff", "border:none", "border-radius:20px",
      "padding:8px 14px", "font-size:11.5px", "font-weight:700",
      "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif",
      "cursor:pointer", "box-shadow:0 2px 8px rgba(0,0,0,0.25)", "opacity:0.85",
    ].join(";");
    const ëmschalten = () => {
      localStorage.setItem(SCHLESSEL, un ? "0" : "1");
      window.location.reload();
    };
    knapp.onclick = ëmschalten;
    document.addEventListener("keydown", (ev) => {
      if (ev.ctrlKey && ev.shiftKey && (ev.key === "A" || ev.key === "a")) { ev.preventDefault(); ëmschalten(); }
    });
    const knappBäi = () => { if (document.body && !knapp.isConnected) document.body.appendChild(knapp); };
    document.addEventListener("DOMContentLoaded", knappBäi);
    if (document.readyState !== "loading") setTimeout(knappBäi, 0);

    if (!un) return;

    // Bis d'Nimm gelueden sinn: Säit verschwommen, sou datt keen Numm kuerz opblénkt.
    const stil = document.createElement("style");
    stil.textContent = "html.ppren-anonym-lued body > *:not(button){filter:blur(10px)!important}";
    document.documentElement.appendChild(stil);
    document.documentElement.classList.add("ppren-anonym-lued");
    const fräi = () => document.documentElement.classList.remove("ppren-anonym-lued");
    setTimeout(fräi, 8000); // Sécherheet: ni méi wéi 8 s verschwommen

    const esc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    let reegelen = [];
    const ersetzen = (text) => {
      let r = text;
      for (const [re, ersatz] of reegelen) if (re.test(r)) { re.lastIndex = 0; r = r.replace(re, ersatz); }
      return r;
    };
    const behandelen = (wuerzel) => {
      if (!reegelen.length || !wuerzel) return;
      const gänger = document.createTreeWalker(wuerzel, NodeFilter.SHOW_TEXT);
      const knuet = [];
      while (gänger.nextNode()) knuet.push(gänger.currentNode);
      for (const n of knuet) {
        const nei = ersetzen(n.nodeValue);
        if (nei !== n.nodeValue) n.nodeValue = nei;
      }
      if (wuerzel.querySelectorAll) {
        // Input-/Textfelder ginn NET ugepaakt: soss kéint e Pseudonym aus Versinn gespäichert ginn.
        wuerzel.querySelectorAll("[title]").forEach((el) => {
          const t = el.getAttribute("title");
          if (t) { const nei = ersetzen(t); if (nei !== t) el.setAttribute("title", nei); }
        });
      }
    };

    fetch(API + "?personenVoll=1&token=" + encodeURIComponent(token))
      .then((r) => r.json())
      .then((data) => {
        const schüler = ((data && data.personen) || []).filter((p) => p.rolle === "Schüler");
        const BUCHST = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const lis = [];
        schüler.forEach((p, i) => {
          const pseudo = "Schüler " + (BUCHST[i % 26] + (i >= 26 ? Math.floor(i / 26) : ""));
          if (p.numm) lis.push([String(p.numm), pseudo]);
          if (p.email) lis.push([String(p.email), "•••@•••"]);
          if (p.virnumm && String(p.virnumm).trim().length >= 2) lis.push([String(p.virnumm).trim(), pseudo]);
        });
        lis.sort((a, b) => b[0].length - a[0].length); // laang Wierder fir d'éischt
        reegelen = lis.map(([w, ersatz]) => [new RegExp("(?<![\\p{L}\\p{N}])" + esc(w) + "(?![\\p{L}\\p{N}])", "gu"), ersatz]);
        behandelen(document.body);
        new MutationObserver((ännerungen) => {
          for (const ä of ännerungen) {
            if (ä.type === "characterData") {
              const nei = ersetzen(ä.target.nodeValue);
              if (nei !== ä.target.nodeValue) ä.target.nodeValue = nei;
            } else {
              ä.addedNodes.forEach((n) => {
                if (n.nodeType === 3) { const nei = ersetzen(n.nodeValue); if (nei !== n.nodeValue) n.nodeValue = nei; }
                else if (n.nodeType === 1) behandelen(n);
              });
              if (ä.type === "attributes" && ä.target) behandelen(ä.target.parentNode || ä.target);
            }
          }
        }).observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ["title"] });
        fräi();
      })
      .catch(fräi);
  }
  try { anonymModus(); } catch (e) { /* den Anonym-Modus däerf ni d'Säit blockéieren */ }

  // Knäppchen "📅 Terminer" (NËMMEN Proffen): schnelle Zougrëff op d'Iwwersiicht vun de Terminer mat de Schüler.
  function fügTerminerKnappBäi() {
    if (rolle !== "Prof") return;
    if (/terminer\.html$/.test(window.location.pathname)) return;
    const a = document.createElement("a");
    a.textContent = "📅 Terminer";
    a.href = "terminer.html";
    a.title = "Terminer mat de Schüler";
    a.style.cssText = [
      "position:fixed", "bottom:52px", "right:70px", "z-index:9999",
      "background:#fff", "color:#00A94F", "border:1px solid #c9d5cd", "border-radius:20px",
      "padding:8px 14px", "font-size:12px", "font-weight:700",
      "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif",
      "text-decoration:none", "cursor:pointer", "box-shadow:0 2px 8px rgba(0,0,0,0.25)",
    ].join(";");
    let drun = false;
    const bäisetzen = () => {
      if (drun || !document.body) return;
      drun = true;
      document.body.appendChild(a);
    };
    document.addEventListener("DOMContentLoaded", bäisetzen);
    if (document.readyState === "complete" || document.readyState === "interactive") {
      setTimeout(bäisetzen, 0);
    }
  }
  try { fügTerminerKnappBäi(); } catch (e) { /* ni blockéieren */ }

  // Knäppchen "🎯 Virbereeden" (NËMMEN Proffen): Dossier + gezielt Froen pro Schüler fir d'Zwëschegespréicher.
  function fügVirbereedungKnappBäi() {
    if (rolle !== "Prof") return;
    if (/vorbereitung\.html$/.test(window.location.pathname)) return;
    const a = document.createElement("a");
    a.textContent = "🎯 Virbereeden";
    a.href = "vorbereitung.html";
    a.title = "Gespréicher virbereeden";
    a.style.cssText = [
      "position:fixed", "bottom:52px", "right:176px", "z-index:9999",
      "background:#fff", "color:#00A94F", "border:1px solid #c9d5cd", "border-radius:20px",
      "padding:8px 14px", "font-size:12px", "font-weight:700",
      "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif",
      "text-decoration:none", "cursor:pointer", "box-shadow:0 2px 8px rgba(0,0,0,0.25)",
    ].join(";");
    let drun = false;
    const bäisetzen = () => {
      if (drun || !document.body) return;
      drun = true;
      document.body.appendChild(a);
    };
    document.addEventListener("DOMContentLoaded", bäisetzen);
    if (document.readyState === "complete" || document.readyState === "interactive") {
      setTimeout(bäisetzen, 0);
    }
  }
  try { fügVirbereedungKnappBäi(); } catch (e) { /* ni blockéieren */ }

  // Feedback vum Zwëschegespréich um Enn vum Suivi (NËMMEN Schüler, NËMMEN wat de Prof ausdrécklech fräigi gëtt).
  function fügFeedbackAmSuiviBäi() {
    if (rolle !== "Schüler") return;
    if (!/suivi\.html$/.test(window.location.pathname)) return;
    const API = "https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec";
    const KRITAEREN = [
      "Fortschrëtt zënter der leschter Etapp",
      "Erreechung vun de gesate Ziler",
      "Autonomie / Eegenstännegkeet bei der Ëmsetzung",
      "Ëmgang mat opgetruedene Problemer",
      "Nächst Schrëtt sënnvoll geplangt",
    ];
    const FONT = "font-family:-apple-system,'Segoe UI',Inter,Arial,sans-serif";
    const el = (tag, css, text) => {
      const x = document.createElement(tag);
      if (css) x.style.cssText = css;
      if (text !== undefined) x.textContent = text;
      return x;
    };
    const maachen = () => {
      fetch(API + "?zwischenFeedback=1&token=" + encodeURIComponent(token))
        .then((res) => res.json())
        .then((data) => {
          const eintraeg = data && data.ok && Array.isArray(data.eintraeg) ? data.eintraeg : [];
          if (!eintraeg.length) return;
          const wurzel = document.getElementById("root");
          if (!wurzel || !wurzel.parentNode) return;
          const box = el("section", "max-width:980px;margin:16px auto 90px;padding:0 16px;" + FONT + ";color:#1c2621");
          const kaart = el("div", "background:#fff;border:1px solid #e2e8e3;border-radius:14px;padding:16px 18px;box-shadow:0 1px 2px rgba(28,38,33,0.03)");
          kaart.appendChild(el("h2", "margin:0 0 10px;font-size:17px", "💬 Feedback zu de Zwëschegespréicher"));
          eintraeg.forEach((e) => {
            const inh = e.inhalt || {};
            const eent = el("div", "border-top:1px solid #e2e8e3;padding-top:10px;margin-top:10px");
            eent.appendChild(el("div", "font-weight:700;font-size:14px", (e.typ || "Zwëschegespréich") + " vum " + e.datum));
            if (String(inh.gesamt || "").trim()) {
              eent.appendChild(el("div", "background:#eaf6ee;border:1px solid #cfe8d7;border-radius:10px;padding:8px 12px;margin:8px 0;white-space:pre-wrap;font-size:13.5px", inh.gesamt));
            }
            (inh.kriterien || []).filter((k) => String(k.text || "").trim()).forEach((k) => {
              const zeil = el("div", "border-left:4px solid #00A94F;padding:2px 0 2px 12px;margin:8px 0");
              zeil.appendChild(el("div", "font-weight:700;font-size:12.5px;color:#00753a", k.name || KRITAEREN[(Number(k.nr) || 1) - 1] || ""));
              zeil.appendChild(el("div", "white-space:pre-wrap;font-size:13.5px", k.text));
              eent.appendChild(zeil);
            });
            kaart.appendChild(eent);
          });
          box.appendChild(kaart);
          wurzel.parentNode.insertBefore(box, wurzel.nextSibling);
        })
        .catch(() => {});
    };
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", maachen);
    else setTimeout(maachen, 0);
  }
  try { fügFeedbackAmSuiviBäi(); } catch (e) { /* ni blockéieren */ }

  // Session am Hannergrond validéieren (ofgelaf Sessioune erausfannen), ouni
  // d'Säit ze blockéieren — bei Ongëltegkeet gëtt sanft op login.html geleet.
  // De Server ass déi eenzeg authoritativ Quell fir d'Roll: den initialen Check
  // hei driwwer baséiert nëmmen op localStorage, dat kann ëmgeängt ginn (z.B. iwwer
  // DevTools). Dofir gëtt d'Roll hei mat der Server-Äntwert ofgeglach a lokal
  // korrigéiert/geläscht, wann se net iwwerteneestëmmen.
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec";
  fetch(WEB_APP_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ typ: "sessionPruefen", token }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.valid) {
        localStorage.removeItem("ppren_token");
        localStorage.removeItem("ppren_numm");
        localStorage.removeItem("ppren_rolle");
        localStorage.removeItem("ppren_klasse");
        laafBassNoLogin();
        return;
      }
      if (data.rolle && data.rolle !== rolle) {
        // Lokal gespäichert Roll stëmmt net mat där um Server iwwereneen —
        // lokal Wäert war entweder ëmgeängt oder ass veraltet. Ëmmer der
        // Server-Wourecht vertrauen; wa se net méi op dëser Säit erlaabt ass
        // (a keen Prof ass, well Proffen iwwerall dierfen), direkt erausschmeisen
        // amplaz just d'localStorage ze aktualiséieren.
        if (erfuerdertRolle && data.rolle !== erfuerdertRolle && data.rolle !== "Prof") {
          localStorage.removeItem("ppren_token");
          localStorage.removeItem("ppren_numm");
          localStorage.removeItem("ppren_rolle");
          localStorage.removeItem("ppren_klasse");
          window.location.href = "login.html";
          return;
        }
        localStorage.setItem("ppren_numm", data.numm || numm);
        localStorage.setItem("ppren_rolle", data.rolle);
        localStorage.setItem("ppren_klasse", data.klasse || "");
      }
    })
    .catch(() => {}); // Bei Netzwerkfehler: Säit bleift nutzbar, näischt blockéieren
})();
