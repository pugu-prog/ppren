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

  // Klengt Ofmell-Knäppchen um Rand vun der Säit derbäisetzen, sou datt
  // et automatesch iwwerall do ass, ouni all Säit enzeel upassen ze mussen.
  function fügOfmellKnäppchenBäi() {
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
    document.addEventListener("DOMContentLoaded", () => document.body.appendChild(knäppchen));
    if (document.readyState === "complete" || document.readyState === "interactive") {
      setTimeout(() => document.body && document.body.appendChild(knäppchen), 0);
    }
  }
  fügOfmellKnäppchenBäi();

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
