/**
 * ===== GSE-Sektioun-Websäit-Integratioun =====
 * Separat Apps-Script-Datei (deelt den globale Scope mat Code.gs).
 *
 * Gëtt aus Code.gs opgeruff, wann e Projektplang fräigeschalt gëtt:
 *   veroeffentlechGseProjet(data.titel, data.gseThemen);
 * (eng Zeil am "Frei"-Block vun erstelleOderAktualisiereProjektplan)
 *
 * Setzt de Projekttitel automatesch an déi jeweileg Themen-Säit(en) op der
 * GSE-Websäit (https://gse-ltett-sektioun.netlify.app, Repo op GitHub), an
 * der Lëscht ënner "project-list-full".
 *
 * Brauch en GitHub Personal Access Token (Contents: Read and write op dem
 * gse-ltett-sektioun-Repo) ënnert dem Numm "GITHUB_TOKEN" an de Script
 * Properties. Ouni Token gëtt just e Log-Eintrag gemaach, kee Feeler —
 * de Projektplang selwer funktionéiert ëmmer, onofhängeg dovun.
 */
const GSE_GITHUB_REPO = "pugu-prog/gse-ltett-sektioun";

function veroeffentlechGseProjet(titel, gseThemen) {
  if (!titel || !gseThemen || gseThemen.length === 0) return;
  const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) {
    Logger.log("ℹ️ GITHUB_TOKEN feelt an de Script Properties — GSE-Websäit-Veröffentlechung iwwersprongen.");
    return;
  }
  gseThemen.forEach((thema) => {
    try {
      fuegProjetToGseThemaPage(thema, titel, token);
    } catch (e) {
      Logger.log("⚠️ GSE-Websäit-Veröffentlechung feelgeschloen fir Thema '" + thema + "': " + e.message);
    }
  });
}

function fuegProjetToGseThemaPage(thema, titel, token) {
  const path = "pages/" + thema + ".html";
  const url = "https://api.github.com/repos/" + GSE_GITHUB_REPO + "/contents/" + path;
  const headers = {
    Authorization: "Bearer " + token,
    Accept: "application/vnd.github+json",
  };

  const getResp = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
  if (getResp.getResponseCode() !== 200) {
    throw new Error("Konnt " + path + " net lueden (HTTP " + getResp.getResponseCode() + ")");
  }
  const fileData = JSON.parse(getResp.getContentText());
  const contentBytes = Utilities.base64Decode(fileData.content.replace(/\n/g, ""));
  let html = Utilities.newBlob(contentBytes).getDataAsString("UTF-8");

  const titelEscaped = String(titel).trim()
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  if (!titelEscaped) return;
  if (html.indexOf(">" + titelEscaped + "<") !== -1) {
    return; // Projet steet do scho, näischt duebel derbäisetzen
  }

  const listeIdx = html.indexOf('class="project-list-full"');
  if (listeIdx === -1) throw new Error('"project-list-full" net fonnt op ' + path);
  const zoumaachIdx = html.indexOf("</ul>", listeIdx);
  if (zoumaachIdx === -1) throw new Error("</ul> net fonnt op " + path);

  const neiElement = "<li>" + titelEscaped + "</li>\n";
  html = html.slice(0, zoumaachIdx) + neiElement + html.slice(zoumaachIdx);

  const neiContentB64 = Utilities.base64Encode(Utilities.newBlob(html, "text/html", "x").getBytes());
  const putPayload = {
    message: "Neie Projet derbäigesat: " + titelEscaped + " (" + thema + ") — automatesch vu PPREN",
    content: neiContentB64,
    sha: fileData.sha,
    branch: "main",
  };
  const putResp = UrlFetchApp.fetch(url, {
    method: "put",
    headers: Object.assign({ "Content-Type": "application/json" }, headers),
    payload: JSON.stringify(putPayload),
    muteHttpExceptions: true,
  });
  if (putResp.getResponseCode() >= 300) {
    throw new Error("GitHub-Push feelgeschloen (HTTP " + putResp.getResponseCode() + "): " + putResp.getContentText().substring(0, 200));
  }
}
