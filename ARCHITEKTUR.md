# PPREN Projet — Architektur-Iwwersiicht

Kuerz Iwwersiicht wou wat läit — als Orientéierung, well d'System iwwer méi Plattformen verspreet ass.

## 1. Front-End — GitHub

- **Repo**: [pugu-prog/ppren](https://github.com/pugu-prog/ppren)
- **Live-URL**: https://pugu-prog.github.io/ppren/ (GitHub Pages) — Astigspunkt fir Proffen a Schüler ass normalerweis [links.html](https://pugu-prog.github.io/ppren/links.html)
- ~26 statesch HTML-Säiten, React iwwer CDN + Babel am Browser (keng Build-Kette, kee `npm install` néideg)
- Jiddereng Säit huet hire ganze Code (JS + CSS) inline an der `.html`-Datei; `auth-guard.js` ass déi eenzeg gedeelt Datei
- Ännerung = Datei am Repo upassen a pushen → gëtt bannent Minutten iwwer GitHub Pages live

## 2. Back-End — Google Apps Script

- **Projet**: "PPREN web" (Editor iwwer [script.google.com/home](https://script.google.com/home) am `gseltett@gmail.com`-Konto opmaachen)
- **Deployte Web-App-URL** (déi d'Front-End benotzt): `https://script.google.com/macros/s/AKfycbymF-ntji7pYfnRIYxgRREtxVTUUgP38AHUjlbcpxyumpT3G2B1WdBf4M1ehmjXL4c/exec`
- Eng eenzeg Datei `Code.gs` (~2300 Zeilen): Login/PIN/Session-Verwaltung, `doGet`/`doPost`-Router fir all Aktioun, Google-Doc-Generatioun (Projektplang, Suivi, Bewertung), E-Mail-Versand
- **Wichteg**: no all Ännerung um Code muss iwwer *Deploy → Manage deployments → ✏️ Edit → New version* nei deployt ginn — de rohen Editor-Stand gëllt net automatesch fir déi live URL

## 3. Datebank — Eng Google Sheet

- **Sheet-ID**: `1eAdAoDkiQMnowCV2sl1mrmTh3GT6W-_946uI06L3JAQ`
- Enthält all Daten als Tabs (Apps Script erstellt feelend Tabs automatesch): Iwwersiicht (Projeten/Status), `Bewertungen`, `Meilensteng`, `Wochenberichte`, `OffiziellZaitplang`, `Fachgespraeche`, `Zieluewerpreiwungen`, `Rendezvousen`, `Login` (Numm/Roll/PIN-Hash), `Sessions` (aktiv Tokens)
- Keng richteg Datebank-Motor — all Zeil gëtt bei all Ufro komplett gelies/duerchsicht (funktionéiert gutt bei där Gréisst, awer skaléiert net onendlech)

## 4. Späicherung — Google Drive

- **Ordner** "PPREN Projekte" (ID `12eeECkttG0U1-zRHM__age2FdZp3MOmy`)
- Dran: automatesch generéiert Google Docs pro Schüler (Projektplang, Suivi, Bewertung), d'Word-/Doc-Virlag, d'Logoen, an eropgelueden Ahänk aus de Wochenberichter

## 5. E-Mail — Gmail (`MailApp`)

- Rendez-vous-Erënnerungen un d'Schüler (2 Deeg virdrun)
- Leeft ënner dem Google-Konto, mat deem d'Apps-Script-Projet deployt ass

## 6. Zäitgesteierten Auslöser

- E Apps-Script-Trigger (net am Code sichtbar — muss am Editor ënner ⏰ *Triggers* separat agestallt sinn) rifft `sendeRendezvousErennerungen()` all Dag ~7:00 op

## Ofhängegkeet vu Google

Nëmmen d'statesch Front-End (Punkt 1) läit op GitHub. **Alles anert** — Auth, Datebank, Doc-Generatioun, E-Mail, Zäitplang — leeft iwwer Google Apps Script / Sheets / Drive / Gmail ënner engem eenzege Google-Konto. E Wiessel vun der Plattform géif e komplette Rewrite vum Backend bedeiten (net just d'Front-End-Säiten).
