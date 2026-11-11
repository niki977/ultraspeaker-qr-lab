# The Ultraspeaker QR Lab

Crea QR code personalizzati e citazioni scientifiche con un QR che porta al paper, pronti per le slide.
Fa parte della famiglia dei Lab The Ultraspeaker (QR, Image, Flipbook, Projector, Check) e segue la
*Linea guida Lab Ultraspeaker 1.3*.

- **Versione web:** https://niki977.github.io/ultraspeaker-qr-lab/
- **Add-in di PowerPoint:** scheda **Home → The Ultraspeaker → QR Lab** (manifest `manifest.xml`)
- **Pagine di servizio:** `support.html`, `privacy.html`, `terms.html`, `GUIDA-INSTALLAZIONE.html`
- **Guida d'utilizzo:** `guide/qr-lab-guide-<lingua>.pdf` (it, en, es, fr, de)
- **Contatto:** info@theultraspeaker.com · Titolare: Nicholas D. Charles (The Ultraspeaker)

## Cosa fa

- **Sezione QR:** link, testo, email, telefono, SMS, WhatsApp, Wi‑Fi, contatto, posizione, evento, social.
  Design con modelli rapidi, due modelli personali, forme, occhi, colori, logo e cornice.
  Nel web scarica PNG o SVG (sul telefono «Salva o condividi»); nel pannello inserisce nella slide.
  La password del Wi‑Fi non viene mai salvata.
- **Sezione Reference:** ricerca per DOI, PMID, link o testo (Crossref, NCBI/PubMed, Unpaywall),
  stili Vancouver, APA 7, AMA, Harvard; citazione nella slide breve, numerica, autore-anno o completa;
  inserimento con QR, bibliografia in slide, **Aggiungi alla bibliografia**, elenco cliccabile e
  riordinabile, **bibliografia in PDF** (jsPDF, creata sul dispositivo).
- **Collegamento dal Feedback Lab:** solo web, `?url=https://…&name=…&lang=it|en|es|fr|de`
  apre la scheda QR con il link, nella lingua indicata; il file scaricato si chiama `qr-<name>`.

## Struttura

| Percorso | Contenuto |
|---|---|
| `index.html` | Pagina unica, per il web e per il pannello (carica sempre Office.js) |
| `taskpane.html` | Pagina indicata nel manifest: rimanda a `index.html` conservando i parametri |
| `css/app.css` | Stile proprio del Lab, dopo `shared/ultraspeaker-lab.css` |
| `app.js` | Sezione QR, lingua, memoria, avvio; espone `window.QRLab` |
| `reference.js` | Sezione Reference |
| `bib-pdf.js` | Bibliografia in PDF (carica `vendor/jspdf.umd.min.js` e `fonts/pdf/` al primo uso) |
| `qrcode.js`, `i18n.js` | Motore QR (MIT) e testi in cinque lingue |
| `legal.js` | Lingua e menu delle pagine di servizio |
| `shared/`, `fonts/` | Kit comune 1.3: non si modificano qui |
| `guide/` | Guide d'utilizzo: sorgenti, screenshot, `capture.mjs`, PDF |
| `assets/` | Logo, Arena, favicon e icone (16–512 px) |
| `site.webmanifest`, `manifest.xml` | Versione web installabile e add-in di PowerPoint |

**Memoria del browser:** `ultraspeaker.lang` (comune a tutti i Lab) e `ultraspeaker.qr.*`
(`state`, `recent`, `custom`, `lastDesigns`, `mode`, `refs`, `refPrefs`, `speaker`). Le vecchie chiavi
`qrlab.*` e `qrstudio.*` vengono copiate una volta e poi cancellate. Nel pannello i riferimenti stanno
nelle impostazioni del documento (`ultraspeakerReferences`).

## Installazione e aggiornamento

1. Copiare i file nella cartella del repository, poi in GitHub Desktop **Commit to main** e **Push origin**.
2. A ogni rilascio cambiare il valore `?v=AAAAMMGG` su CSS e JS (`index.html`, pagine di servizio,
   `bib-pdf.js`), altrimenti la cache mostra la versione vecchia.
3. Se cambia `manifest.xml`, ricopiarlo nella cartella `wef` di PowerPoint (vedi `GUIDA-INSTALLAZIONE.html`)
   e validarlo con `npx office-addin-manifest validate manifest.xml`.

## Aggiornare la guida d'utilizzo

Serve Playwright con Chromium, installato fuori dal repository (`node_modules` è escluso da `.gitignore`).

```
node guide/capture.mjs              # rifà gli screenshot in guide/img/<lingua>/
node shared/guide/build-guide.mjs   # genera guide/qr-lab-guide-<lingua>.pdf
```

## Versioni

- **1.3.0 (novembre 2026):** uniformazione alla linea guida 1.3 (kit comune, pagina unica, scheda Home,
  lingua condivisa, chiavi `ultraspeaker.qr.*`, pagine di servizio nuove, guida PDF).
  Reference: «Aggiungi alla bibliografia», citazione «Completa», elenco cliccabile e riordinabile,
  bibliografia in PDF. Corretto lo spazio sotto «Il tuo logo».
- **1.2.x (ottobre–novembre 2026):** sezione Reference, pubblicazione dello store preparata,
  collegamento con il Feedback Lab.
