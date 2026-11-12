# The Ultraspeaker Image Lab

Regola le foto delle slide come un programma di fotoritocco: istogramma con livelli (nero, mezzitoni, bianco),
esposizione, contrasto, saturazione, luci, ombre, nitidezza, riduzione rumore, temperatura e tinta, stili rapidi,
punto neutro, «Migliora qualità», ritaglio a proporzioni (social, slide, stampa) o a misura, rotazione
(90°, 180° e raddrizzamento fino a ±45°), DuoTone/TriTone, zoom, annulla/ripeti.
Interfaccia in italiano, inglese, spagnolo, francese e tedesco. Fa parte dei Lab The Ultraspeaker
e segue la linea guida comune 1.3.

- **Versione web:** <https://niki977.github.io/ultraspeaker-image-lab/>
- **Pannello di PowerPoint:** manifest `image-lab-manifest.xml` (in prova, fase 2)
- **Versione:** 1.4.0 (novembre 2026)

## Come funziona

**Nel browser** si apre una foto (o la si trascina sull'anteprima), si regola e si scarica.
Sul telefono «Salva o condividi» apre il menu del sistema (Foto, WhatsApp…).

**In PowerPoint** si seleziona una foto e si apre Home → The Ultraspeaker → Image Lab.
«Applica alla slide» inserisce la foto regolata nella stessa posizione, misura e rotazione;
l'originale resta nella slide, nascosto. Le regolazioni stanno nelle etichette della foto
(`ULAB_ORIG`, `ULAB_PARAMS`, `ULAB_ROLE`): riselezionando la foto si riprendono, e «Originale» torna alla foto di partenza.

Il ritaglio per social, slide e stampa imposta solo la proporzione: la foto conserva la risoluzione della zona scelta.
Solo «Dimensioni personalizzate» porta la foto a una misura esatta (px, oppure cm con dpi).

## Requisiti

PowerPoint di Microsoft 365 con PowerPointApi 1.10 e ImageCoercion 1.1: Windows 2601 o successivo,
Mac 16.105 o successivo, oppure PowerPoint per il web. iPad non è supportato.

## Struttura

| Percorso | Contenuto |
|---|---|
| `index.html` | Versione web e pannello (una sola pagina) |
| `css/app.css` | Stile proprio del Lab, dopo `shared/ultraspeaker-lab.css` |
| `app.js` | Interfaccia, anteprima, istogramma, ritaglio, rotazione, DuoTone, collegamento con PowerPoint |
| `imaging.js` | Elaborazione dei pixel |
| `i18n.js` | Testi nelle cinque lingue |
| `support.html`, `privacy.html`, `terms.html`, `GUIDA-INSTALLAZIONE.html`, `legal.js` | Pagine di servizio in cinque lingue |
| `image-lab-manifest.xml`, `site.webmanifest` | Manifest dell'add-in e della versione web |
| `shared/`, `fonts/` | Kit comune 1.3 dei Lab (non si modifica qui) |
| `assets/` | Loghi, Arena, favicon e icone (16–512 px) |
| `guide/` | Guida d'utilizzo PDF in cinque lingue, sorgenti, screenshot e `capture.mjs` |
| `store/` | Testi e screenshot per AppSource, presentazione di prova |

Memoria del browser: solo la lingua, in `ultraspeaker.lang` (comune ai Lab). La vecchia chiave `imagelab.lang`
viene copiata al primo avvio e poi cancellata.

## Guida d'utilizzo

Le guide PDF in cinque lingue stanno in `guide/` (`image-lab-guide-<lingua>.pdf`) e si aprono dal pulsante «Guida».
I sorgenti sono `guide/src/guide.<lingua>.html`, gli screenshot `guide/img/<lingua>/`. Per aggiornarle dopo una modifica
all'interfaccia, dalla radice del repository e con Playwright installato:

1. `node guide/capture.mjs` rifà tutti gli screenshot (e `guide/img/pins.json`, la posizione dei richiami);
2. aggiorna i testi in `guide/src/` se serve (l'italiano è la versione di riferimento);
3. `node shared/guide/build-guide.mjs` genera i cinque PDF.

Il capitolo 08 (PowerPoint) ha il riquadro «In arrivo»: quando il pannello esce, va tolto e i PDF rigenerati.

## Pubblicare un aggiornamento

1. Copia i file nella cartella del repository e, in GitHub Desktop, fai «Commit to main» e «Push origin».
2. Aggiorna `?v=AAAAMMGG` su CSS e JS a ogni rilascio (GitHub Pages e PowerPoint tengono i file in cache).
3. Se cambia il manifest, aumenta `<Version>` e sostituiscilo in PowerPoint
   (Mac: `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`); controllalo con
   `npx office-addin-manifest validate image-lab-manifest.xml`.

## Versioni

- **1.4.0** (novembre 2026): uniformazione alla linea guida 1.3 (kit comune, aspetto, lingua comune, pagine di servizio
  scure in cinque lingue, guida d'utilizzo PDF, nuovo indirizzo `ultraspeaker-image-lab`); ritaglio a proporzioni.
- **1.3.0** (ottobre 2026): DuoTone/TriTone, ritaglio, rotazione, «Migliora qualità», semplificazione del codice.
