/* The Ultraspeaker QR Lab — bibliografia in PDF (A4, pagine bianche, Champagne & Limousines).
   Riceve le voci già formattate da reference.js e restituisce un Blob PDF.
   jsPDF e i font si caricano solo al primo uso, per non appesantire l'apertura del Lab. */
(function () {
  "use strict";
  const V = "?v=20261111b";
  const W = 210, H = 297, M = 20, BOTTOM = H - 26;
  const C = { black: "#000000", ash: "#8E8E8E", silver: "#D0D0D0" };
  const PT = 0.3528; // 1 pt in mm
  const QR = 16;     // lato del QR accanto a ogni voce, in mm

  function loadScript(src) {
    return new Promise((res, rej) => { const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
  }
  async function fileB64(url) {
    const b = new Uint8Array(await (await fetch(url)).arrayBuffer());
    let s = "";
    for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }
  // Logo The Ultraspeaker nero, rasterizzato una volta: jsPDF non disegna gli SVG
  function logoPng() {
    return new Promise((res, rej) => {
      const img = new Image();
      img.onload = () => {
        const c = document.createElement("canvas"); c.width = 1200; c.height = Math.round(1200 * 802 / 3955);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL("image/png"));
      };
      img.onerror = rej;
      img.src = "assets/ultraspeaker-logo-black.svg" + V;
    });
  }
  let assets = null;
  function prepare() { // si chiama all'apertura della finestra, così al clic il PDF è pronto subito
    if (!assets) {
      assets = Promise.all([
        window.jspdf ? null : loadScript("vendor/jspdf.umd.min.js" + V),
        fileB64("fonts/pdf/champagne-limousines-regular.ttf"),
        fileB64("fonts/pdf/champagne-limousines-bold.ttf"),
        logoPng(),
      ]).then(([, regular, bold, logo]) => ({ regular, bold, logo }));
      assets.catch(() => { assets = null; });
    }
    return assets;
  }

  /* m = { title, speaker, date, label, powered, entries: [{ text, url, qr, qrUrl }] }
     text: citazione (vuota se lo speaker l'ha esclusa); url: link cliccabile; qr: PNG in data URL */
  async function build(m) {
    const a = await prepare();
    const doc = new window.jspdf.jsPDF({ unit: "mm", format: "a4", compress: true });
    doc.addFileToVFS("cl-r.ttf", a.regular); doc.addFont("cl-r.ttf", "CL", "normal");
    doc.addFileToVFS("cl-b.ttf", a.bold); doc.addFont("cl-b.ttf", "CL", "bold");
    doc.setProperties({ title: m.label + " — " + m.title, author: m.speaker || "The Ultraspeaker", creator: "The Ultraspeaker QR Lab" });

    const font = (style, size, color) => { doc.setFont("CL", style); doc.setFontSize(size); doc.setTextColor(color); };
    const lh = (size, f) => size * PT * f; // altezza di riga in mm
    // etichetta in maiuscolo spaziato 0,32 em, come nelle guide
    function label(txt, x, y) {
      font("bold", 7.5, C.ash);
      const cs = 7.5 * 0.32 * PT;
      doc.setCharSpace(cs); doc.text(String(txt).toUpperCase(), x, y); doc.setCharSpace(0);
    }
    function footer() {
      font("normal", 7, C.ash);
      const pw = doc.getTextWidth(m.powered), lw = 26, lhh = lw * 802 / 3955, gap = 2;
      const x = (W - (pw + gap + lw)) / 2, y = H - 13;
      doc.text(m.powered, x, y);
      doc.addImage(a.logo, "PNG", x + pw + gap, y - lhh + 0.9, lw, lhh);
    }

    // ---------- intestazione ----------
    let y = 24;
    label(m.label, M, y); y += 9;
    font("bold", 20, C.black);
    doc.splitTextToSize(m.title, W - 2 * M).forEach((l) => { doc.text(l, M, y); y += lh(20, 1.2); });
    y += 1;
    if (m.speaker) { font("normal", 12, C.black); doc.text(m.speaker, M, y); y += lh(12, 1.6); }
    if (m.date) { font("normal", 10, C.ash); doc.text(m.date, M, y); y += lh(10, 1.6); }
    y += 2;
    doc.setDrawColor(C.silver); doc.setLineWidth(0.3); doc.line(M, y, W - M, y);
    y += 8;
    footer();

    // ---------- voci, senza numeri ----------
    doc.setLineHeightFactor(1.6);
    m.entries.forEach((e) => {
      const tw = W - 2 * M - (e.qr ? QR + 5 : 0);
      font("normal", 10, C.black);
      const lines = e.text ? doc.splitTextToSize(e.text, tw) : [];
      const textH = lines.length * lh(10, 1.6);
      // senza citazione il link è la voce: stessa grandezza e colore del testo
      const us = e.text ? 8.5 : 10, uc = e.text ? C.ash : C.black;
      font("normal", us, uc);
      const urlLines = e.url ? doc.splitTextToSize(e.url, tw) : [];
      const urlH = urlLines.length * lh(us, 1.6);
      const h = Math.max(textH + urlH, e.qr ? QR : 0);
      if (y + h > BOTTOM) { doc.addPage(); footer(); y = 24; }
      let ty = y + 10 * PT; // linea di base della prima riga
      font("normal", 10, C.black);
      if (lines.length) doc.text(lines, M, ty);
      ty += textH;
      if (e.url) {
        font("normal", us, uc);
        doc.text(urlLines, M, ty);
        doc.link(M, ty - us * PT, tw, urlH, { url: e.url });
      }
      if (e.qr) {
        doc.addImage(e.qr, "PNG", W - M - QR, y, QR, QR);
        if (e.qrUrl) doc.link(W - M - QR, y, QR, QR, { url: e.qrUrl });
      }
      y += h + 5;
    });
    return doc.output("blob");
  }

  window.QRBibPdf = { prepare, build };
})();
