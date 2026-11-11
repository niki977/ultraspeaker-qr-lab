// The Ultraspeaker QR Lab — screenshot della guida d'utilizzo.
// Uso, dalla radice del repository:   node guide/capture.mjs
// Serve il Lab in locale, lo apre con Playwright e salva in guide/img/<lingua>/ gli screenshot di ogni capitolo
// (scala 2×: 1440 px web, 390 px telefono, 350 px pannello simulato con la classe us-addin), più pins.json
// con la posizione dei richiami. Articoli e nomi sono inventati (DOI 10.5555, prefisso di prova di Crossref).
import { createServer } from "node:http";
import { createRequire } from "node:module";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { extname, join, resolve } from "node:path";

let chromium;
try { ({ chromium } = await import("playwright")); }
catch { ({ chromium } = createRequire(import.meta.url)("playwright")); }

const ROOT = resolve(process.cwd());
const LANGS = ["it", "en", "es", "fr", "de"];
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".woff2": "font/woff2", ".ttf": "font/ttf", ".json": "application/json", ".webmanifest": "application/manifest+json" };

// Server statico minimo: niente dipendenze
const server = createServer(async (req, res) => {
  try {
    const path = decodeURIComponent(new URL(req.url, "http://x").pathname);
    const file = join(ROOT, path === "/" ? "index.html" : path);
    if (!file.startsWith(ROOT)) throw new Error("fuori dalla cartella");
    const body = await readFile(file);
    res.writeHead(200, { "content-type": TYPES[extname(file)] || "application/octet-stream" });
    res.end(body);
  } catch { res.writeHead(404); res.end(); }
});
await new Promise((r) => server.listen(0, r));
const BASE = `http://localhost:${server.address().port}/`;

// Tre articoli inventati, risposte di Crossref simulate (PubMed e Unpaywall non rispondono)
const art = (doi, title, journal, abbr, year, vol, issue, page, authors) => ({ message: { DOI: doi, title: [title], "container-title": [journal], "short-container-title": [abbr],
  issued: { "date-parts": [[year]] }, volume: vol, issue, page, author: authors.map(([family, given]) => ({ family, given })) } });
const ARTICLES = {
  "10.5555/qrlab.2024.017": art("10.5555/qrlab.2024.017", "Masticatory efficiency during clear aligner therapy: a randomised clinical trial", "Journal of Dental Research Examples", "J Dent Res Ex", 2024, "12", "3", "211-219",
    [["Bianchi", "Laura"], ["Moreau", "Claire"], ["Keller", "Tobias"], ["Santos", "Ana"]]),
  "10.5555/qrlab.2021.044": art("10.5555/qrlab.2021.044", "Digital workflows in implant planning: a systematic review", "Clinical Oral Examples", "Clin Oral Ex", 2021, "8", "2", "77-90",
    [["Novak", "Petra"], ["Ferri", "Marco"]]),
  "10.5555/qrlab.2019.103": art("10.5555/qrlab.2019.103", "Patient-reported outcomes after periodontal surgery", "Periodontal Science Examples", "Periodontal Sci Ex", 2019, "5", "4", "301-308",
    [["Larsen", "Erik"], ["Dubois", "Marie"], ["Rossi", "Paolo"], ["Weber", "Anna"], ["Costa", "Rita"], ["Klein", "Jonas"], ["Moretti", "Sara"]]),
};
const DOIS = Object.keys(ARTICLES);
const SAMPLE = {
  link: "https://www.example.com/corso-2026",
  wifi: { it: "Sala congressi", en: "Conference room", es: "Sala de congresos", fr: "Salle de congrès", de: "Kongresssaal" },
  lecture: { it: "Allineatori trasparenti: cosa dice la letteratura", en: "Clear aligners: what the literature says", es: "Alineadores transparentes: qué dice la literatura", fr: "Aligneurs transparents : ce que dit la littérature", de: "Transparente Aligner: was die Literatur sagt" },
  speaker: "Dr. Laura Bianchi",
};

// PowerPoint simulato per le immagini del pannello
const MOCK = `(() => {
  const settings = { d: {}, get(k) { return this.d[k] || null; }, set(k, v) { this.d[k] = v; }, saveAsync(cb) { cb && cb({}); } };
  window.Office = { onReady(cb) { setTimeout(() => cb({ host: "PowerPoint" }), 20); }, HostType: { PowerPoint: "PowerPoint" },
    AsyncResultStatus: { Succeeded: "succeeded" }, CoercionType: { Text: "text", Image: "image", XmlSvg: "xmlSvg" },
    context: { displayLanguage: "en-GB", requirements: { isSetSupported: (n) => n !== "PowerPointApi" }, document: { settings, setSelectedDataAsync(d, o, cb) { cb({ status: "succeeded" }); }, getSelectedDataAsync(t, cb) { cb({ status: "succeeded", value: "" }); } },
      ui: { openBrowserWindow() {} } } };
})();`;

const browser = await chromium.launch();
async function open(lang, { width, height = 900, office = false }) {
  const ctx = await browser.newContext({ viewport: { width, height }, deviceScaleFactor: 2, locale: "en-GB" });
  await ctx.addInitScript(() => { if (!sessionStorage.getItem("seeded")) { localStorage.clear(); sessionStorage.setItem("seeded", "1"); } });
  await ctx.route("**/appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: office ? MOCK : "" }));
  await ctx.route(/api\.crossref\.org/, (r) => { const u = decodeURIComponent(r.request().url()); const k = DOIS.find((d) => u.includes(d)) || DOIS[0]; r.fulfill({ json: ARTICLES[k] }); });
  await ctx.route(/eutils\.ncbi|api\.unpaywall/, (r) => r.fulfill({ status: 404, body: "" }));
  const page = await ctx.newPage();
  await page.goto(`${BASE}${office ? "taskpane.html" : "index.html"}?lang=${lang}`);
  await page.waitForTimeout(office ? 700 : 400);
  await page.evaluate(() => document.fonts.ready);
  await page.addStyleTag({ content: ".toast{display:none!important}" }); // i messaggi temporanei non finiscono negli screenshot
  return page;
}
const box = (p, s) => p.$eval(s, (e) => { const r = e.getBoundingClientRect(); return { x: r.x + scrollX, y: r.y + scrollY, w: r.width, h: r.height }; });
// Ritaglio dalla pagina intera attorno agli elementi indicati, con un margine
async function shot(p, file, sels, pad = 10, full = false) {
  const bs = await Promise.all(sels.map((s) => box(p, s)));
  const x0 = full ? 0 : Math.max(0, Math.min(...bs.map((b) => b.x)) - pad), x1 = full ? await p.evaluate(() => innerWidth) : Math.max(...bs.map((b) => b.x + b.w)) + pad;
  const y0 = Math.max(0, Math.min(...bs.map((b) => b.y)) - pad), y1 = Math.max(...bs.map((b) => b.y + b.h)) + pad;
  const clip = { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
  await p.screenshot({ path: file, fullPage: true, clip });
  return clip;
}
// posizione di un richiamo in percentuale del ritaglio
const pin = async (p, clip, s, dx = 0, dy = 0.5) => { const b = await box(p, s); return [+(100 * (b.x + b.w * dx - clip.x) / clip.width).toFixed(1), +(100 * (b.y + b.h * dy - clip.y) / clip.height).toFixed(1)]; };
// link di esempio, poi il cursore esce dal campo: niente bordo di focus negli screenshot
async function fillLink(p) {
  await p.fill("#fields input[data-k=url]", SAMPLE.link);
  await p.$eval("#fields input[data-k=url]", (e) => e.blur()); await p.waitForTimeout(150);
}
async function search(p, doi) {
  await p.fill("#refQuery", doi); await p.click("#refSearchBtn");
  await p.waitForSelector("#refCiteCard:not([hidden])"); await p.waitForTimeout(200);
}

for (const lang of LANGS) {
  const dir = join(ROOT, "guide/img", lang); await mkdir(dir, { recursive: true });
  const pins = {};

  // 01-web: versione web con un link, colonna della pagina; richiami sulle parti principali
  {
    const p = await open(lang, { width: 1440, height: 760 });
    await fillLink(p);
    const col = await box(p, ".us-page"), clip = { x: col.x - 30, y: 0, width: col.w + 60, height: 760 };
    await p.screenshot({ path: join(dir, "01-web.png"), clip });
    pins.web = { tabs: await pin(p, clip, ".modes", 0.02, 0.5), type: await pin(p, clip, "#types", -0.01, 0.2), content: await pin(p, clip, "#fields", -0.01, 0.3),
      design: await pin(p, clip, "#designCard .design-tabs", -0.01, 0.5), preview: await pin(p, clip, "#previewImg", 0.0, 0.1), download: await pin(p, clip, "#insertBtn", 0.02, 0.5),
      tools: await pin(p, clip, ".us-header-tools", -0.02, 0.5) };
    await p.context().close();
  }
  // 01-panel: pannello di PowerPoint (simulato) con il link
  {
    const p = await open(lang, { width: 350, height: 900, office: true });
    await fillLink(p);
    await p.screenshot({ path: join(dir, "01-panel.png"), clip: { x: 0, y: 0, width: 350, height: 640 } });
    await p.context().close();
  }
  // 03-preview: anteprima e download; 04-wifi: tipo Wi-Fi con il nome della rete; 05-*: design
  {
    const p = await open(lang, { width: 1440, height: 900 });
    await fillLink(p);
    await shot(p, join(dir, "03-preview.png"), ["#previewCard"]);
    await p.click('#types [data-type="wifi"]'); await p.fill("#fields input[data-k=ssid]", SAMPLE.wifi[lang]); await p.$eval("#fields input[data-k=ssid]", (e) => e.blur()); await p.waitForTimeout(150);
    const parent = (s) => p.$eval(s, (e) => { e.closest(".us-panel").id = e.closest(".us-panel").id || "g-" + e.id; return "#" + e.closest(".us-panel").id; });
    await shot(p, join(dir, "04-wifi.png"), [await parent("#types"), await parent("#fields")]);
    await p.click('#types [data-type="url"]'); await p.waitForTimeout(100);
    await shot(p, join(dir, "05-style.png"), ["#designCard .step-h", "#myPresets"]);
    await p.click("[data-tab=logo]"); await p.click('#iconLogos [data-ico="brand"]'); await p.waitForTimeout(400);
    await shot(p, join(dir, "05-logo.png"), ["#designCard", "#previewImg"]);
    await p.context().close();
  }
  // 06-*: Reference, ricerca e citazione; 07-*: elenco, finestra del PDF
  {
    const p = await open(lang, { width: 1440, height: 900 });
    await p.click("[data-mode=ref]"); await search(p, DOIS[0]);
    const clip = await shot(p, join(dir, "06-search.png"), ["#modeRef main > section:first-child", "#refPreviewCard"]);
    pins.search = { query: await pin(p, clip, "#refQuery", -0.01, 0.5), preview: await pin(p, clip, "#refPreview .ref-title", -0.02, 0.5) };
    await shot(p, join(dir, "06-cite.png"), ["#refCiteCard"]);
    await p.click("#refAddBtn");
    await p.click("details.ref-more summary"); await p.fill("#refBulk", DOIS.slice(1).join("\n")); await p.click("#refBulkBtn");
    await p.waitForFunction(() => document.querySelectorAll(".ref-item").length === 3, null, { timeout: 15000 }); await p.waitForTimeout(200);
    await shot(p, join(dir, "07-list.png"), ["#refList", "#refBibBtns"]);
    await p.click("#refPdfBtn");
    await p.fill("#bibTitle", SAMPLE.lecture[lang]); await p.fill("#bibSpeaker", SAMPLE.speaker); await p.fill("#bibDate", "2026-12-05");
    await p.click("#bibExtra [data-v=link]"); await p.click("#bibExtra [data-v=qr]"); await p.$eval("#bibDate", (e) => e.blur()); await p.waitForTimeout(150);
    await shot(p, join(dir, "07-pdf.png"), ["#bibDialog"], 0);
    await p.context().close();
  }
  // 08-phone: versione web sul telefono, con la barra fissa in basso
  {
    const p = await open(lang, { width: 390, height: 780 });
    await fillLink(p);
    await p.screenshot({ path: join(dir, "08-phone.png") });
    await p.context().close();
  }
  await writeFile(join(dir, "pins.json"), JSON.stringify(pins, null, 1));
  console.log(`ok  guide/img/${lang}/`);
}
await browser.close();
server.close();
