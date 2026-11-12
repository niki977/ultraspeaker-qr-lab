// The Ultraspeaker Image Lab — screenshot della guida d'utilizzo.
// Uso, dalla radice del repository:   node guide/capture.mjs
// Serve il Lab in locale, lo apre con Playwright e salva in guide/img/<lingua>/ gli screenshot di ogni capitolo
// (scala 2×: 1440 px web, 390 px telefono, 350 px pannello simulato con la classe us-addin), più pins.json
// con la posizione dei richiami. La foto è quella di esempio della versione web: nessuna persona, nessun dato.
import { createServer } from "node:http";
import { createRequire } from "node:module";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { extname, join, resolve } from "node:path";

let chromium;
try { ({ chromium } = await import("playwright")); }
catch { ({ chromium } = createRequire(import.meta.url)("playwright")); }

const ROOT = resolve(process.cwd());
const LANGS = ["it", "en", "es", "fr", "de"];
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".woff2": "font/woff2", ".json": "application/json", ".webmanifest": "application/manifest+json" };

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

// PowerPoint simulato: una slide con la foto di esempio selezionata
const MOCK = `(() => {
  const M = window.__MOCK, shapes = [];
  const tags = () => { const items = []; return { items, load() {}, add(k, v) { const i = items.find((x) => x.key === k); if (i) i.value = v; else items.push({ key: k, value: v }); },
    delete(k) { const i = items.findIndex((x) => x.key === k); if (i >= 0) items.splice(i, 1); }, getItemOrNullObject(k) { return items.find((x) => x.key === k) || { isNullObject: true }; } }; };
  const shape = (b64, box) => { const s = Object.assign({ id: "S" + (shapes.length + 1), type: "Image", name: "Foto", rotation: 0, visible: true, load() {}, b64,
    getParentSlide: () => ({ id: "SL1", load() {} }), getImageAsBase64: () => ({ value: s.b64 }), setZOrder() {}, delete() { shapes.splice(shapes.indexOf(s), 1); } }, box);
    s.tags = tags(); Object.defineProperty(s, "zOrderPosition", { get: () => shapes.indexOf(s) }); shapes.push(s); return s; };
  if (M.b64) shape(M.b64, { left: 60, top: 60, width: 480, height: 270 });
  let sel = M.b64 ? ["S1"] : [];
  const byId = (id) => shapes.find((x) => x.id === id);
  const coll = { load() {}, get items() { return shapes.slice(); }, getItem: byId, getItemOrNullObject: (id) => byId(id) || { isNullObject: true, load() {} } };
  const ctx = { sync: () => Promise.resolve(), presentation: { getSelectedShapes: () => ({ load() {}, items: sel.map(byId).filter(Boolean) }), setSelectedShapes(ids) { sel = ids; },
    slides: { getItem: () => ({ id: "SL1", load() {}, shapes: coll }) } } };
  window.PowerPoint = { run: (fn) => fn(ctx) };
  window.Office = { HostType: { PowerPoint: "PowerPoint" }, CoercionType: { Image: "image" }, AsyncResultStatus: { Succeeded: "succeeded" }, EventType: { DocumentSelectionChanged: "s" },
    onReady(cb) { setTimeout(() => cb({ host: "PowerPoint" }), 20); },
    context: { displayLanguage: M.lang, requirements: { isSetSupported: () => true }, ui: { openBrowserWindow() {} },
      document: { addHandlerAsync() {}, setSelectedDataAsync(b64, o, cb) { shape(b64, { left: o.imageLeft, top: o.imageTop, width: o.imageWidth, height: o.imageHeight }); cb({ status: "succeeded" }); } } } };
})();`;

const browser = await chromium.launch();
const STATIC = ".pin,.foot{position:static!important}";
const click = (p, sel) => p.locator(sel).first().evaluate((e) => e.click());
const pins = {};

async function open(lang, width, height, opts = {}) {
  const ctx = await browser.newContext({ viewport: { width, height }, deviceScaleFactor: 2, locale: lang, acceptDownloads: true });
  await ctx.addInitScript(() => { try { localStorage.clear(); } catch (e) { /* niente memoria */ } });
  if (opts.panel) {
    await ctx.addInitScript((m) => { window.__MOCK = m; }, { lang, b64: opts.b64 || "" });
    await ctx.route(/appsforoffice/, (r) => r.fulfill({ body: MOCK, contentType: "text/javascript" }));
  } else await ctx.route(/appsforoffice/, (r) => r.fulfill({ body: "", contentType: "text/javascript" }));
  const page = await ctx.newPage();
  page.on("pageerror", (e) => console.error("errore nella pagina:", e.message));
  await page.goto(BASE + "index.html?lang=" + lang, { waitUntil: "networkidle" });
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(600);
  return { ctx, page };
}
// rettangolo di un elemento in coordinate della pagina
const box = (p, sel) => p.locator(sel).first().evaluate((e) => { const r = e.getBoundingClientRect(); return { x: r.left, y: r.top + scrollY, w: r.width, h: r.height }; });
// ritaglio dalla cima del primo elemento al fondo dell'ultimo, sulla larghezza della colonna
async function shotRange(p, file, from, to, pad = 12) {
  await p.addStyleTag({ content: STATIC }); await p.waitForTimeout(150);
  const col = await box(p, ".app"), a = await box(p, from), b = await box(p, to);
  await p.screenshot({ path: file, fullPage: true, clip: { x: col.x, y: Math.max(0, a.y - pad), width: col.w, height: b.y + b.h - a.y + pad * 2 } });
}
// richiami in percentuale rispetto al ritaglio
async function pinsFor(p, clip, list) {
  const out = [];
  for (const [n, sel, dx = 0.5, dy = 0.5] of list) {
    const b = await box(p, sel);
    out.push({ n, left: +(((b.x + b.w * dx - clip.x) / clip.width) * 100).toFixed(1), top: +(((b.y + b.h * dy - clip.y) / clip.height) * 100).toFixed(1) });
  }
  return out;
}

// foto di esempio in PNG per il pannello simulato: la si scarica dalla versione web senza regolazioni
let sampleB64 = "";
{
  const { ctx, page } = await open("en", 1440, 900);
  const [dl] = await Promise.all([page.waitForEvent("download"), click(page, "#applyBtn")]);
  sampleB64 = (await readFile(await dl.path())).toString("base64");
  await ctx.close();
}

for (const lang of LANGS) {
  const dir = join(ROOT, "guide/img", lang);
  await mkdir(dir, { recursive: true });
  pins[lang] = {};

  // 02 · versione web su computer, con i richiami
  {
    const { ctx, page } = await open(lang, 1440, 900);
    const col = await box(page, ".app");
    const clip = { x: col.x - 24, y: 0, width: col.w + 48, height: 900 };
    await page.screenshot({ path: join(dir, "overview.png"), clip });
    pins[lang].overview = await pinsFor(page, clip, [[1, ".us-header-tools", -0.04, 0.5], [2, ".preview", 0.5, 0.45], [3, "#duoBtn", 1.75, 0.5], [4, "#zoombar", 0.5, 0.25], [5, "#loadBtn", -0.02, 0.5], [6, ".hist", 0.5, 0.35], [7, "#applyBtn", 0.96, 0.5]]);
    await ctx.close();
  }
  // 03 · stili rapidi e cursori (stile «Palco»)
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, '#presets [data-p="stage"]'); await page.waitForTimeout(300);
    await shotRange(page, join(dir, "adjust.png"), "#presets", "#sliders");
    await ctx.close();
  }
  // 04 · istogramma e livelli automatici; punto neutro
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, "#autoLv"); await page.waitForTimeout(300);
    await shotRange(page, join(dir, "levels.png"), ".hist-head", "#resetLv");
    await click(page, "#pickBtn"); await page.waitForTimeout(200);
    await shotRange(page, join(dir, "neutral.png"), "#autoLv", "#pickHint");
    await ctx.close();
  }
  // 05 · Migliora qualità e DuoTone / TriTone
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, "#qBtn"); await page.waitForTimeout(400);
    await shotRange(page, join(dir, "quality.png"), "#qBtn", "#qRow");
    await ctx.close();
  }
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, "#duoBtn"); await page.waitForTimeout(200);
    await click(page, "#duoM3"); await page.waitForTimeout(200);
    await click(page, "#duoPresets button:nth-child(5)"); await page.waitForTimeout(400);
    await shotRange(page, join(dir, "duotone.png"), ".pin", "#duoPresets");
    await ctx.close();
  }
  // 06 · ritaglio e rotazione
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, "#cropBtn"); await page.waitForTimeout(150);
    await page.selectOption("#cropFmt", "ig-portrait"); await page.waitForTimeout(300);
    await shotRange(page, join(dir, "crop.png"), ".pin", "#cropPanel");
    await ctx.close();
  }
  {
    const { ctx, page } = await open(lang, 390, 844);
    await click(page, "#rotBtn"); await page.waitForTimeout(150);
    await page.fill("#rotNum", "4.5"); await page.dispatchEvent("#rotNum", "change"); await page.locator("#rotNum").blur(); await page.waitForTimeout(300);
    await shotRange(page, join(dir, "rotate.png"), ".pin", "#rotPanel");
    await ctx.close();
  }
  // 07 · prima/dopo e scarica sul telefono
  {
    const { ctx, page } = await open(lang, 390, 620);
    await click(page, '#presets [data-p="bw"]'); await page.waitForTimeout(200);
    await click(page, "#beforeBtn"); await page.waitForTimeout(300);
    await page.screenshot({ path: join(dir, "before-after.png") });
    await ctx.close();
  }
  // 08 · pannello di PowerPoint con la foto selezionata
  {
    const { ctx, page } = await open(lang, 350, 620, { panel: true, b64: sampleB64 });
    await click(page, '#presets [data-p="stage"]'); await page.waitForTimeout(300);
    await page.screenshot({ path: join(dir, "panel.png") });
    pins[lang].panel = await pinsFor(page, { x: 0, y: 0, width: 350, height: 620 }, [[1, "#status", -0.03, 0.5], [2, ".preview", 0.5, 0.5], [3, "#applyBtn", 0.95, 0.5]]);
    await ctx.close();
  }
  console.log("ok", lang);
}
await writeFile(join(ROOT, "guide/img/pins.json"), JSON.stringify(pins, null, 1));
await browser.close();
server.close();
