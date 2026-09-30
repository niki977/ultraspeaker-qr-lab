/* The Ultraspeaker QR Lab – add-in PowerPoint per creare QR code personalizzati */
(function () {
  "use strict";

  /* ---------- Utilità ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage non disponibile */ } },
  };
  let inOffice = false;
  let toastT;
  function toast(msg) {
    const t = $("#toast"); t.textContent = msg; t.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), 2600);
  }
  /* ---------- Lingua ---------- */
  const I18N = window.QR_I18N;
  const LANGS = I18N.langs.map((l) => l[0]);
  let LANG = "it";
  function t(k, vars) {
    let v = (I18N[LANG] && I18N[LANG][k]) || I18N.it[k] || k;
    if (vars) for (const n in vars) v = v.replace("{" + n + "}", vars[n]);
    return v;
  }
  function detectLang() {
    const saved = store.get("qrlab.lang", null);
    if (saved && LANGS.includes(saved)) return saved;
    let l = "";
    try { if (inOffice && Office.context && Office.context.displayLanguage) l = Office.context.displayLanguage; } catch (e) { /* ignora */ }
    if (!l) l = (navigator.languages && navigator.languages[0]) || navigator.language || "it";
    l = l.slice(0, 2).toLowerCase();
    return LANGS.includes(l) ? l : "en";
  }
  // Applica le traduzioni a tutti gli elementi con data-i18n
  function applyStatic() {
    document.documentElement.lang = LANG;
    $$("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $$("[data-i18n-aria]").forEach((el) => el.setAttribute("aria-label", t(el.dataset.i18nAria)));
    $$("[data-i18n-alt]").forEach((el) => el.setAttribute("alt", t(el.dataset.i18nAlt)));
    $$("[data-i18n-title]").forEach((el) => el.setAttribute("title", t(el.dataset.i18nTitle)));
    $$(".color[data-color]").forEach((box) => { const [p, x] = $$("input", box); if (p) { p.setAttribute("aria-label", t("color.pick")); x.setAttribute("aria-label", t("color.code")); } });
  }
  const FRAME_DEFAULTS = () => LANGS.map((l) => I18N[l]["frame.default"]);

  const ICON = {
    url: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>',
    text: '<path d="M5 6h14M12 6v13M9 19h6"/>',
    email: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m4 7 8 6 8-6"/>',
    phone: '<path d="M6 3h3l2 5-2.5 1.5a11 11 0 0 0 6 6L16 13l5 2v3a2 2 0 0 1-2 2A16 16 0 0 1 4 5a2 2 0 0 1 2-2"/>',
    sms: '<path d="M4 5h16v11H9l-5 4z"/><path d="M8 10h8"/>',
    whatsapp: '<path d="M4 20l1.3-4A8 8 0 1 1 8 18.7z"/><path d="M9.5 9.5c.3 2 2 3.8 4 4.2l1-1 1.5.7"/>',
    wifi: '<path d="M2 9a15 15 0 0 1 20 0M5.5 12.5a10 10 0 0 1 13 0M9 16a5 5 0 0 1 6 0"/><circle cx="12" cy="19" r="1"/>',
    vcard: '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M6 16c.6-1.5 1.7-2 3-2s2.4.5 3 2M15 10h3M15 13h3"/>',
    location: '<path d="M12 21s-7-6.2-7-11a7 7 0 0 1 14 0c0 4.8-7 11-7 11z"/><circle cx="12" cy="10" r="2.5"/>',
    event: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    social: '<circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="6" r="2.5"/><circle cx="18" cy="18" r="2.5"/><path d="m8.2 10.8 7.6-3.6M8.2 13.2l7.6 3.6"/>',
  };

  /* ---------- Tipi di contenuto ---------- */
  const TYPES = [
    { id: "url", fields: [{ k: "url", label: "l.url", ph: "ph.url", type: "url", req: true, sel: true }] },
    { id: "text", fields: [{ k: "text", label: "l.text", ph: "ph.text", area: true, req: true, sel: true }] },
    { id: "email", fields: [
      { k: "to", label: "l.to", ph: "ph.to", type: "email", req: true },
      { k: "subject", label: "l.subject", ph: "ph.subject" },
      { k: "body", label: "l.body", ph: "ph.body", area: true }] },
    { id: "phone", fields: [{ k: "phone", label: "l.phone", ph: "ph.phone", type: "tel", req: true }] },
    { id: "sms", fields: [
      { k: "phone", label: "l.number", ph: "ph.phone", type: "tel", req: true },
      { k: "msg", label: "l.msg", ph: "ph.sms", area: true }] },
    { id: "whatsapp", fields: [
      { k: "phone", label: "l.waPhone", ph: "ph.phone", type: "tel", req: true, hint: "hint.wa" },
      { k: "msg", label: "l.waMsg", ph: "ph.wa", area: true }] },
    { id: "wifi", fields: [
      { k: "ssid", label: "l.ssid", ph: "ph.ssid", req: true },
      { k: "pass", label: "l.pass", ph: "••••••••" },
      { k: "enc", label: "l.enc", select: [["WPA", "opt.wpa"], ["WEP", "opt.wep"], ["nopass", "opt.nopass"]] },
      { k: "hidden", label: "l.hidden", check: true }] },
    { id: "vcard", fields: [
      { row: [{ k: "first", label: "l.first", ph: "ph.first", req: true }, { k: "last", label: "l.last", ph: "ph.last" }] },
      { row: [{ k: "org", label: "l.org", ph: "ph.org" }, { k: "title", label: "l.title", ph: "ph.title" }] },
      { row: [{ k: "mobile", label: "l.mobile", ph: "ph.phone", type: "tel" }, { k: "work", label: "l.work", ph: "ph.work", type: "tel" }] },
      { k: "email", label: "l.email", ph: "ph.email", type: "email" },
      { k: "web", label: "l.web", ph: "ph.web", type: "url" },
      { k: "street", label: "l.street", ph: "ph.street" },
      { row: [{ k: "city", label: "l.city", ph: "ph.city" }, { k: "zip", label: "l.zip", ph: "ph.zip" }] },
      { k: "country", label: "l.country", ph: "ph.country" }] },
    { id: "location", fields: [
      { k: "mode", label: "l.mode", select: [["addr", "opt.addr"], ["coords", "opt.coords"]] },
      { k: "addr", label: "l.addr", ph: "ph.addr", req: true, when: (v) => v.mode !== "coords" },
      { row: [{ k: "lat", label: "l.lat", ph: "45.5210", req: true }, { k: "lng", label: "l.lng", ph: "9.0880", req: true }], when: (v) => v.mode === "coords" }] },
    { id: "event", fields: [
      { k: "title", label: "l.evTitle", ph: "ph.evTitle", req: true },
      { row: [{ k: "start", label: "l.start", type: "datetime-local", req: true }, { k: "end", label: "l.end", type: "datetime-local" }] },
      { k: "loc", label: "l.loc", ph: "ph.loc" },
      { k: "desc", label: "l.desc", ph: "ph.desc", area: true }] },
    { id: "social", fields: [
      { k: "net", label: "l.net", select: [["instagram", "Instagram"], ["linkedin", "LinkedIn"], ["facebook", "Facebook"], ["youtube", "YouTube"], ["tiktok", "TikTok"], ["x", "X (Twitter)"]] },
      { k: "handle", label: "l.handle", ph: "ph.handle", req: true }] },
  ];

  /* ---------- Codifica del contenuto ---------- */
  const trim = (s) => (s || "").trim();
  // Aggiunge https:// se manca. "sito.it:8080" non è uno schema; gli spazi diventano %20 (alcuni lettori troncano il link allo spazio)
  const normUrl = (u) => {
    u = trim(u); if (!u) return "";
    u = u.replace(/\s/g, "%20");
    return /^[a-z][a-z0-9+.-]*:\/\//i.test(u) || /^(mailto|tel|sms|smsto|geo|facetime|whatsapp|skype|maps):/i.test(u) ? u : "https://" + u;
  };
  const cleanPhone = (p) => trim(p).replace(/[^\d+]/g, "");
  const vEsc = (s) => trim(s).replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
  const wEsc = (s) => (s || "").replace(/([\\;,:"])/g, "\\$1");
  const icsDate = (s) => s ? s.replace(/[-:]/g, "").slice(0, 13) + "00" : "";

  function encode(type, v) {
    switch (type) {
      case "url": return normUrl(v.url);
      case "text": return trim(v.text);
      case "email": {
        if (!trim(v.to)) return "";
        const q = [];
        if (trim(v.subject)) q.push("subject=" + encodeURIComponent(trim(v.subject)));
        if (trim(v.body)) q.push("body=" + encodeURIComponent(trim(v.body)));
        return "mailto:" + trim(v.to) + (q.length ? "?" + q.join("&") : "");
      }
      case "phone": return cleanPhone(v.phone) ? "tel:" + cleanPhone(v.phone) : "";
      case "sms": return cleanPhone(v.phone) ? "SMSTO:" + cleanPhone(v.phone) + ":" + trim(v.msg) : "";
      case "whatsapp": {
        let n = trim(v.phone).replace(/[^\d+]/g, "");
        if (!n) return "";
        if (n.startsWith("+")) n = n.slice(1); else if (n.startsWith("00")) n = n.slice(2);
        else if (LANG === "it" && n.length === 10 && n.startsWith("3")) n = "39" + n; // cellulari italiani senza prefisso
        n = n.replace(/\D/g, "");
        return "https://wa.me/" + n + (trim(v.msg) ? "?text=" + encodeURIComponent(trim(v.msg)) : "");
      }
      case "wifi": {
        if (!trim(v.ssid)) return "";
        const enc = v.enc || "WPA";
        return "WIFI:T:" + enc + ";S:" + wEsc(v.ssid) + ";" + (enc !== "nopass" && v.pass ? "P:" + wEsc(v.pass) + ";" : "") + (v.hidden ? "H:true;" : "") + ";";
      }
      case "vcard": {
        if (!trim(v.first) && !trim(v.last)) return "";
        const L = ["BEGIN:VCARD", "VERSION:3.0", "N:" + vEsc(v.last) + ";" + vEsc(v.first) + ";;;", "FN:" + vEsc([trim(v.first), trim(v.last)].filter(Boolean).join(" "))];
        if (trim(v.org)) L.push("ORG:" + vEsc(v.org));
        if (trim(v.title)) L.push("TITLE:" + vEsc(v.title));
        if (trim(v.mobile)) L.push("TEL;TYPE=CELL:" + cleanPhone(v.mobile));
        if (trim(v.work)) L.push("TEL;TYPE=WORK:" + cleanPhone(v.work));
        if (trim(v.email)) L.push("EMAIL:" + trim(v.email));
        if (trim(v.web)) L.push("URL:" + normUrl(v.web));
        if (trim(v.street) || trim(v.city) || trim(v.zip) || trim(v.country)) L.push("ADR;TYPE=WORK:;;" + vEsc(v.street) + ";" + vEsc(v.city) + ";;" + vEsc(v.zip) + ";" + vEsc(v.country));
        L.push("END:VCARD");
        return L.join("\n");
      }
      case "location": {
        if (v.mode === "coords") {
          const la = parseFloat(String(v.lat || "").replace(",", ".")), lo = parseFloat(String(v.lng || "").replace(",", "."));
          if (isNaN(la) || isNaN(lo) || Math.abs(la) > 90 || Math.abs(lo) > 180) return "";
          return "https://www.google.com/maps/search/?api=1&query=" + la + "," + lo;
        }
        return trim(v.addr) ? "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(trim(v.addr)) : "";
      }
      case "event": {
        if (!trim(v.title) || !v.start) return "";
        let end = v.end;
        // fine mancante o precedente all'inizio: durata di 1 ora
        if (!end || end <= v.start) { const d = new Date(v.start); d.setHours(d.getHours() + 1); end = new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16); }
        const L = ["BEGIN:VEVENT", "SUMMARY:" + vEsc(v.title), "DTSTART:" + icsDate(v.start), "DTEND:" + icsDate(end)];
        if (trim(v.loc)) L.push("LOCATION:" + vEsc(v.loc));
        if (trim(v.desc)) L.push("DESCRIPTION:" + vEsc(v.desc));
        L.push("END:VEVENT");
        return L.join("\n");
      }
      case "social": {
        const h = trim(v.handle);
        if (!h) return "";
        if (/^https?:\/\//i.test(h) || /^www\./i.test(h)) return normUrl(h);
        const u = h.replace(/^@/, "").replace(/\s+/g, "");
        const base = { instagram: "https://www.instagram.com/", linkedin: "https://www.linkedin.com/in/", facebook: "https://www.facebook.com/", youtube: "https://www.youtube.com/@", tiktok: "https://www.tiktok.com/@", x: "https://x.com/" }[v.net || "instagram"];
        return base + u;
      }
    }
    return "";
  }

  /* ---------- Stato ---------- */
  const DEFAULT_DESIGN = {
    dotType: "rounded", dotColor: "#000000", useGradient: false, gradColor2: "#6b6b6b", gradType: "linear", gradRot: 45,
    cornerSqType: "extra-rounded", cornerDotType: "circle", eyesSame: true, cornerSqColor: "#000000", cornerDotColor: "#000000",
    bgColor: "#ffffff", transparentBg: false,
    logo: null, logoName: "", logoSize: 0.3, logoMargin: 6, hideDots: true,
    frame: "none", frameText: "SCANSIONAMI", frameColor: "#000000", frameTextColor: "#ffffff",
    ec: "M", margin: 20,
  };
  const state = store.get("qrstudio.state", null) || {};
  const S = {
    type: state.type || "url",
    values: state.values || (window.QR_DEMO || window.QR_WEB ? { url: { url: "www.theultraspeaker.com" } } : {}),
    design: Object.assign({}, DEFAULT_DESIGN, state.design || {}),
    format: state.format || "png",
    sizeCm: state.sizeCm || 5,
  };
  // Il vecchio nero predefinito (#171a26) diventa il nero pieno del marchio
  ["dotColor", "cornerSqColor", "cornerDotColor", "frameColor"].forEach((k) => { if (S.design[k] === "#171a26") S.design[k] = "#000000"; });
  // La forma "Elegante" (classy-rounded) è stata tolta: diventa "Classy"
  if (S.design.dotType === "classy-rounded") S.design.dotType = "classy";
  // Migrazione dai nomi usati nella prima versione
  if (S.design.cornerSqType === "dot") S.design.cornerSqType = "circle";
  if (S.design.cornerDotType === "dot") S.design.cornerDotType = "circle";
  if (S.design.logo && /^data:image\/svg/.test(S.design.logo) && !(S.design.logoName || "").startsWith("icon:")) { /* verrà convertito all'avvio */ }
  /* La password Wi-Fi non viene mai salvata: resta solo in memoria finché il pannello è aperto */
  const noPass = (v) => { const c = JSON.parse(JSON.stringify(v || {})); delete c.pass; return c; };
  const safeValues = (all) => { const c = JSON.parse(JSON.stringify(all || {})); if (c.wifi) delete c.wifi.pass; return c; };
  // Pulizia dei dati salvati dalle versioni precedenti, che potevano contenere la password
  (function purgeSavedPasswords() {
    if (S.values.wifi && S.values.wifi.pass) delete S.values.wifi.pass;
    const st = store.get("qrstudio.state", null);
    if (st && st.values && st.values.wifi && "pass" in st.values.wifi) { delete st.values.wifi.pass; store.set("qrstudio.state", st); }
    const rec = store.get("qrstudio.recent", []);
    // i QR Wi-Fi recenti salvati prima contengono la password anche nella miniatura: li elimino
    const clean = rec.filter((r) => !(r.type === "wifi" && /;P:/.test(r.data || "")));
    if (clean.length !== rec.length) store.set("qrstudio.recent", clean);
  })();
  const saveState = () => {
    const d = Object.assign({}, S.design);
    if (d.logo && d.logo.length > 300000) { d.logo = null; d.logoName = ""; }
    store.set("qrstudio.state", { type: S.type, values: safeValues(S.values), design: d, format: S.format, sizeCm: S.sizeCm });
  };
  const vals = () => (S.values[S.type] = S.values[S.type] || {});

  /* ---------- UI: tipi e campi ---------- */
  function renderTypes() {
    $("#types").innerHTML = TYPES.map((ty) =>
      `<button type="button" class="type" data-type="${ty.id}" aria-pressed="${ty.id === S.type}"><svg viewBox="0 0 24 24">${ICON[ty.id]}</svg>${esc(t("type." + ty.id))}</button>`).join("");
    $$("#types .type").forEach((b) => b.addEventListener("click", () => {
      S.type = b.dataset.type; renderTypes(); renderFields(); update();
    }));
  }

  function fieldHtml(f, v) {
    const id = "f_" + f.k;
    const val = v[f.k] != null ? v[f.k] : "";
    let inp;
    const ph = esc(f.ph ? t(f.ph) : "");
    if (f.select) inp = `<select id="${id}" data-k="${f.k}">${f.select.map(([o, l]) => `<option value="${o}" ${val === o ? "selected" : ""}>${esc(t(l))}</option>`).join("")}</select>`;
    else if (f.check) return `<label class="check"><input type="checkbox" id="${id}" data-k="${f.k}" ${val ? "checked" : ""}> ${esc(t(f.label))}</label>`;
    else if (f.area) inp = `<textarea id="${id}" data-k="${f.k}" placeholder="${ph}">${esc(val)}</textarea>`;
    else inp = `<input type="${f.type || "text"}" id="${id}" data-k="${f.k}" placeholder="${ph}" value="${esc(val)}" autocomplete="off">`;
    const sel = f.sel && inOffice ? `<button type="button" class="inline-btn" data-sel="${f.k}">${esc(t("useSelection"))}</button>` : "";
    return `<div class="field"><label for="${id}">${esc(t(f.label))}${f.req ? "" : ` <span style="font-weight:400;color:var(--ink-3)">${esc(t("optional"))}</span>`}</label>${inp}${f.hint ? `<div class="hint">${esc(t(f.hint))}</div>` : ""}${sel}</div>`;
  }

  function renderFields() {
    const ty = TYPES.find((x) => x.id === S.type);
    const v = vals();
    ty.fields.forEach((f) => { if (f.select && v[f.k] == null) v[f.k] = f.select[0][0]; });
    $("#contentTitle").textContent = t("type." + ty.id + ".t");
    $("#fields").innerHTML = ty.fields.filter((f) => !f.when || f.when(v)).map((f) =>
      f.row ? `<div class="row">${f.row.map((g) => fieldHtml(g, v)).join("")}</div>` : fieldHtml(f, v)).join("");
    $$("#fields [data-k]").forEach((el) => {
      const ev = el.tagName === "SELECT" || el.type === "checkbox" ? "change" : "input";
      el.addEventListener(ev, () => {
        v[el.dataset.k] = el.type === "checkbox" ? el.checked : el.value;
        if (el.tagName === "SELECT" && ty.fields.some((f) => f.when)) renderFields();
        update();
      });
    });
    $$("#fields [data-sel]").forEach((b) => b.addEventListener("click", () => readSelection(b.dataset.sel)));
  }

  function readSelection(k) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded && trim(r.value)) {
        vals()[k] = trim(r.value); renderFields(); update();
      } else toast(t("toast.selectFirst"));
    });
  }

  /* ---------- UI: design ---------- */
  const SAMPLE = "https://example.com/qr";
  // Miniature: stesso motore del QR vero, con un contenuto di esempio
  function miniQR(over, el, zoom) {
    const d = Object.assign({}, DEFAULT_DESIGN, { logo: null, frame: "none", margin: 0, transparentBg: false, bgColor: "#ffffff" }, over);
    try {
      const out = buildSVG(SAMPLE, d, "L", zoom ? { crop: true } : {});
      el.innerHTML = out.svg;
    } catch (e) { /* miniatura non essenziale */ }
  }
  const DOTS = [["square", "dot.square"], ["rounded", "dot.rounded"], ["extra-rounded", "dot.extra"], ["dots", "dot.dots"], ["classy", "dot.classy"], ["connected", "dot.connected"], ["vpills", "dot.vpills"], ["lines", "dot.lines"]];
  const CSQ = [["square", "eye.square"], ["rounded", "eye.rounded"], ["extra-rounded", "eye.extra"], ["circle", "eye.circle"],
    ["leaf", "eye.leaf"], ["leaf-inv", "eye.leafInv"], ["drop-in", "eye.dropIn"], ["drop-out", "eye.dropOut"], ["corner-in", "eye.corner"], ["octagon", "eye.octagon"], ["dots", "eye.dots"]];
  const CDOT = [["square", "eye.square"], ["rounded", "eye.rounded"], ["circle", "eye.circle"], ["leaf", "eye.leaf"], ["leaf-inv", "eye.leafInv"],
    ["drop-in", "eye.dropIn"], ["drop-out", "eye.dropOut"], ["corner-in", "eye.corner"], ["diamond", "eye.diamond"], ["octagon", "eye.octagon"], ["plus", "eye.plus"], ["star", "eye.star"], ["dots9", "eye.dots9"]];

  /* ---------- Occhi (marcatori d'angolo) disegnati in proprio ----------
     Le forme sono definite per l'occhio in alto a sinistra; gli altri due vengono ruotati
     così gli angoli "speciali" (foglia, goccia) puntano sempre verso il centro del QR. */
  const f2 = (n) => Math.round(n * 1000) / 1000;
  const circ = (cx, cy, r) => `M${f2(cx - r)} ${f2(cy)}a${f2(r)} ${f2(r)} 0 1 0 ${f2(2 * r)} 0a${f2(r)} ${f2(r)} 0 1 0 ${f2(-2 * r)} 0Z`;
  function rrect(x, y, w, h, r) { // r = [tl, tr, br, bl]
    const [a, b, c, d] = r.map((v) => Math.max(0, Math.min(v, w / 2, h / 2)));
    return `M${f2(x + a)} ${f2(y)}H${f2(x + w - b)}` + (b ? `A${f2(b)} ${f2(b)} 0 0 1 ${f2(x + w)} ${f2(y + b)}` : "") +
      `V${f2(y + h - c)}` + (c ? `A${f2(c)} ${f2(c)} 0 0 1 ${f2(x + w - c)} ${f2(y + h)}` : "") +
      `H${f2(x + d)}` + (d ? `A${f2(d)} ${f2(d)} 0 0 1 ${f2(x)} ${f2(y + h - d)}` : "") +
      `V${f2(y + a)}` + (a ? `A${f2(a)} ${f2(a)} 0 0 1 ${f2(x + a)} ${f2(y)}` : "") + "Z";
  }
  function octo(x, y, s, k) { const c = s * k; return `M${f2(x + c)} ${f2(y)}H${f2(x + s - c)}L${f2(x + s)} ${f2(y + c)}V${f2(y + s - c)}L${f2(x + s - c)} ${f2(y + s)}H${f2(x + c)}L${f2(x)} ${f2(y + s - c)}V${f2(y + c)}Z`; }
  const RADII = { square: [0, 0, 0, 0], rounded: [0.2, 0.2, 0.2, 0.2], "extra-rounded": [0.36, 0.36, 0.36, 0.36], circle: [0.5, 0.5, 0.5, 0.5],
    leaf: [0.5, 0, 0.5, 0], "leaf-inv": [0, 0.5, 0, 0.5], "drop-in": [0.5, 0.5, 0, 0.5], "drop-out": [0, 0.5, 0.5, 0.5], "corner-in": [0.3, 0.3, 0, 0.3] };
  // Cornice esterna 7x7 moduli (u = lato del modulo)
  function eyeFrame(kind, x, y, u, fill) {
    const S7 = 7 * u, S5 = 5 * u;
    if (kind === "dots") {
      let o = "";
      for (let i = 0; i < 7; i++) for (let j = 0; j < 7; j++) if (i === 0 || j === 0 || i === 6 || j === 6) o += circ(x + (i + 0.5) * u, y + (j + 0.5) * u, u * 0.5);
      return `<path d="${o}" fill="${fill}"/>`;
    }
    let d;
    if (kind === "octagon") d = octo(x, y, S7, 0.26) + octo(x + u, y + u, S5, 0.2);
    else {
      const R = (RADII[kind] || RADII.square).map((k) => k * S7);
      d = rrect(x, y, S7, S7, R) + rrect(x + u, y + u, S5, S5, R.map((r) => (r ? Math.max(0, r - u) : 0)));
    }
    return `<path fill-rule="evenodd" d="${d}" fill="${fill}"/>`;
  }
  // Punto interno 3x3 moduli
  function eyeBall(kind, x, y, u, fill) {
    const S3 = 3 * u, cx = x + S3 / 2, cy = y + S3 / 2;
    switch (kind) {
      case "diamond": return `<path d="M${f2(cx)} ${f2(y - 0.15 * u)}L${f2(x + S3 + 0.15 * u)} ${f2(cy)}L${f2(cx)} ${f2(y + S3 + 0.15 * u)}L${f2(x - 0.15 * u)} ${f2(cy)}Z" fill="${fill}"/>`;
      case "octagon": return `<path d="${octo(x, y, S3, 0.29)}" fill="${fill}"/>`;
      case "plus": return `<path d="${rrect(x + u * 0.9, y, u * 1.2, S3, [0.2 * u, 0.2 * u, 0.2 * u, 0.2 * u])}${rrect(x, y + u * 0.9, S3, u * 1.2, [0.2 * u, 0.2 * u, 0.2 * u, 0.2 * u])}" fill="${fill}"/>`;
      case "star": {
        let d = "";
        for (let i = 0; i < 10; i++) { const r = i % 2 ? S3 * 0.4 : S3 * 0.72, a = -Math.PI / 2 + (i * Math.PI) / 5; d += (i ? "L" : "M") + f2(cx + r * Math.cos(a)) + " " + f2(cy + 0.1 * u + r * Math.sin(a)); }
        return `<path d="${d}Z" fill="${fill}"/>`;
      }
      case "dots9": {
        let o = "";
        for (let i = 0; i < 3; i++) for (let j = 0; j < 3; j++) o += circ(x + (i + 0.5) * u, y + (j + 0.5) * u, u * 0.6);
        return `<path d="${o}" fill="${fill}"/>`;
      }
      default: return `<path d="${rrect(x, y, S3, S3, (RADII[kind] || RADII.square).map((k) => k * S3))}" fill="${fill}"/>`;
    }
  }
  function eyeThumb(outer, inner, active) {
    const on = "#222", off = "#c9ccd6";
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="-0.6 -0.6 8.2 8.2">${eyeFrame(outer, 0, 0, 1, active === "outer" ? on : off)}${eyeBall(inner, 2, 2, 1, active === "inner" ? on : off)}</svg>`;
  }
  const PRESETS = [
    { name: "preset.brand", brand: true, d: { dotType: "rounded", cornerSqType: "drop-in", cornerDotType: "circle", dotColor: "#000000", useGradient: false, eyesSame: true, bgColor: "#ffffff" } },
    { name: "preset.classic", d: { dotType: "square", cornerSqType: "square", cornerDotType: "square", dotColor: "#000000", useGradient: false, eyesSame: true, bgColor: "#ffffff" } },
    { name: "preset.soft", d: { dotType: "rounded", cornerSqType: "extra-rounded", cornerDotType: "circle", dotColor: "#171a26", useGradient: false, eyesSame: true, bgColor: "#ffffff" } },
    { name: "preset.clinic", d: { dotType: "rounded", cornerSqType: "extra-rounded", cornerDotType: "circle", dotColor: "#0a5cad", useGradient: false, eyesSame: false, cornerSqColor: "#062f5c", cornerDotColor: "#16a3c9", bgColor: "#ffffff" } },
    { name: "preset.elegant", d: { dotType: "vpills", cornerSqType: "extra-rounded", cornerDotType: "square", dotColor: "#1d2b64", useGradient: true, gradColor2: "#6b4fd8", gradType: "linear", gradRot: 45, eyesSame: true, bgColor: "#ffffff" } },
    { name: "preset.dots", d: { dotType: "dots", cornerSqType: "circle", cornerDotType: "circle", dotColor: "#0f766e", useGradient: false, eyesSame: true, bgColor: "#ffffff" } },
    { name: "preset.sunset", d: { dotType: "extra-rounded", cornerSqType: "extra-rounded", cornerDotType: "circle", dotColor: "#c2185b", useGradient: true, gradColor2: "#ef6c00", gradType: "linear", gradRot: 135, eyesSame: true, bgColor: "#ffffff" } },
  ];
  const FRAMES = [["none", "frame.none"], ["bottom", "frame.bottom"], ["top", "frame.top"], ["border", "frame.border"], ["bubble", "frame.bubble"], ["corners", "frame.corners"]];
  const BRAND_LOGO = "assets/ultraspeaker-logo.svg"; // simbolo The Ultraspeaker su cerchio bianco
  const LOGO_ICONS = ["url", "email", "phone", "wifi", "location", "event", "vcard", "social"];

  function iconLogo(key, color) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48"><circle cx="24" cy="24" r="23" fill="${color}"/><g transform="translate(12 12)" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${ICON[key]}</g></svg>`;
    return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
  }

  function tileGroup(sel, items, key, thumbFn) {
    const box = $(sel);
    box.innerHTML = items.map(([v, l]) => `<button type="button" class="tile" data-v="${v}" aria-pressed="${S.design[key] === v}"><span class="thumb"></span><span data-i18n="${l}">${esc(t(l))}</span></button>`).join("");
    $$(".tile", box).forEach((b, i) => {
      thumbFn && thumbFn(items[i][0], $(".thumb", b));
      b.addEventListener("click", () => { S.design[key] = b.dataset.v; syncDesignUI(); update(); });
    });
  }

  function buildDesignUI() {
    // Modelli
    $("#presets").innerHTML = PRESETS.map((p, i) => `<button type="button" class="tile" data-i="${i}"><span class="thumb"></span><span data-i18n="${p.name}">${esc(t(p.name))}</span></button>`).join("");
    $$("#presets .tile").forEach((b) => {
      const p = PRESETS[+b.dataset.i].d;
      miniQR(p, $(".thumb", b));
      b.addEventListener("click", () => {
        Object.assign(S.design, p);
        if (PRESETS[+b.dataset.i].brand) { S.design.logoName = "icon:brand"; setIconLogo(); }
        syncDesignUI(); update(); toast(t("toast.preset"));
      });
    });
    renderMyPresets();
    tileGroup("#dotTypes", DOTS, "dotType", (v, el) => miniQR({ dotType: v, dotColor: "#222222", useGradient: false, eyesSame: true, cornerSqType: "square", cornerDotType: "square" }, el, true));
    tileGroup("#cornerSqTypes", CSQ, "cornerSqType", (v, el) => { el.innerHTML = eyeThumb(v, "square", "outer"); });
    tileGroup("#cornerDotTypes", CDOT, "cornerDotType", (v, el) => { el.innerHTML = eyeThumb("square", v, "inner"); });
    tileGroup("#frames", FRAMES, "frame", (v, el) => { el.innerHTML = frameThumb(v); });

    // Loghi a icona
    const ic = $("#iconLogos");
    ic.innerHTML = `<button type="button" class="tile" data-ico=""><span class="thumb" style="font-size:20px;color:var(--ink-3)">∅</span><span data-i18n="logo.none">${esc(t("logo.none"))}</span></button>` +
      `<button type="button" class="tile" data-ico="brand"><span class="thumb"><img alt="" src="${BRAND_LOGO}"></span><span>Ultraspeaker</span></button>` +
      LOGO_ICONS.map((k) => `<button type="button" class="tile" data-ico="${k}"><span class="thumb"><img alt="" src="${iconLogo(k, "#000000")}"></span><span data-i18n="type.${k}">${esc(t("type." + k))}</span></button>`).join("");
    $$("#iconLogos .tile").forEach((b) => b.addEventListener("click", () => {
      const k = b.dataset.ico;
      if (!k) { S.design.logo = null; S.design.logoName = ""; }
      else { S.design.logoName = "icon:" + k; setIconLogo(); return; }
      syncDesignUI(); update();
    }));

    // Colori
    $$(".color[data-color]").forEach((box) => {
      const k = box.dataset.color;
      box.innerHTML = `<input type="color" aria-label="Scegli colore"><input type="text" maxlength="7" spellcheck="false" aria-label="Codice colore">`;
      const [pick, txt] = $$("input", box);
      pick.addEventListener("input", () => { S.design[k] = pick.value; txt.value = pick.value; afterColor(k); });
      txt.addEventListener("change", () => { let c = txt.value.trim(); if (!c.startsWith("#")) c = "#" + c; if (/^#[0-9a-f]{6}$/i.test(c)) { S.design[k] = c.toLowerCase(); pick.value = S.design[k]; afterColor(k); } else txt.value = S.design[k]; });
    });
    function afterColor(k) {
      if (S.design.logoName && S.design.logoName.startsWith("icon:") && S.design.logoName !== "icon:brand" && (k === "dotColor" || k === "gradColor2")) { clearTimeout(iconT); iconT = setTimeout(setIconLogo, 120); }
      update();
    }
    const bindCheck = (id, k, re) => $(id).addEventListener("change", (e) => { S.design[k] = e.target.checked; if (re) syncDesignUI(); update(); });
    bindCheck("#useGradient", "useGradient", true);
    bindCheck("#eyesSame", "eyesSame", true);
    bindCheck("#transparentBg", "transparentBg", true);
    bindCheck("#hideDots", "hideDots");
    const bindRange = (id, k, fmt, num = true) => $(id).addEventListener("input", (e) => { S.design[k] = num ? +e.target.value : e.target.value; syncDesignUI(); update(); });
    bindRange("#gradRot", "gradRot");
    bindRange("#logoSize", "logoSize");
    bindRange("#logoMargin", "logoMargin");
    bindRange("#qrMargin", "margin");
    $("#frameText").addEventListener("input", (e) => { S.design.frameText = e.target.value; update(); });
    segBind("#gradType", (v) => { S.design.gradType = v; });
    segBind("#ecLevel", (v) => { S.design.ec = v; if (S.design.logo && v !== "H") toast(t("toast.ecLogo")); });

    // Upload logo
    $("#logoFile").addEventListener("change", (e) => {
      const f = e.target.files[0]; if (!f) return;
      if (f.size > 4 * 1024 * 1024) { toast(t("toast.tooBig")); return; }
      const r = new FileReader();
      r.onload = () => shrinkImage(r.result, f.type).then((url) => { S.design.logo = url; S.design.logoName = f.name; syncDesignUI(); update(); });
      r.readAsDataURL(f);
      e.target.value = "";
    });

    // Tab
    $$(".tabs [role=tab]").forEach((b) => b.addEventListener("click", () => {
      $$(".tabs [role=tab]").forEach((x) => x.setAttribute("aria-selected", x === b));
      $$(".panel").forEach((p) => (p.hidden = p.dataset.panel !== b.dataset.tab));
    }));
    $("#resetDesign").addEventListener("click", () => { S.design = Object.assign({}, DEFAULT_DESIGN, { frameText: t("frame.default") }); syncDesignUI(); update(); toast(t("toast.reset")); });
  }

  let iconT;
  // Le icone vengono convertite in PNG: PowerPoint non mostra SVG annidati dentro un altro SVG
  function setIconLogo() {
    const k = S.design.logoName.slice(5);
    const src = k === "brand" ? BRAND_LOGO : iconLogo(k, S.design.useGradient ? S.design.gradColor2 : S.design.dotColor);
    rasterize(src, k === "brand" ? 400 : 256).then((url) => {
      if (S.design.logoName !== "icon:" + k) return;
      S.design.logo = url; syncDesignUI(); update();
    });
  }
  function rasterize(src, max) {
    return new Promise((res) => {
      const img = new Image();
      img.onload = () => {
        const w = img.naturalWidth || max, h = img.naturalHeight || max, s = max / Math.max(w, h);
        const c = document.createElement("canvas"); c.width = Math.round(w * s); c.height = Math.round(h * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL("image/png"));
      };
      img.onerror = () => res(src);
      img.src = src;
    });
  }

  /* ---------- I miei modelli: 2 personali (con nome) + 2 ultimi usati ---------- */
  const ICO_SAVE = '<svg viewBox="0 0 24 24"><path d="M5 3h11l3 3v15H5z"/><path d="M8 3v6h8V3M8 21v-7h8v7"/></svg>';
  const ICO_EDIT = '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M13 7l4 4"/></svg>';
  const cleanDesign = (d) => {
    const c = JSON.parse(JSON.stringify(d));
    if (c.logo && c.logo.length > 150000) { c.logo = null; if (!(c.logoName || "").startsWith("icon:")) c.logoName = ""; }
    return c;
  };
  const designKey = (d) => JSON.stringify(Object.assign({}, d, { logo: d.logo ? d.logo.length : 0 }));
  // Porta un design salvato (anche da versioni precedenti) nel formato attuale
  function loadDesign(d) {
    S.design = Object.assign({}, DEFAULT_DESIGN, { frameText: t("frame.default") }, JSON.parse(JSON.stringify(d || {})));
    if (S.design.dotType === "classy-rounded") S.design.dotType = "classy";
    ["cornerSqType", "cornerDotType"].forEach((k) => { if (S.design[k] === "dot") S.design[k] = "circle"; });
    ["dotColor", "cornerSqColor", "cornerDotColor", "frameColor"].forEach((k) => { if (S.design[k] === "#171a26") S.design[k] = "#000000"; });
    if ((S.design.logoName || "").startsWith("icon:") && (!S.design.logo || S.design.logoName === "icon:brand")) setIconLogo();
  }
  function applyDesign(d) {
    loadDesign(d);
    syncDesignUI(); update(); toast(t("toast.preset"));
  }
  // Ultimi 2 design usati (indipendenti dal contenuto del QR)
  function lastDesigns() {
    const saved = store.get("qrlab.lastDesigns", null);
    if (saved) return saved;
    const out = [], seen = new Set(); // prima volta: li ricavo dai QR recenti
    for (const r of store.get("qrstudio.recent", [])) {
      const k = designKey(r.design); if (seen.has(k)) continue;
      seen.add(k); out.push(r.design); if (out.length === 2) break;
    }
    return out;
  }
  function rememberDesign(d) {
    const c = cleanDesign(d), k = designKey(c);
    const list = [c].concat(lastDesigns().filter((x) => designKey(x) !== k)).slice(0, 2);
    store.set("qrlab.lastDesigns", list);
  }
  function renderMyPresets() {
    const box = $("#myPresets"); if (!box) return;
    const slots = store.get("qrlab.custom", [null, null]);
    const last = lastDesigns();
    let html = "";
    [0, 1].forEach((i) => {
      const sl = slots[i], name = sl ? (sl.name || t("custom.default", { n: i + 1 })) : t("custom.default", { n: i + 1 });
      html += `<div class="tile my" data-slot="${i}">
        <button type="button" class="my-main" title="${esc(sl ? name : t("custom.save"))}"><span class="thumb${sl ? "" : " empty"}">${sl ? "" : "+"}</span><span class="nm">${esc(sl ? name : t("custom.empty"))}</span></button>
        ${sl ? `<div class="my-acts"><button type="button" data-act="save" title="${esc(t("custom.save"))}" aria-label="${esc(t("custom.save"))}">${ICO_SAVE}</button><button type="button" data-act="rename" title="${esc(t("custom.rename"))}" aria-label="${esc(t("custom.rename"))}">${ICO_EDIT}</button></div>` : ""}
      </div>`;
    });
    [0, 1].forEach((i) => {
      const d = last[i];
      html += `<button type="button" class="tile" data-last="${i}" ${d ? "" : "disabled"} title="${esc(d ? t("last." + (i + 1)) : t("last.emptyHint"))}"><span class="thumb${d ? "" : " empty"}">${d ? "" : "–"}</span><span>${esc(t("last." + (i + 1)))}</span></button>`;
    });
    box.innerHTML = html;
    [0, 1].forEach((i) => { if (slots[i]) miniQR(slots[i].design, $(`[data-slot="${i}"] .thumb`, box)); if (last[i]) miniQR(last[i], $(`[data-last="${i}"] .thumb`, box)); });
    $$("[data-slot]", box).forEach((el) => {
      const i = +el.dataset.slot;
      const save = () => {
        const cur = store.get("qrlab.custom", [null, null]);
        cur[i] = { name: cur[i] ? cur[i].name : null, design: cleanDesign(S.design) };
        store.set("qrlab.custom", cur); renderMyPresets();
        toast(t("toast.customSaved", { name: cur[i].name || t("custom.default", { n: i + 1 }) }));
      };
      $(".my-main", el).addEventListener("click", () => { const cur = store.get("qrlab.custom", [null, null]); if (cur[i]) applyDesign(cur[i].design); else save(); });
      const bs = $('[data-act="save"]', el), br = $('[data-act="rename"]', el);
      if (bs) bs.addEventListener("click", save);
      if (br) br.addEventListener("click", () => {
        const nm = $(".nm", el), cur = store.get("qrlab.custom", [null, null]);
        const inp = document.createElement("input"); inp.type = "text"; inp.maxLength = 18; inp.placeholder = t("custom.namePh");
        inp.value = cur[i].name || ""; inp.setAttribute("aria-label", t("custom.rename"));
        el.replaceChild(inp, $(".my-acts", el)); inp.focus(); inp.select();
        nm.textContent = "…";
        let done = false;
        const commit = (ok) => {
          if (done) return; done = true;
          if (ok) { const c2 = store.get("qrlab.custom", [null, null]); c2[i].name = inp.value.trim() || null; store.set("qrlab.custom", c2); toast(t("toast.renamed")); }
          renderMyPresets();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") commit(true); if (e.key === "Escape") commit(false); });
        inp.addEventListener("blur", () => commit(true));
      });
    });
    $$("[data-last]", box).forEach((b) => b.addEventListener("click", () => { const d = lastDesigns()[+b.dataset.last]; if (d) applyDesign(d); }));
  }

  function segBind(sel, fn) {
    $$(sel + " button").forEach((b) => b.addEventListener("click", () => { fn(b.dataset.v); syncDesignUI(); update(); }));
  }

  // Riduce loghi raster molto grandi, così il file resta leggero
  function shrinkImage(dataUrl, mime) {
    if (mime === "image/svg+xml") return rasterize(dataUrl, 512);
    return new Promise((res) => {
      const img = new Image();
      img.onload = () => {
        const max = 400, s = Math.min(1, max / Math.max(img.width, img.height));
        if (s === 1 && dataUrl.length < 200000) return res(dataUrl);
        const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        res(c.toDataURL("image/png"));
      };
      img.onerror = () => res(dataUrl);
      img.src = dataUrl;
    });
  }

  function frameThumb(v) {
    const q = '<rect x="14" y="10" width="20" height="20" rx="2" fill="none" stroke="#222" stroke-width="2"/><rect x="18" y="14" width="5" height="5" fill="#222"/><rect x="25" y="21" width="5" height="5" fill="#222"/>';
    const m = {
      none: '<rect x="12" y="12" width="24" height="24" rx="2" fill="none" stroke="#222" stroke-width="2"/><rect x="16" y="16" width="6" height="6" fill="#222"/><rect x="26" y="26" width="6" height="6" fill="#222"/>',
      bottom: '<rect x="9" y="5" width="30" height="38" rx="4" fill="#222"/><rect x="12" y="8" width="24" height="24" rx="2" fill="#fff"/><rect x="16" y="12" width="6" height="6" fill="#222"/><rect x="26" y="22" width="6" height="6" fill="#222"/><rect x="15" y="35" width="18" height="3" rx="1.5" fill="#fff"/>',
      top: '<rect x="9" y="5" width="30" height="38" rx="4" fill="#222"/><rect x="12" y="16" width="24" height="24" rx="2" fill="#fff"/><rect x="16" y="20" width="6" height="6" fill="#222"/><rect x="26" y="30" width="6" height="6" fill="#222"/><rect x="15" y="9" width="18" height="3" rx="1.5" fill="#fff"/>',
      border: '<rect x="9" y="5" width="30" height="38" rx="4" fill="none" stroke="#222" stroke-width="2"/>' + q + '<rect x="15" y="35" width="18" height="3" rx="1.5" fill="#222"/>',
      bubble: q + '<path d="M24 32l-3 3h6z" fill="#222"/><rect x="9" y="35" width="30" height="9" rx="4.5" fill="#222"/><rect x="15" y="38.5" width="18" height="2" rx="1" fill="#fff"/>',
      corners: '<path d="M9 12V6h6M33 6h6v6M39 30v6h-6M15 36H9v-6" fill="none" stroke="#222" stroke-width="2.2"/><rect x="14" y="11" width="20" height="20" rx="2" fill="none" stroke="#222" stroke-width="2"/><rect x="18" y="15" width="5" height="5" fill="#222"/><rect x="25" y="22" width="5" height="5" fill="#222"/><rect x="15" y="40" width="18" height="3" rx="1.5" fill="#222"/>',
    };
    return `<svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">${m[v]}</svg>`;
  }

  function syncDesignUI() {
    const d = S.design;
    const press = (sel, key) => $$(sel + " .tile").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === d[key]));
    press("#dotTypes", "dotType"); press("#cornerSqTypes", "cornerSqType"); press("#cornerDotTypes", "cornerDotType"); press("#frames", "frame");
    $$("#iconLogos .tile").forEach((b) => b.setAttribute("aria-pressed", b.dataset.ico ? d.logoName === "icon:" + b.dataset.ico : !d.logo));
    const setV = (el, v) => { if (el !== document.activeElement && el.value !== String(v)) el.value = v; };
    $$(".color[data-color]").forEach((box) => { const [p, t] = $$("input", box); setV(p, d[box.dataset.color]); setV(t, d[box.dataset.color]); });
    $("#useGradient").checked = d.useGradient; $("#gradOpts").hidden = !d.useGradient; $("#grad2Wrap").hidden = !d.useGradient;
    $("#gradRot").value = d.gradRot; $("#gradRotVal").textContent = d.gradRot + "°";
    $$("#gradType button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === d.gradType));
    $("#eyesSame").checked = d.eyesSame; $("#eyeColors").hidden = d.eyesSame;
    $("#transparentBg").checked = d.transparentBg;
    $("#logoSize").value = d.logoSize; $("#logoSizeVal").textContent = Math.round(d.logoSize * 100) + "%";
    $("#logoMargin").value = d.logoMargin; $("#logoMarginVal").textContent = d.logoMargin;
    $("#hideDots").checked = d.hideDots;
    $("#logoOpts").style.opacity = d.logo ? 1 : 0.45;
    $("#uploadText").textContent = d.logo && !d.logoName.startsWith("icon:") ? t("upload.done", { name: d.logoName || t("upload.defaultName") }) : t("upload.idle");
    $("#frameOpts").style.display = d.frame === "none" ? "none" : "";
    setV($("#frameText"), d.frameText);
    const ec = effectiveEC();
    $$("#ecLevel button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === ec));
    $("#qrMargin").value = d.margin; $("#marginVal").textContent = d.margin;
    $$("#format button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === S.format));
    setV($("#sizeCm"), S.sizeCm);
  }
  const effectiveEC = () => (S.design.logo ? "H" : S.design.ec);

  /* ---------- Generazione ---------- */
  function hexLum(h) {
    const n = parseInt(h.slice(1), 16), c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((x) => { x /= 255; return x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.04045) / 1.055, 2.4); });
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  }
  const contrast = (a, b) => { const x = hexLum(a), y = hexLum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };

  /* Motore di disegno leggero: tutti i moduli in un unico tracciato, niente maschere.
     Il QR è disegnato in "unità modulo" dentro un gruppo scalato: file piccolo e veloce da gestire in PowerPoint. */
  function matrix(data, ec) {
    qrcode.stringToBytes = qrcode.stringToBytesFuncs["UTF-8"];
    const q = qrcode(0, ec);
    q.addData(data, "Byte");
    q.make();
    return q;
  }
  const DOT_R = { // raggi degli angoli esposti [tl, tr, br, bl], in frazioni di modulo
    rounded: [0.3, 0.3, 0.3, 0.3], "extra-rounded": [0.5, 0.5, 0.5, 0.5], classy: [0.5, 0, 0.5, 0], "classy-rounded": [0.5, 0.2, 0.5, 0.2] };

  function bodyPath(q, n, skip, type) {
    const dark = (r, c) => r >= 0 && c >= 0 && r < n && c < n && !skip(r, c) && q.isDark(r, c);
    if (type === "classy-rounded") type = "classy"; // forma rimossa
    // Capsule: punti che si fondono in capsule verticali
    if (type === "vpills") {
      let d = "";
      for (let c = 0; c < n; c++) {
        let r = 0;
        while (r < n) {
          if (!dark(r, c)) { r++; continue; }
          let e = r; while (e + 1 < n && dark(e + 1, c)) e++;
          d += rrect(c + 0.06, r + 0.06, 0.88, e - r + 1 - 0.12, [0.44, 0.44, 0.44, 0.44]);
          r = e + 1;
        }
      }
      return d;
    }
    // Linee: soprattutto tratti verticali sottili; orizzontali solo per i moduli rimasti; punti per gli isolati
    if (type === "lines") {
      let d = ""; const W = 0.5, o = (1 - W) / 2, R = W / 2;
      const used = new Set();
      for (let c = 0; c < n; c++) { let r = 0; while (r < n) { if (!dark(r, c)) { r++; continue; } let e = r; while (e + 1 < n && dark(e + 1, c)) e++;
        if (e > r) { d += rrect(c + o, r + o, W, e - r + W, [R, R, R, R]); for (let k = r; k <= e; k++) used.add(k * n + c); } r = e + 1; } }
      const free = (r, c) => dark(r, c) && !used.has(r * n + c);
      for (let r = 0; r < n; r++) { let c = 0; while (c < n) { if (!free(r, c)) { c++; continue; } let e = c; while (e + 1 < n && free(r, e + 1)) e++;
        if (e > c) { d += rrect(c + o, r + o, e - c + W, W, [R, R, R, R]); for (let k = c; k <= e; k++) used.add(r * n + k); } c = e + 1; } }
      for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (free(r, c)) d += circ(c + 0.5, r + 0.5, 0.33);
      return d;
    }
    const E = 0.04; // leggera sovrapposizione: niente righine bianche tra moduli vicini
    let d = "";
    if (type === "square") {
      for (let r = 0; r < n; r++) {
        let c = 0;
        while (c < n) {
          if (!dark(r, c)) { c++; continue; }
          let e = c; while (e + 1 < n && dark(r, e + 1)) e++;
          d += `M${c} ${r}h${e - c + 1}v${1 + E}h${-(e - c + 1)}Z`;
          c = e + 1;
        }
      }
      return d;
    }
    for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) {
      if (!dark(r, c)) continue;
      if (type === "dots") { d += circ(c + 0.5, r + 0.5, 0.48); continue; }
      // Connessi: quadratini staccati uniti da ponticelli (formano croci)
      if (type === "connected") {
        d += rrect(c + 0.1, r + 0.1, 0.8, 0.8, [0.08, 0.08, 0.08, 0.08]);
        if (dark(r, c + 1)) d += `M${c + 0.5} ${r + 0.3}h1v0.4h-1Z`;
        if (dark(r + 1, c)) d += `M${c + 0.3} ${r + 0.5}h0.4v1h-0.4Z`;
        continue;
      }
      const T = dark(r - 1, c), B = dark(r + 1, c), L = dark(r, c - 1), R = dark(r, c + 1);
      const k = DOT_R[type] || DOT_R.rounded;
      const rad = [!T && !L ? k[0] : 0, !T && !R ? k[1] : 0, !B && !R ? k[2] : 0, !B && !L ? k[3] : 0];
      d += rrect(c, r, 1 + (R ? E : 0), 1 + (B ? E : 0), rad);
    }
    return d;
  }

  function buildSVG(data, d = S.design, ec = effectiveEC(), opt = {}) {
    const q = matrix(data, ec), n = q.getModuleCount();
    const uid = "q" + Math.random().toString(36).slice(2, 7);
    // Zona del logo (in moduli), centrata e con numero dispari di moduli
    let logoBox = null;
    if (d.logo) {
      // Stesso limite di sicurezza di prima: l'area coperta resta entro ciò che la correzione "Max" recupera
      const side = n * Math.sqrt(d.logoSize * 0.2);
      let hid = Math.ceil(side + (2 * d.logoMargin * n) / 1000);
      if (hid % 2 !== n % 2) hid++;
      hid = Math.min(hid, n - 16);
      const sideFit = Math.min(side, hid - 0.4); // sui QR piccoli il logo resta dentro la zona libera
      const h0 = (n - hid) / 2;
      logoBox = { h0, hid, x: (n - sideFit) / 2, side: sideFit };
    }
    const inFinder = (r, c) => (r < 7 && c < 7) || (r < 7 && c >= n - 7) || (r >= n - 7 && c < 7);
    const inLogo = (r, c) => logoBox && d.hideDots && r >= logoBox.h0 && r < logoBox.h0 + logoBox.hid && c >= logoBox.h0 && c < logoBox.h0 + logoBox.hid;
    const body = bodyPath(q, n, (r, c) => inFinder(r, c) || inLogo(r, c), d.dotType);

    // Colori: una sola sfumatura condivisa da moduli e occhi
    let defs = "", fill = d.dotColor;
    if (d.useGradient) {
      const stops = `<stop offset="0" stop-color="${d.dotColor}"/><stop offset="1" stop-color="${d.gradColor2}"/>`;
      if (d.gradType === "radial") defs = `<radialGradient id="${uid}g" gradientUnits="userSpaceOnUse" cx="${n / 2}" cy="${n / 2}" r="${f2(n / Math.SQRT2)}">${stops}</radialGradient>`;
      else {
        const a = (d.gradRot * Math.PI) / 180, L = n / 2, dx = Math.cos(a) * L, dy = Math.sin(a) * L;
        defs = `<linearGradient id="${uid}g" gradientUnits="userSpaceOnUse" x1="${f2(L - dx)}" y1="${f2(L - dy)}" x2="${f2(L + dx)}" y2="${f2(L + dy)}">${stops}</linearGradient>`;
      }
      fill = `url(#${uid}g)`;
    }
    const outerFill = d.eyesSame ? fill : d.cornerSqColor, innerFill = d.eyesSame ? fill : d.cornerDotColor;
    let eyes = "";
    [[0, 0, 0], [1, 0, 90], [0, 1, -90]].forEach(([cx, cy, rot]) => {
      const x = cx * (n - 7), y = cy * (n - 7);
      const ball = eyeBall(d.cornerDotType, x + 2, y + 2, 1, innerFill);
      const keepUp = d.cornerDotType === "star";
      eyes += `<g transform="rotate(${rot} ${x + 3.5} ${y + 3.5})">${eyeFrame(d.cornerSqType, x, y, 1, outerFill)}${keepUp ? "" : ball}</g>${keepUp ? ball : ""}`;
    });
    let logo = "";
    if (logoBox) {
      const lx = f2(logoBox.x), ls = f2(logoBox.side);
      // Base del colore dello sfondo sotto al logo: separa i loghi scuri dai moduli e aiuta la lettura
      if (d.hideDots) { const pad = 0.15, px0 = f2(logoBox.h0 + pad), ps = f2(logoBox.hid - 2 * pad); logo += `<rect x="${px0}" y="${px0}" width="${ps}" height="${ps}" rx="${f2(Math.min(1.2, ps / 6))}" fill="${d.transparentBg ? "#ffffff" : d.bgColor}"/>`; }
      logo += `<image x="${lx}" y="${lx}" width="${ls}" height="${ls}" preserveAspectRatio="xMidYMid meet" href="${d.logo}" xlink:href="${d.logo}"/>`;
    }
    const inner = `${defs ? `<defs>${defs}</defs>` : ""}<path d="${body}" fill="${fill}"/>${eyes}${logo}`;

    // Miniature: solo il QR, eventualmente ingrandito sull'angolo
    if (opt.crop) return { svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${f2(n * 0.3)} ${f2(n * 0.3)} ${f2(n * 0.4)} ${f2(n * 0.4)}"><rect x="-1" y="-1" width="${n + 2}" height="${n + 2}" fill="#fff"/>${inner}</svg>` };

    const Q = 1000, m = d.margin, sc = (Q - 2 * m) / n;
    const bgFill = d.transparentBg ? null : d.bgColor;
    const place = (x, y) => (bgFill && d.frame === "none" ? `<rect x="${x}" y="${y}" width="${Q}" height="${Q}" fill="${bgFill}"/>` : "") +
      `<g transform="translate(${f2(x + m)} ${f2(y + m)}) scale(${f2(sc)})">${inner}</g>`;
    const bg = d.transparentBg ? "none" : d.bgColor;
    const txt = esc(trim(d.frameText) || " ");
    const fontFam = "'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif";
    const fs = (w, base) => Math.min(base, Math.floor(w / (Math.max(txt.length, 4) * 0.62)));
    let W = Q, H = Q, body2 = "";
    const fc = d.frameColor, tc = d.frameTextColor;
    const text = (x, y, w, color, base) => `<text x="${x}" y="${y}" text-anchor="middle" dominant-baseline="central" font-family="${fontFam}" font-weight="700" font-size="${fs(w, base)}" letter-spacing="2" fill="${color}">${txt}</text>`;
    switch (d.frame) {
      case "none": body2 = place(0, 0); break;
      case "bottom": case "top": {
        const p = 50, band = 200; W = Q + 2 * p; H = Q + 2 * p + band - p;
        const qy = d.frame === "top" ? band : p;
        const ty = d.frame === "top" ? band / 2 + 8 : qy + Q + (H - qy - Q) / 2 - 4;
        body2 = `<rect width="${W}" height="${H}" rx="70" fill="${fc}"/><rect x="${p - 10}" y="${qy - 10}" width="${Q + 20}" height="${Q + 20}" rx="40" fill="${d.transparentBg ? "#ffffff" : d.bgColor}"/>` + place(p, qy) + text(W / 2, ty, W - 120, tc, 120);
        break;
      }
      case "border": {
        const p = 60, band = 190; W = Q + 2 * p; H = Q + p + band;
        body2 = `<rect x="14" y="14" width="${W - 28}" height="${H - 28}" rx="60" fill="${bg}" stroke="${fc}" stroke-width="28"/>` + place(p, p) + text(W / 2, p + Q + band / 2 - 20, W - 160, fc, 110);
        break;
      }
      case "bubble": {
        const gap = 40, bh = 180; W = Q; H = Q + gap + bh;
        body2 = (bg !== "none" ? `<rect width="${Q}" height="${Q}" fill="${bg}"/>` : "") + place(0, 0) +
          `<path d="M${W / 2 - 50} ${Q + gap + 2}L${W / 2} ${Q + gap - 45}L${W / 2 + 50} ${Q + gap + 2}z" fill="${fc}"/><rect x="40" y="${Q + gap}" width="${W - 80}" height="${bh}" rx="${bh / 2}" fill="${fc}"/>` + text(W / 2, Q + gap + bh / 2, W - 200, tc, 100);
        break;
      }
      case "corners": {
        const p = 70, band = 190, L = 170, sw = 26; W = Q + 2 * p; H = Q + 2 * p + band - 30;
        const x2 = W - sw / 2, y2 = Q + 2 * p - sw / 2, o = sw / 2;
        body2 = (bg !== "none" ? `<rect width="${W}" height="${H}" rx="40" fill="${bg}"/>` : "") +
          `<path d="M${o} ${o + L}V${o + 30}Q${o} ${o} ${o + 30} ${o}H${o + L}M${x2 - L} ${o}H${x2 - 30}Q${x2} ${o} ${x2} ${o + 30}V${o + L}M${x2} ${y2 - L}V${y2 - 30}Q${x2} ${y2} ${x2 - 30} ${y2}H${x2 - L}M${o + L} ${y2}H${o + 30}Q${o} ${y2} ${o} ${y2 - 30}V${y2 - L}" fill="none" stroke="${fc}" stroke-width="${sw}" stroke-linecap="round"/>` +
          place(p, p) + text(W / 2, Q + 2 * p + (band - 30) / 2, W - 120, fc, 110);
        break;
      }
    }
    return { svg: `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${body2}</svg>`, W, H };
  }

  function svgToPng(svg, W, H, targetW) {
    return new Promise((res, rej) => {
      const img = new Image();
      const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
      img.onload = () => {
        const draw = () => {
          const s = targetW / W, c = document.createElement("canvas");
          c.width = Math.round(W * s); c.height = Math.round(H * s);
          const ctx = c.getContext("2d"); ctx.imageSmoothingQuality = "high";
          ctx.drawImage(img, 0, 0, c.width, c.height);
          URL.revokeObjectURL(url);
          res(c.toDataURL("image/png"));
        };
        // Safari (PowerPoint per Mac) a volte non ha ancora decodificato il logo annidato nell'SVG
        if (svg.includes("<image")) setTimeout(draw, 80); else draw();
      };
      img.onerror = () => { URL.revokeObjectURL(url); rej(new Error("render")); };
      img.src = url;
    });
  }

  /* ---------- Aggiornamento anteprima ---------- */
  let timer, saveT, current = null, previewUrl = null;
  function update() {
    syncDesignUI();
    clearTimeout(timer); timer = setTimeout(render, 60);
    clearTimeout(saveT); saveT = setTimeout(saveState, 600); // salvataggio differito: niente scatti mentre scrivi
  }

  function render() {
    const data = encode(S.type, vals());
    const empty = !data;
    const warns = [];
    const d = S.design;
    let out;
    try {
      out = buildSVG(empty ? "https://example.com" : data);
    } catch (e) {
      current = null; setReady(false);
      $("#warnBox").innerHTML = ""; $("#contrastWarn").innerHTML = ""; $("#preview").classList.add("dim");
      const msg = String(e && (e.message || e));
      $("#status").textContent = /overflow|length/i.test(msg) ? t("status.tooLong") : t("status.error");
      return;
    }
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = URL.createObjectURL(new Blob([out.svg], { type: "image/svg+xml" }));
    $("#previewImg").src = previewUrl; $("#barImg").src = previewUrl;
    $("#preview").classList.toggle("dim", empty);
    current = empty ? null : Object.assign({ data }, out);
    setReady(!empty);
    prepShare();
    if (empty) { $("#status").textContent = t("status.empty"); }
    else {
      const bytes = new Blob([data]).size;
      $("#status").textContent = t("status.ready", { n: bytes });
      if (bytes > 300) warns.push(t("warn.long"));
    }
    const bgC = d.transparentBg ? "#ffffff" : d.bgColor;
    const cols = [d.dotColor].concat(d.useGradient ? [d.gradColor2] : []).concat(d.eyesSame ? [] : [d.cornerSqColor, d.cornerDotColor]);
    const minC = Math.min.apply(null, cols.map((c) => contrast(c, bgC)));
    if (minC < 3) warns.push(t("warn.contrast"));
    else if (hexLum(d.dotColor) > hexLum(bgC)) warns.push(t("warn.inverted"));
    if (d.transparentBg) warns.push(t("warn.transparent"));
    if (d.logo && d.logoSize > 0.38) warns.push(t("warn.bigLogo"));
    // Moduli sottili + occhio molto decorativo: combinazione delicata per alcuni lettori
    if (["lines", "dots"].includes(d.dotType) && ["star", "plus", "diamond", "dots9"].includes(d.cornerDotType)) warns.push(t("warn.fragile"));
    if (d.margin < 10 && d.frame === "none") warns.push(t("warn.margin"));
    const html = warns.map((w) => `<div class="warn">${esc(w)}</div>`).join("");
    $("#warnBox").innerHTML = html;
    $("#contrastWarn").innerHTML = minC < 3 ? '<div class="warn">' + esc(t("warn.contrastShort")) + '</div>' : "";
  }

  function setReady(ok) {
    ["#insertBtn", "#barInsert", "#dlPng", "#dlSvg"].forEach((s) => ($(s).disabled = !ok));
  }

  /* ---------- Inserimento e download ---------- */
  const CM_TO_PT = 28.3465;
  async function insert() {
    if (!current) return;
    const wPt = Math.min(30, Math.max(1, +S.sizeCm || 5)) * CM_TO_PT;
    const hPt = (wPt * current.H) / current.W;
    if (window.QR_DEMO) { toast(t("toast.demo")); addRecent(); return; }
    if (window.QR_WEB) { if (shareMode) { share(); return; } download("png"); toast(t("toast.downloaded")); return; }
    if (!inOffice) {
      toast(t("toast.browser"));
      download("png"); return;
    }
    const btns = ["#insertBtn", "#barInsert"].map((s) => $(s));
    btns.forEach((b) => (b.disabled = true));
    try {
      if (S.format === "svg" && Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
        await setData(current.svg, { coercionType: Office.CoercionType.XmlSvg, imageWidth: wPt, imageHeight: hPt });
      } else {
        if (S.format === "svg") toast(t("toast.noSvg"));
        const px = Math.max(600, Math.min(1800, Math.round(((+S.sizeCm || 5) / 2.54) * 300))); // 300 dpi
        const png = await svgToPng(current.svg, current.W, current.H, px);
        await setData(png.split(",")[1], { coercionType: Office.CoercionType.Image, imageWidth: wPt, imageHeight: hPt });
      }
      toast(t("toast.inserted"));
      addRecent();
    } catch (e) {
      toast(t("toast.insertFail"));
    } finally { btns.forEach((b) => (b.disabled = !current)); }
  }
  function setData(val, opts) {
    return new Promise((res, rej) => Office.context.document.setSelectedDataAsync(val, opts, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error))));
  }
  function fileBase() { return "qr-" + S.type + "-" + new Date().toISOString().slice(0, 10); }
  /* ---------- Condivisione da cellulare (versione web) ----------
     Sul telefono il pulsante principale apre il menu Condividi del sistema:
     da lì si salva nelle Foto (iPhone: "Salva immagine") o si invia con WhatsApp, Messaggi, Telegram…
     Il PNG viene preparato in anticipo, perché Safari apre il menu solo se la chiamata parte subito dal tocco. */
  const MOBILE = window.QR_WEB && (matchMedia("(pointer: coarse)").matches || /Android|iPhone|iPad|iPod/i.test(navigator.userAgent));
  let shareMode = false, shareFile = null, shareFor = null, shareT;
  function pngPx() { return Math.max(600, Math.min(3000, Math.round(((+S.sizeCm || 5) / 2.54) * 300))); }
  function detectShare() {
    try {
      if (!MOBILE || !navigator.share || !navigator.canShare) return false;
      return navigator.canShare({ files: [new File([new Uint8Array([137, 80, 78, 71])], "t.png", { type: "image/png" })] });
    } catch (e) { return false; }
  }
  function prepShare() {
    if (!shareMode) return;
    clearTimeout(shareT); shareFile = null; shareFor = null;
    if (!current) return;
    const cur = current, px = pngPx();
    shareT = setTimeout(async () => {
      try {
        const url = await svgToPng(cur.svg, cur.W, cur.H, px);
        const blob = await (await fetch(url)).blob();
        if (current === cur) { shareFile = new File([blob], fileBase() + ".png", { type: "image/png" }); shareFor = cur.svg + "|" + px; }
      } catch (e) { /* al tocco lo rigenero */ }
    }, 350);
  }
  async function share() {
    if (!current) return;
    let file = shareFile;
    if (!file || shareFor !== current.svg + "|" + pngPx()) {
      const url = await svgToPng(current.svg, current.W, current.H, pngPx());
      file = new File([await (await fetch(url)).blob()], fileBase() + ".png", { type: "image/png" });
    }
    try {
      await navigator.share({ files: [file] });
      addRecent();
    } catch (e) {
      if (e && e.name === "AbortError") return; // menu chiuso dall'utente
      download("png"); toast(t("toast.downloaded"));
    }
  }

  async function download(kind) {
    if (!current) return;
    let href;
    if (kind === "svg") href = URL.createObjectURL(new Blob([current.svg], { type: "image/svg+xml" }));
    else href = await svgToPng(current.svg, current.W, current.H, window.QR_WEB ? pngPx() : 2000);
    const a = document.createElement("a"); a.href = href; a.download = fileBase() + "." + kind;
    document.body.appendChild(a); a.click(); a.remove();
    if (kind === "svg") setTimeout(() => URL.revokeObjectURL(href), 2000);
    addRecent();
  }

  /* ---------- Recenti ---------- */
  async function addRecent() {
    if (!current) return;
    // Per il Wi-Fi salvo una versione senza password: dati, campi e anche la miniatura (che è un QR leggibile)
    const isWifi = S.type === "wifi";
    const values = isWifi ? noPass(vals()) : JSON.parse(JSON.stringify(vals()));
    const data = isWifi ? encode("wifi", values) : current.data;
    let thumb;
    try {
      const src = isWifi ? buildSVG(data) : current;
      thumb = await svgToPng(src.svg, src.W, src.H, 96);
    } catch (e) { return; }
    const design = Object.assign({}, S.design);
    if (design.logo && design.logo.length > 150000) { design.logo = null; design.logoName = ""; }
    let list = store.get("qrstudio.recent", []);
    list = list.filter((r) => r.data !== data);
    list.unshift({ data, type: S.type, values, design, thumb });
    store.set("qrstudio.recent", list.slice(0, 8));
    rememberDesign(S.design);
    renderRecent(); renderMyPresets();
  }
  function renderRecent() {
    const list = store.get("qrstudio.recent", []);
    const box = $("#recent");
    if (!list.length) { box.innerHTML = '<span class="recent-empty">' + esc(t("recent.empty")) + '</span>'; return; }
    box.innerHTML = list.map((r, i) => `<button type="button" data-i="${i}" title="${esc(r.data.slice(0, 80))}"><img alt="${esc(t("recent.alt"))}" src="${esc(r.thumb)}"></button>`).join("");
    $$("#recent button").forEach((b) => b.addEventListener("click", () => {
      const r = list[+b.dataset.i];
      S.type = r.type; S.values[r.type] = r.values; loadDesign(r.design);
      renderTypes(); renderFields(); update(); toast(t("toast.restored"));
    }));
  }

  /* ---------- Avvio ---------- */
  function labelButtons() {
    if (window.QR_WEB && shareMode) { $("#insertLabel").textContent = t("btn.share"); $("#barInsert").textContent = t("btn.shareShort"); $("#dlPngLabel").textContent = "PNG"; $("#dlSvgLabel").textContent = "SVG"; return; }
    if (window.QR_WEB) { $("#insertLabel").textContent = t("btn.downloadPng"); $("#barInsert").textContent = t("btn.downloadPng"); $("#dlSvgLabel").textContent = t("btn.downloadSvg"); return; }
    if (window.QR_DEMO) { $("#insertLabel").textContent = t("btn.insertDemo"); $("#barInsert").textContent = t("btn.insert"); }
    else if (!inOffice) { $("#insertLabel").textContent = t("btn.insertBrowser"); $("#barInsert").textContent = t("btn.downloadPng"); }
    else { $("#insertLabel").textContent = t("btn.insert"); $("#barInsert").textContent = t("btn.insert"); }
    // In PowerPoint i pulsanti di download non funzionano su tutte le piattaforme (es. Mac): li nascondo, resta l'inserimento
    if (inOffice) $(".btns").hidden = true;
  }
  function applyLang() {
    applyStatic(); labelButtons();
    $("#langSel").value = LANG;
    renderTypes(); renderFields(); syncDesignUI(); renderRecent(); renderMyPresets(); render();
  }
  function setLang(l) {
    if (!LANGS.includes(l) || l === LANG) return;
    // Se il testo della cornice è quello predefinito, lo traduco; se l'hai scritto tu, resta com'è
    if (FRAME_DEFAULTS().includes(S.design.frameText)) S.design.frameText = I18N[l]["frame.default"];
    LANG = l; store.set("qrlab.lang", l); saveState();
    applyLang();
  }

  function init() {
    LANG = detectLang();
    if (window.QR_DEMO) $("#appSub").dataset.i18n = "app.subDemo";
    if (FRAME_DEFAULTS().includes(S.design.frameText)) S.design.frameText = t("frame.default");
    $("#langSel").innerHTML = I18N.langs.map(([c, n]) => `<option value="${c}">${n}</option>`).join("");
    $("#langSel").addEventListener("change", (e) => setLang(e.target.value));
    buildDesignUI();
    $$("#format button").forEach((b) => b.addEventListener("click", () => { S.format = b.dataset.v; update(); }));
    $("#sizeCm").addEventListener("input", (e) => { S.sizeCm = +e.target.value || 5; saveState(); });
    $("#insertBtn").addEventListener("click", insert);
    $("#barInsert").addEventListener("click", insert);
    $("#barImg").addEventListener("click", () => $("#previewCard").scrollIntoView({ behavior: "smooth" }));
    $("#dlPng").addEventListener("click", () => download("png"));
    $("#dlSvg").addEventListener("click", () => download("svg"));
    // Link in fondo al pannello.
    // Versione web: "Richiedi una consulenza" (email). Add-in: "Aiuto e supporto", perché le regole di
    // Microsoft Marketplace non consentono di indirizzare gli utenti verso servizi esterni dall'add-in.
    const cta = $("#ctaMail");
    if (!window.QR_WEB && !window.QR_DEMO) {
      cta.dataset.i18n = "cta.help";
      cta.href = new URL("support.html", location.href).href;
      cta.target = "_blank"; cta.rel = "noopener"; cta.removeAttribute("title");
    }
    cta.addEventListener("click", (e) => {
      if (!inOffice) return; // nel browser il link funziona normalmente
      try {
        // nel pannello di PowerPoint i link vanno aperti dal sistema
        if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) { e.preventDefault(); Office.context.ui.openBrowserWindow(cta.href); }
      } catch (err) { /* lascio agire il link normale */ }
    });
    if (window.QR_DEMO) $("#dlPng").parentElement.hidden = true;
    if (window.QR_WEB) { // versione web: download al posto dell'inserimento
      $("#appSub").dataset.i18n = "app.subWeb";
      $("#widthLbl").dataset.i18n = "lbl.printWidth";
      $("#formatBox").hidden = true;
      shareMode = detectShare();
      if (shareMode) { // telefono: Condividi / Salva nelle Foto, con PNG e SVG come alternative
        $("#insertBtn svg").innerHTML = '<path d="M12 15V4M8 8l4-4 4 4"/><path d="M6 11H5v9h14v-9h-1"/>';
        $("#webHint").dataset.i18n = "web.hintMobile";
      } else {
        $("#dlPng").style.display = "none";
        $("#insertBtn svg").innerHTML = '<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>';
        $(".btns").style.gridTemplateColumns = "1fr";
      }
      $("#webHint").hidden = false;
    }
    // Le icone salvate come SVG e il logo Ultraspeaker vengono rigenerati (il logo del marchio ora è vettoriale)
    if (S.design.logoName && S.design.logoName.startsWith("icon:") && (/^data:image\/svg/.test(S.design.logo || "") || S.design.logoName === "icon:brand")) setIconLogo();
    applyLang();
  }

  window.__qrLab = { svg: () => (current ? current.svg : null) }; // per diagnostica

  let started = false;
  const start = () => { if (!started) { started = true; init(); } };
  // PowerPoint ha risposto dopo che il pannello era già partito (computer lento)
  function officeArrivedLate() {
    if (!store.get("qrlab.lang", null)) {
      const l = detectLang();
      if (l !== LANG) { if (FRAME_DEFAULTS().includes(S.design.frameText)) S.design.frameText = I18N[l]["frame.default"]; LANG = l; }
    }
    applyLang();
  }
  if (window.Office && Office.onReady) {
    Office.onReady((info) => {
      const host = !!(info && info.host);
      if (!started) { inOffice = host; start(); }
      else if (host && !inOffice) { inOffice = true; officeArrivedLate(); }
    });
    setTimeout(start, 2500); // fallback se Office.js non risponde (es. aperto nel browser)
  } else {
    document.readyState === "loading" ? document.addEventListener("DOMContentLoaded", start) : start();
  }
})();
