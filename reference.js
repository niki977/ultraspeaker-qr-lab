/* The Ultraspeaker QR Lab – sezione Reference (citazioni scientifiche + QR)
   Usa il motore QR del QR Lab tramite window.QRLab (definito in app.js).
   Metadati da API pubbliche senza chiave: Crossref, NCBI E-utilities (PubMed), Unpaywall. Nessun backend. */
(function () {
  "use strict";

  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const L = () => window.QRLab;
  const t = (k, v) => L().t(k, v);
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const CM_TO_PT = 28.3465;
  const CONTACT = "info@theultraspeaker.com"; // contatto dell'app richiesto dalle API (non è un dato dell'utente)
  const NAME_PREFIX = "UltraspeakerRef:";

  /* =====================================================================
     1. RETE: richieste con gestione di offline, timeout, 404 e limiti di frequenza
     ===================================================================== */
  class RefError extends Error { constructor(code) { super(code); this.code = code; } }
  let ncbiLast = 0;
  async function ncbiSlot() { // NCBI senza chiave: massimo 3 richieste al secondo
    const wait = ncbiLast + 380 - Date.now();
    ncbiLast = Math.max(Date.now(), ncbiLast + 380);
    if (wait > 0) await sleep(wait);
  }
  async function fetchApi(url, kind = "json", retries = 2) {
    if (navigator.onLine === false) throw new RefError("offline");
    for (let i = 0; ; i++) {
      const ctl = typeof AbortController !== "undefined" ? new AbortController() : null;
      const tm = setTimeout(() => ctl && ctl.abort(), 15000);
      let r;
      try { r = await fetch(url, ctl ? { signal: ctl.signal } : {}); }
      catch (e) {
        clearTimeout(tm);
        if (e && e.name === "AbortError") throw new RefError("timeout");
        throw new RefError(navigator.onLine === false ? "offline" : "network");
      }
      clearTimeout(tm);
      if (r.status === 429 || r.status === 503) {
        if (i < retries) { await sleep(1200 * (i + 1)); continue; }
        throw new RefError("rate");
      }
      if (r.status === 404) throw new RefError("notfound");
      if (!r.ok) throw new RefError("http");
      try { return kind === "text" ? await r.text() : await r.json(); }
      catch (e) { throw new RefError("http"); }
    }
  }
  const CR = "https://api.crossref.org/works";
  const EU = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/";
  const EU_ID = "&tool=UltraspeakerQRLab&email=" + encodeURIComponent(CONTACT);
  const crossref = (doi) => fetchApi(`${CR}/${encodeURIComponent(doi)}?mailto=${encodeURIComponent(CONTACT)}`);
  const crossrefSearch = (q) => fetchApi(`${CR}?query.bibliographic=${encodeURIComponent(q)}&rows=8&mailto=${encodeURIComponent(CONTACT)}`);
  async function eutils(path, params, kind = "json") { await ncbiSlot(); return fetchApi(`${EU}${path}?${params}${EU_ID}`, kind); }
  const unpaywall = (doi) => fetchApi(`https://api.unpaywall.org/v2/${encodeURIComponent(doi)}?email=${encodeURIComponent(CONTACT)}`);

  /* =====================================================================
     2. RICONOSCIMENTO DELL'INPUT E NORMALIZZAZIONE DEI METADATI
     ===================================================================== */
  const DOI_RE = /10\.\d{4,9}\/[^\s"<>]+/i;
  function cleanDoi(d) {
    d = d.replace(/[.,;:\]}>]+$/, "");
    // una ")" finale si toglie solo se non chiude una parentesi del DOI, es. 10.1016/S0140-6736(97)11096-0
    while (d.endsWith(")") && (d.match(/\(/g) || []).length < (d.match(/\)/g) || []).length) d = d.slice(0, -1).replace(/[.,;:]+$/, "");
    return d.toLowerCase();
  }
  function parseInput(raw) {
    const s = String(raw || "").trim();
    if (!s) return null;
    let m = s.match(/pubmed\.ncbi\.nlm\.nih\.gov\/(\d{1,9})/i) || s.match(/ncbi\.nlm\.nih\.gov\/pubmed\/(\d{1,9})/i);
    if (m) return { kind: "pmid", id: m[1] };
    let dec = s; try { dec = decodeURIComponent(s); } catch (e) { /* lascio com'è */ }
    m = dec.match(DOI_RE);
    if (m) return { kind: "doi", id: cleanDoi(m[0]) };
    m = s.match(/^(?:pmid\s*:?\s*)?(\d{1,9})$/i);
    if (m) return { kind: "pmid", id: m[1] };
    if (s.length < 3) return null;
    return { kind: "text", q: s };
  }
  // Elenco incollato: DOI, PMID o link, uno per riga (o separati da virgole / punti e virgola).
  // I numeri dentro una frase (es. un anno) NON vengono presi per PMID: un PMID deve stare da solo o dopo "PMID".
  function parseList(raw) {
    const out = [], seen = new Set();
    const push = (p) => { const k = p.kind + ":" + p.id; if (!seen.has(k)) { seen.add(k); out.push(p); } };
    String(raw || "").split(/[\n\r;,\t]+/).map((x) => x.trim()).filter(Boolean).forEach((chunk) => {
      let dec = chunk; try { dec = decodeURIComponent(chunk); } catch (e) { /* lascio com'è */ }
      const dois = dec.match(new RegExp(DOI_RE.source, "gi")) || [];
      const pms = [...chunk.matchAll(/(?:pubmed\.ncbi\.nlm\.nih\.gov\/|ncbi\.nlm\.nih\.gov\/pubmed\/|\bpmid\s*:?\s*)(\d{1,9})/gi)].map((m) => m[1]);
      dois.forEach((d) => push({ kind: "doi", id: cleanDoi(d) }));
      pms.forEach((id) => push({ kind: "pmid", id }));
      if (!dois.length && !pms.length) chunk.split(/\s+/).forEach((w) => { if (/^\d{1,9}$/.test(w) && chunk.split(/\s+/).every((x) => /^\d{1,9}$/.test(x))) push({ kind: "pmid", id: w }); });
    });
    return out;
  }
  const decodeEnt = (s) => { const el = document.createElement("textarea"); el.innerHTML = s; return el.value; };
  const cleanText = (s) => decodeEnt(String(s || "").replace(/<[^>]+>/g, "")).replace(/\s+/g, " ").trim();
  const initials = (g) => String(g || "").split(/[\s.\-]+/).filter(Boolean).map((w) => w[0].toUpperCase()).join("");
  const uid = () => "r" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  const RETRACT = ["retraction", "withdrawal", "removal", "retracted"];
  const CORRECT = ["correction", "erratum", "corrigendum", "addendum", "clarification", "new_version"];
  const CONCERN = ["expression_of_concern", "expression-of-concern"];
  function statusFrom(notices, retractedFlag) {
    const st = { retracted: !!retractedFlag, corrected: false, concern: false, notices: [] };
    (notices || []).forEach((n) => {
      const ty = String(n.type || "").toLowerCase();
      if (RETRACT.includes(ty)) st.retracted = true;
      else if (CONCERN.includes(ty)) st.concern = true;
      else if (CORRECT.includes(ty)) st.corrected = true;
      else return;
      if (n.doi && !st.notices.some((x) => x.doi === n.doi)) st.notices.push({ type: ty, doi: String(n.doi).toLowerCase() });
    });
    return st;
  }
  function fromCrossref(m) {
    const authors = (m.author || []).map((a) => (a.family ? { family: cleanText(a.family), given: cleanText(a.given || "") } : { family: cleanText(a.name || ""), given: "", org: true })).filter((a) => a.family);
    const dp = ["issued", "published-print", "published-online", "published"].map((k) => m[k] && m[k]["date-parts"] && m[k]["date-parts"][0]).find((x) => x && x[0]) || [];
    const notices = (m["updated-by"] || []).map((u) => ({ type: u.type, doi: u.DOI }));
    return {
      doi: m.DOI ? String(m.DOI).toLowerCase() : "", pmid: "",
      authors, title: cleanText((m.title || [])[0]), journal: cleanText((m["container-title"] || [])[0]),
      journalAbbr: cleanText((m["short-container-title"] || [])[0]), year: dp[0] ? String(dp[0]) : "",
      volume: cleanText(m.volume), issue: cleanText(m.issue), pages: cleanText(m.page || m["article-number"] || ""),
      status: statusFrom(notices),
    };
  }
  // PubMed scrive i nomi completi delle riviste in minuscolo ("Journal of dentistry"): li porto in Title Case per APA/Harvard
  const MINOR = /^(of|and|the|in|for|on|a|an|to|at|by|de|der|und|la|le|et|di|e|y|del)$/i;
  const titleCase = (s) => String(s || "").split(" ").map((w, i) => (i > 0 && MINOR.test(w) ? w.toLowerCase() : w.charAt(0).toUpperCase() + w.slice(1))).join(" ");
  function fromPubmed(s) {
    const authors = (s.authors || []).filter((a) => a.name && a.authtype !== "Investigator").map((a) => {
      const m = String(a.name).match(/^(.*?)\s+([A-Z]{1,5})$/);
      return m && a.authtype !== "CollectiveName" ? { family: m[1], given: m[2], init: true } : { family: a.name, given: "", org: true };
    });
    const ids = s.articleids || [];
    const idOf = (ty) => { const x = ids.find((i) => i.idtype === ty); return x ? x.value : ""; };
    const year = (String(s.pubdate || s.sortpubdate || "").match(/\d{4}/) || [""])[0];
    const pubtypes = (s.pubtype || []).map((p) => String(p).toLowerCase());
    return {
      doi: (idOf("doi") || (String(s.elocationid || "").match(DOI_RE) || [""])[0]).toLowerCase(),
      pmid: String(s.uid || idOf("pubmed") || ""), pmcid: idOf("pmc") || idOf("pmcid") || "",
      authors, title: cleanText(s.title).replace(/\.$/, ""), journal: titleCase(cleanText(s.fulljournalname)), journalAbbr: cleanText(s.source),
      year, volume: cleanText(s.volume), issue: cleanText(s.issue), pages: cleanText(s.pages || ""),
      status: statusFrom([], pubtypes.includes("retracted publication")),
    };
  }
  async function pubmedSummary(ids) {
    const j = await eutils("esummary.fcgi", "db=pubmed&retmode=json&id=" + ids.join(","));
    const res = (j && j.result) || {};
    return (res.uids || ids).map((u) => res[u]).filter((x) => x && !x.error);
  }
  async function pmidForDoi(doi) {
    try {
      const j = await eutils("esearch.fcgi", "db=pubmed&retmode=json&retmax=1&term=" + encodeURIComponent(doi + "[doi]"));
      return ((j.esearchresult || {}).idlist || [])[0] || "";
    } catch (e) { return ""; }
  }
  // Recupera un riferimento completo da DOI o PMID, con stato (ritrattato/corretto) e open access
  async function lookup(p) {
    let ref;
    if (p.kind === "doi") {
      const j = await crossref(p.id);
      ref = fromCrossref(j.message || {});
      ref.doi = ref.doi || p.id;
      const pmid = await pmidForDoi(ref.doi);
      if (pmid) {
        ref.pmid = pmid;
        try { // per le riviste biomediche le abbreviazioni NLM di PubMed sono quelle corrette per Vancouver/AMA
          const s = (await pubmedSummary([pmid]))[0];
          if (s) { const pm = fromPubmed(s); ref.journalAbbr = pm.journalAbbr || ref.journalAbbr; ref.pmcid = pm.pmcid; if (pm.status.retracted) ref.status.retracted = true; }
        } catch (e) { /* PubMed facoltativo */ }
      }
    } else if (p.kind === "pmid") {
      const s = (await pubmedSummary([p.id]))[0];
      if (!s) throw new RefError("notfound");
      ref = fromPubmed(s);
      if (ref.doi) {
        try {
          const cr = fromCrossref((await crossref(ref.doi)).message || {});
          ref.status = statusFrom(cr.status.notices, ref.status.retracted || cr.status.retracted);
          if (!ref.issue) ref.issue = cr.issue;
          if (cr.journal) ref.journal = cr.journal; // nome ufficiale della rivista
        } catch (e) { /* Crossref facoltativo */ }
      }
    } else throw new RefError("notfound");
    if (!ref.title && !ref.authors.length) throw new RefError("notfound");
    await addOpenAccess(ref);
    ref.id = uid(); ref.added = Date.now();
    return ref;
  }
  async function addOpenAccess(ref) {
    ref.oaUrl = ""; ref.oaChecked = false;
    if (ref.doi) {
      try {
        const u = await unpaywall(ref.doi);
        const b = u && u.best_oa_location;
        if (u && u.is_oa && b) ref.oaUrl = b.url_for_landing_page || b.url || b.url_for_pdf || "";
        ref.oaChecked = true;
      } catch (e) { /* Unpaywall facoltativo */ }
    }
    if (!ref.oaUrl && ref.pmcid) ref.oaUrl = "https://pmc.ncbi.nlm.nih.gov/articles/" + ref.pmcid + "/";
  }
  async function searchText(q) {
    try {
      const j = await eutils("esearch.fcgi", "db=pubmed&retmode=json&retmax=10&sort=relevance&term=" + encodeURIComponent(q));
      const ids = (j.esearchresult || {}).idlist || [];
      if (ids.length) return (await pubmedSummary(ids)).map((s) => Object.assign(fromPubmed(s), { src: "pubmed" }));
    } catch (e) { if (e.code === "offline" || e.code === "rate") throw e; }
    const j = await crossrefSearch(q);
    return ((j.message || {}).items || []).filter((m) => (m.title || [])[0]).map((m) => Object.assign(fromCrossref(m), { src: "crossref" }));
  }
  async function fetchAbstract(pmid) {
    const xml = await eutils("efetch.fcgi", "db=pubmed&rettype=abstract&retmode=xml&id=" + pmid, "text");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    return $$("AbstractText", doc).map((n) => (n.getAttribute("Label") ? n.getAttribute("Label") + ": " : "") + n.textContent.replace(/\s+/g, " ").trim()).join("\n\n");
  }

  /* =====================================================================
     3. STILI DI CITAZIONE
     ===================================================================== */
  const STYLES = ["vancouver", "apa", "ama", "harvard"];
  const SHORTS = ["authorJournal", "numeric", "authorYear"];
  const STYLE_NAME = { vancouver: "Vancouver", apa: "APA 7", ama: "AMA", harvard: "Harvard" };
  const NUMBERED = (style) => style === "vancouver" || style === "ama";
  const endDot = (s) => (s && !/[.?!]$/.test(s) ? s + "." : s);
  const dash = (p) => String(p || "").replace(/\s*[-–]\s*/g, "–");
  const hyph = (p) => String(p || "").replace(/\s*[–-]\s*/g, "-");
  const nlmName = (a) => (a.org ? a.family : (a.family + " " + (a.init ? a.given : initials(a.given))).trim());
  const apaName = (a) => { if (a.org) return a.family; const i = (a.init ? a.given.split("") : initials(a.given).split("")).map((c) => c + ".").join(" "); return i ? `${a.family}, ${i}` : a.family; };
  const doiUrl = (r) => (r.doi ? "https://doi.org/" + r.doi : r.pmid ? "https://pubmed.ncbi.nlm.nih.gov/" + r.pmid + "/" : "");
  function srcPart(r) { // ;vol(issue):pagine
    let s = r.year || "";
    if (r.volume) s += ";" + r.volume;
    if (r.issue) s += "(" + r.issue + ")";
    if (r.pages) s += ":" + hyph(r.pages);
    return s;
  }
  function formatFull(r, style) {
    const A = r.authors || [];
    const jAbbr = r.journalAbbr || r.journal, jFull = r.journal || r.journalAbbr;
    if (style === "vancouver" || style === "ama") {
      let names;
      if (style === "vancouver") names = A.length > 6 ? A.slice(0, 6).map(nlmName).join(", ") + ", et al" : A.map(nlmName).join(", ");
      else names = A.length > 6 ? A.slice(0, 3).map(nlmName).join(", ") + ", et al" : A.map(nlmName).join(", ");
      const parts = [];
      if (names) parts.push(endDot(names));
      if (r.title) parts.push(endDot(r.title));
      if (jAbbr) parts.push(endDot(jAbbr));
      const sp = srcPart(r); if (sp) parts.push(endDot(sp));
      if (r.doi) parts.push(style === "ama" ? "doi:" + r.doi : "doi: " + r.doi);
      else if (r.pmid) parts.push("PMID: " + r.pmid);
      return parts.join(" ");
    }
    const yr = r.year || "n.d.";
    if (style === "apa") {
      const n = A.map(apaName);
      let names;
      if (n.length === 0) names = "";
      else if (n.length === 1) names = n[0];
      else if (n.length <= 20) names = n.slice(0, -1).join(", ") + ", & " + n[n.length - 1];
      else names = n.slice(0, 19).join(", ") + ", . . . " + n[n.length - 1];
      let s = names ? `${endDot(names)} (${yr}). ` : "";
      s += r.title ? endDot(r.title) + " " : "";
      if (!names) s += `(${yr}). `;
      let src = jFull || "";
      if (r.volume) src += (src ? ", " : "") + r.volume + (r.issue ? `(${r.issue})` : "");
      if (r.pages) src += (src ? ", " : "") + dash(r.pages);
      if (src) s += endDot(src) + " ";
      if (doiUrl(r)) s += doiUrl(r);
      return s.trim();
    }
    // Harvard (Cite Them Right)
    const n = A.map(apaName);
    let names;
    if (n.length === 0) names = "";
    else if (n.length === 1) names = n[0];
    else if (n.length <= 3) names = n.slice(0, -1).join(", ") + " and " + n[n.length - 1];
    else names = n[0] + " et al.";
    let s = names ? `${names} (${yr}) ` : `(${yr}) `;
    s += r.title ? `‘${r.title}’, ` : "";
    s += jFull || "";
    if (r.volume) s += ", " + r.volume + (r.issue ? `(${r.issue})` : "");
    if (r.pages) s += ", " + (/[–-]/.test(r.pages) ? "pp. " : "p. ") + dash(r.pages);
    s = endDot(s.replace(/,\s*$/, ""));
    if (doiUrl(r)) s += ` ${r.doi ? "doi: " + r.doi : "Available at: " + doiUrl(r)}.`;
    return s;
  }
  function firstAuthors(r) {
    const A = r.authors || [];
    if (!A.length) return "";
    if (A.length === 1) return A[0].family;
    if (A.length === 2) return A[0].family + " & " + A[1].family;
    return A[0].family + " et al.";
  }
  function formatShort(r, fmt, num) {
    if (fmt === "numeric") return "[" + (num || "?") + "]";
    const who = firstAuthors(r) || r.journalAbbr || r.journal || "";
    if (fmt === "authorYear") return `(${who}${r.year ? ", " + r.year : ""})`;
    const j = r.journalAbbr || r.journal || "";
    return [who, [j, r.year].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  }

  /* =====================================================================
     4. ARCHIVIO: impostazioni del documento (restano nel .pptx) o browser
     ===================================================================== */
  const DOC_KEY = "ultraspeakerReferences";
  const LOCAL_KEY = "qrlab.refs";
  const PREF_KEY = "qrlab.refPrefs";
  const docSettings = () => { try { return L().isOffice() && Office.context.document.settings ? Office.context.document.settings : null; } catch (e) { return null; } };
  const emptyDoc = () => ({ v: 1, refs: [], style: "vancouver", shortFmt: "authorJournal" });
  let DOC = emptyDoc();
  function loadDoc() {
    let d = null;
    const ds = docSettings();
    if (ds) { try { d = ds.get(DOC_KEY); } catch (e) { d = null; } }
    else d = L().store.get(LOCAL_KEY, null);
    if (typeof d === "string") { try { d = JSON.parse(d); } catch (e) { d = null; } }
    DOC = Object.assign(emptyDoc(), d && typeof d === "object" ? d : {});
    if (!Array.isArray(DOC.refs)) DOC.refs = [];
    if (!STYLES.includes(DOC.style)) DOC.style = "vancouver";
    if (!SHORTS.includes(DOC.shortFmt)) DOC.shortFmt = "authorJournal";
  }
  let saveT;
  function saveDoc() {
    const ds = docSettings();
    if (!ds) { L().store.set(LOCAL_KEY, DOC); return; }
    try { ds.set(DOC_KEY, DOC); } catch (e) { return; }
    clearTimeout(saveT);
    saveT = setTimeout(() => { try { ds.saveAsync(() => {}); } catch (e) { /* ignora */ } }, 300);
  }
  const numOf = (ref) => DOC.refs.indexOf(ref) + 1;
  const normTitle = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  function sameRef(a, b) {
    if (a === b) return false;
    if (a.doi && b.doi) return a.doi === b.doi;
    if (a.pmid && b.pmid) return a.pmid === b.pmid;
    return !!a.title && normTitle(a.title) === normTitle(b.title) && (a.year || "") === (b.year || "");
  }
  const findSame = (r) => DOC.refs.find((x) => x !== r && (x.id === r.id || sameRef(x, r)));
  const dupIds = () => { const s = new Set(); DOC.refs.forEach((a, i) => DOC.refs.forEach((b, j) => { if (i < j && sameRef(a, b)) { s.add(a.id); s.add(b.id); } })); return s; };
  function addToDoc(ref) { // restituisce il riferimento già presente se è un duplicato
    const ex = findSame(ref);
    if (ex) return { ref: ex, dup: true };
    const c = JSON.parse(JSON.stringify(ref)); delete c.src; delete c.abstract;
    DOC.refs.push(c); saveDoc();
    return { ref: c, dup: false };
  }

  const defaultPrefs = { preset: "discreet", pos: "left", font: "", size: 16, color: "#7f7f7f", qr: true, qrCm: 1.9, target: "doi", notes: false, bibQr: false };
  const PRESETS = {
    discreet: { pos: "left", size: 16, color: "#7f7f7f", qrCm: 1.9 },
    centered: { pos: "center", size: 16, color: "#595959", qrCm: 1.9 },
    bold: { pos: "right", size: 20, color: "#262626", qrCm: 2.5 },
    dark: { pos: "left", size: 16, color: "#d9d9d9", qrCm: 1.9 },
  };
  let P = Object.assign({}, defaultPrefs);
  const savePrefs = () => L().store.set(PREF_KEY, P);

  /* =====================================================================
     5. QR DEL RIFERIMENTO (motore del QR Lab, con il design scelto nella sezione QR)
     ===================================================================== */
  function refUrl(r) { return P.target === "oa" && r.oaUrl ? r.oaUrl : doiUrl(r); }
  function refQR(r, small) {
    const url = refUrl(r);
    if (!url) return null;
    const d = Object.assign({}, L().design(), { frame: "none" }); // la cornice con testo non serve accanto a una citazione
    if (small) { d.logo = null; }
    return L().buildSVG(url, d, d.logo ? "H" : d.ec);
  }

  /* =====================================================================
     6. POWERPOINT: dimensioni slide, inserimento, scansione, bibliografia
     ===================================================================== */
  const ppt = (v) => { try { return L().isOffice() && window.PowerPoint && Office.context.requirements.isSetSupported("PowerPointApi", v); } catch (e) { return false; } };
  const fullApi = () => ppt("1.5"); // caselle di testo posizionate, selezione di slide e forme
  const setData = (val, opts) => new Promise((res, rej) => Office.context.document.setSelectedDataAsync(val, opts, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? res() : rej(r.error))));

  async function slideSize(ctx) {
    if (ppt("1.10")) {
      try { const ps = ctx.presentation.pageSetup; ps.load("slideWidth,slideHeight"); await ctx.sync(); if (ps.slideWidth > 0) return { W: ps.slideWidth, H: ps.slideHeight }; } catch (e) { /* passo al metodo successivo */ }
    }
    try { // stima dai segnaposto del master: la forma più estesa indica la dimensione della slide
      const ms = ctx.presentation.slideMasters.getItemAt(0).shapes; ms.load("items/left,items/top,items/width,items/height"); await ctx.sync();
      let w = 0, h = 0; ms.items.forEach((s) => { w = Math.max(w, s.left + s.width); h = Math.max(h, s.top + s.height); });
      const known = [[960, 540], [720, 540], [960, 600], [792, 612], [1280, 720], [841.89, 595.28]];
      const best = known.filter(([kw, kh]) => kw >= w - 2 && kh >= h - 2).sort((a, b) => a[0] * a[1] - b[0] * b[1])[0];
      if (best) return { W: best[0], H: best[1] };
    } catch (e) { /* uso il 16:9 standard */ }
    return { W: 960, H: 540 };
  }
  const MARGIN = 18;
  function layout(W, H, text, withQr) {
    const q = withQr ? Math.max(28, Math.min(120, P.qrCm * CM_TO_PT)) : 0, gap = withQr ? 10 : 0;
    const boxH = Math.max(P.size * 1.5, 18);
    const textW = Math.min(W - 2 * MARGIN - q - gap, Math.max(60, text.length * P.size * 0.55 + 16));
    const out = { q, boxH, align: "Left" };
    const bottom = H - MARGIN;
    out.ty = bottom - boxH; out.qy = bottom - q;
    if (P.pos === "left") { out.qx = MARGIN; out.tx = MARGIN + q + gap; out.tw = Math.min(W * 0.62, W - out.tx - MARGIN); }
    else if (P.pos === "right") { out.qx = W - MARGIN - q; out.tw = Math.min(W * 0.62, W - 2 * MARGIN - q - gap); out.tx = W - MARGIN - q - gap - out.tw; out.align = "Right"; }
    else {
      if (!withQr) { out.tx = MARGIN; out.tw = W - 2 * MARGIN; out.align = "Center"; }
      else { const total = textW + gap + q; out.tx = (W - total) / 2; out.tw = textW; out.qx = out.tx + textW + gap; }
    }
    return out;
  }
  function styleText(shape, opts) {
    const tf = shape.textFrame;
    try { tf.wordWrap = true; tf.autoSizeSetting = "AutoSizeNone"; tf.verticalAlignment = opts.valign || "Bottom"; tf.leftMargin = 0; tf.rightMargin = 0; tf.topMargin = 0; tf.bottomMargin = 0; } catch (e) { /* proprietà non disponibile */ }
    const tr = tf.textRange;
    tr.font.size = opts.size;
    if (opts.color) tr.font.color = opts.color;
    if (opts.font) tr.font.name = opts.font;
    if (opts.bold != null) tr.font.bold = opts.bold;
    try { tr.paragraphFormat.horizontalAlignment = opts.align || "Left"; } catch (e) { /* ignora */ }
  }
  async function qrImagePayload(ref, sidePt) {
    const out = refQR(ref);
    if (!out) return null;
    if (ppt("1.4") && Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) return { data: out.svg, type: Office.CoercionType.XmlSvg };
    const png = await L().svgToPng(out.svg, out.W, out.H, Math.max(300, Math.min(900, Math.round((sidePt / 72) * 300))));
    return { data: png.split(",")[1], type: Office.CoercionType.Image };
  }
  // Inserisce un'immagine QR nella slide selezionata in una posizione precisa e restituisce l'id della nuova forma
  async function placeQr(ref, x, y, side, slideId, kind = "qr") {
    const pl = kind === "bibqr" ? await qrImagePayloadSmall(ref, side) : await qrImagePayload(ref, side);
    if (!pl) return null;
    let before = null;
    if (fullApi()) {
      await PowerPoint.run(async (ctx) => {
        if (slideId) { ctx.presentation.setSelectedSlides([slideId]); await ctx.sync(); }
        const sl = ctx.presentation.getSelectedSlides(); sl.load("items"); await ctx.sync();
        if (!sl.items.length) return;
        try { sl.items[0].setSelectedShapes([]); } catch (e) { /* ignora */ }
        const sh = sl.items[0].shapes; sh.load("items/id"); await ctx.sync();
        before = new Set(sh.items.map((s) => s.id));
      });
    }
    await setData(pl.data, { coercionType: pl.type, imageLeft: x, imageTop: y, imageWidth: side, imageHeight: side });
    if (!before) return null;
    let id = null;
    await PowerPoint.run(async (ctx) => {
      const sl = ctx.presentation.getSelectedSlides(); sl.load("items"); await ctx.sync();
      if (!sl.items.length) return;
      const sh = sl.items[0].shapes; sh.load("items/id,items/name"); await ctx.sync();
      const fresh = sh.items.filter((s) => !before.has(s.id));
      const pic = fresh[fresh.length - 1];
      if (pic) { pic.name = NAME_PREFIX + ref.id + ":" + kind; id = pic.id; try { pic.tags.add("USREF", ref.id); } catch (e) { /* ignora */ } await ctx.sync(); }
    });
    return id;
  }

  async function insertOnSlide(ref, text) {
    const withQr = P.qr && !!refUrl(ref);
    if (!fullApi()) {
      // PowerPoint meno recente: testo e QR vengono inseriti, ma la posizione va sistemata a mano
      await setData(text, { coercionType: Office.CoercionType.Text });
      if (withQr) { const pl = await qrImagePayload(ref, P.qrCm * CM_TO_PT); if (pl) await setData(pl.data, { coercionType: pl.type, imageWidth: P.qrCm * CM_TO_PT, imageHeight: P.qrCm * CM_TO_PT }); }
      return "basic";
    }
    let slideId = null, textId = null, g = null;
    await PowerPoint.run(async (ctx) => {
      const sl = ctx.presentation.getSelectedSlides(); sl.load("items/id"); await ctx.sync();
      if (!sl.items.length) throw new RefError("noslide");
      const slide = sl.items[0]; slideId = slide.id;
      const { W, H } = await slideSize(ctx);
      g = layout(W, H, text, withQr);
      const tb = slide.shapes.addTextBox(text, { left: g.tx, top: g.ty, width: g.tw, height: g.boxH });
      tb.name = NAME_PREFIX + ref.id + ":txt:" + DOC.shortFmt;
      styleText(tb, { size: P.size, color: P.color, font: P.font, align: g.align });
      try { tb.tags.add("USREF", ref.id); tb.tags.add("USREFFMT", DOC.shortFmt); } catch (e) { /* ignora */ }
      tb.load("id"); await ctx.sync(); textId = tb.id;
    });
    let qrId = null;
    if (withQr) qrId = await placeQr(ref, g.qx, g.qy, g.q, slideId);
    if (qrId && textId && ppt("1.8")) { // testo e QR raggruppati: si spostano insieme
      try {
        await PowerPoint.run(async (ctx) => {
          ctx.presentation.setSelectedSlides([slideId]);
          const sl = ctx.presentation.getSelectedSlides(); sl.load("items"); await ctx.sync();
          const grp = sl.items[0].shapes.addGroup([textId, qrId]);
          grp.name = NAME_PREFIX + ref.id + ":grp"; await ctx.sync();
        });
      } catch (e) { /* il gruppo è facoltativo */ }
    }
    return "full";
  }

  // Legge le citazioni presenti nelle slide, nell'ordine di comparsa
  async function scanSlides(ctx) {
    const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
    const per = slides.items.map((s) => { const c = s.shapes; c.load("items/id,items/name,items/type"); return c; });
    await ctx.sync();
    const groups = [];
    per.forEach((c) => c.items.forEach((sh) => { if (sh.name && sh.name.startsWith(NAME_PREFIX) && /:grp$/.test(sh.name)) { try { const gs = sh.group.shapes; gs.load("items/id,items/name"); groups.push({ id: sh.id, gs }); } catch (e) { /* senza API dei gruppi */ } } }));
    if (groups.length) { try { await ctx.sync(); } catch (e) { groups.length = 0; } }
    const found = [];
    slides.items.forEach((slide, si) => {
      per[si].items.forEach((sh) => {
        const list = /:grp$/.test(sh.name || "") ? ((groups.find((g) => g.id === sh.id) || {}).gs || { items: [] }).items : [sh];
        list.forEach((x) => {
          const m = (x.name || "").match(/^UltraspeakerRef:([^:]+):(txt|qr|bib|bibqr|bibtitle)(?::(\w+))?$/);
          if (m) found.push({ slide, slideIdx: si, shape: x, id: m[1], kind: m[2], fmt: m[3] || "" });
        });
      });
    });
    return found;
  }
  // Riordina in base alla comparsa nelle slide e aggiorna il testo delle citazioni
  async function renumber() {
    const res = { found: 0 };
    await PowerPoint.run(async (ctx) => {
      const found = await scanSlides(ctx);
      const order = [];
      found.filter((f) => f.kind === "txt" || f.kind === "qr").forEach((f) => { if (!order.includes(f.id)) order.push(f.id); });
      const present = order.map((id) => DOC.refs.find((r) => r.id === id)).filter(Boolean);
      const rest = DOC.refs.filter((r) => !present.includes(r));
      DOC.refs = present.concat(rest);
      DOC.refs.forEach((r) => { r.onSlides = present.includes(r); });
      found.filter((f) => f.kind === "txt").forEach((f) => {
        const r = DOC.refs.find((x) => x.id === f.id);
        if (r) f.shape.textFrame.textRange.text = formatShort(r, f.fmt || "authorJournal", numOf(r));
      });
      await ctx.sync();
      res.found = present.length;
    });
    saveDoc();
    return res;
  }
  async function deleteFromSlides(id) {
    await PowerPoint.run(async (ctx) => {
      // prima i gruppi testo+QR di questo riferimento, poi le forme singole rimaste
      const slides = ctx.presentation.slides; slides.load("items"); await ctx.sync();
      const per = slides.items.map((s) => { const c = s.shapes; c.load("items/id,items/name"); return c; }); await ctx.sync();
      per.forEach((c) => c.items.forEach((sh) => { if (sh.name === NAME_PREFIX + id + ":grp") sh.delete(); }));
      await ctx.sync();
      const found = await scanSlides(ctx);
      found.filter((f) => f.id === id).forEach((f) => f.shape.delete());
      await ctx.sync();
    });
  }

  // Bibliografia: una o più slide finali; le precedenti generate dall'add-in vengono sostituite
  function bibEntries() {
    const style = DOC.style, list = DOC.refs.slice();
    if (!NUMBERED(style)) list.sort((a, b) => (firstAuthors(a) || a.title || "").localeCompare(firstAuthors(b) || b.title || "", undefined, { sensitivity: "base" }) || String(a.year).localeCompare(String(b.year)));
    return list.map((r) => ({ ref: r, text: (NUMBERED(style) ? numOf(r) + ". " : "") + formatFull(r, style) }));
  }
  function paginate(entries, W, H, size, withQr) {
    const top = 96, usable = H - top - MARGIN - 6, width = W - 2 * MARGIN - (withQr ? 50 : 0);
    const cpl = Math.max(30, Math.floor(width / (size * 0.56)));
    const pages = []; let cur = [], used = 0;
    entries.forEach((e) => {
      const lines = Math.max(1, Math.ceil(e.text.length / cpl));
      const h = Math.max(withQr ? 40 : 0, lines * size * 1.25) + size * 0.6;
      if (cur.length && used + h > usable) { pages.push(cur); cur = []; used = 0; }
      cur.push(Object.assign({ h, y: top + used }, e)); used += h;
    });
    if (cur.length) pages.push(cur);
    return pages;
  }
  async function generateBibliography(withQr) {
    if (NUMBERED(DOC.style)) { try { await renumber(); } catch (e) { /* numerazione non aggiornata */ } }
    const entries = bibEntries();
    const qrJobs = [];
    let pagesCount = 0;
    await PowerPoint.run(async (ctx) => {
      // elimina le slide di bibliografia generate in precedenza
      const found = await scanSlides(ctx);
      const old = []; found.forEach((f) => { if (f.kind === "bibtitle" && !old.includes(f.slide)) old.push(f.slide); });
      old.forEach((s) => s.delete());
      await ctx.sync();
      const { W, H } = await slideSize(ctx);
      // layout più vuoto disponibile (di solito "Vuota"), così la bibliografia non eredita segnaposto
      let layoutId = null;
      try {
        const lays = ctx.presentation.slideMasters.getItemAt(0).layouts; lays.load("items/id"); await ctx.sync();
        const counts = lays.items.map((l) => l.shapes.getCount()); await ctx.sync();
        let bi = 0; counts.forEach((c, i) => { if (c.value < counts[bi].value) bi = i; });
        layoutId = lays.items[bi] && lays.items[bi].id;
      } catch (e) { layoutId = null; }
      const size = entries.length > 8 ? 12 : 14;
      const pages = paginate(entries, W, H, size, withQr);
      pagesCount = pages.length;
      for (let p = 0; p < pages.length; p++) {
        ctx.presentation.slides.add(layoutId ? { layoutId } : undefined);
        await ctx.sync();
        const all = ctx.presentation.slides; all.load("items/id"); await ctx.sync();
        const slide = all.items[all.items.length - 1];
        const title = t("ref.bibTitle") + (pages.length > 1 ? ` (${p + 1}/${pages.length})` : "");
        const tt = slide.shapes.addTextBox(title, { left: MARGIN * 2, top: 30, width: W - MARGIN * 4, height: 46 });
        tt.name = NAME_PREFIX + "bib:bibtitle";
        styleText(tt, { size: 28, font: P.font, bold: true, valign: "Middle" });
        pages[p].forEach((e) => {
          const tb = slide.shapes.addTextBox(e.text, { left: MARGIN * 2, top: e.y, width: W - MARGIN * 4 - (withQr ? 50 : 0), height: e.h });
          tb.name = NAME_PREFIX + e.ref.id + ":bib";
          styleText(tb, { size, font: P.font, valign: "Top" });
          if (withQr && refUrl(e.ref)) qrJobs.push({ ref: e.ref, x: W - MARGIN * 2 - 40, y: e.y, slideId: slide.id });
        });
        await ctx.sync();
      }
    });
    let qrFail = 0;
    for (const j of qrJobs) { // QR piccoli accanto alle voci (senza logo, per restare leggibili)
      try { await placeQr(j.ref, j.x, j.y, 40, j.slideId, "bibqr"); } catch (e) { qrFail++; }
    }
    return { pages: pagesCount, entries: entries.length, qrFail };
  }
  async function qrImagePayloadSmall(ref, sidePt) {
    const out = refQR(ref, true);
    if (!out) return null;
    if (Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) return { data: out.svg, type: Office.CoercionType.XmlSvg };
    const png = await L().svgToPng(out.svg, out.W, out.H, 300);
    return { data: png.split(",")[1], type: Office.CoercionType.Image };
  }

  /* =====================================================================
     7. APPUNTI
     ===================================================================== */
  function copyText(text) {
    let ok = false;
    try {
      const ta = document.createElement("textarea");
      ta.value = text; ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select(); ok = document.execCommand("copy"); ta.remove();
    } catch (e) { ok = false; }
    if (!ok && navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(text).catch(() => {}); ok = true; }
    return ok;
  }

  /* =====================================================================
     8. INTERFACCIA
     ===================================================================== */
  let cur = null;        // riferimento in anteprima
  let results = [];      // risultati di una ricerca libera
  let editing = null;    // id del riferimento in modifica
  let confirmDel = null; // id del riferimento in attesa di conferma eliminazione
  let busy = false;
  const errMsg = (e) => t("ref.err." + ((e && e.code) || "generic")) || t("ref.err.generic");
  function setStatus(msg, kind) {
    const el = $("#refStatus");
    el.textContent = msg || ""; el.className = "status" + (kind === "err" ? " ref-err" : "");
  }
  const authorsLine = (r) => { const A = r.authors || []; const n = A.slice(0, 6).map(nlmName).join(", "); return A.length > 6 ? n + ", et al." : n; };
  const srcLine = (r) => [r.journalAbbr || r.journal, [r.year, r.volume ? r.volume + (r.issue ? "(" + r.issue + ")" : "") : "", hyph(r.pages)].filter(Boolean).join(" · ")].filter(Boolean).join(" — ");

  function renderResults() {
    const box = $("#refResults");
    if (!results.length) { box.innerHTML = ""; box.hidden = true; return; }
    box.hidden = false;
    box.innerHTML = `<div class="lbl">${esc(t("ref.results"))}</div>` + results.map((r, i) => `<button type="button" class="ref-res" data-i="${i}"><b>${esc(r.title)}</b><span>${esc(authorsLine(r))}</span><span>${esc(srcLine(r))}</span></button>`).join("");
    $$(".ref-res", box).forEach((b) => b.addEventListener("click", async () => {
      const r = results[+b.dataset.i];
      const p = r.pmid ? { kind: "pmid", id: r.pmid } : r.doi ? { kind: "doi", id: r.doi } : null;
      if (!p) { cur = Object.assign({ id: uid(), added: Date.now() }, r); results = []; renderResults(); renderPreview(); return; }
      results = []; renderResults();
      await runLookup(p);
    }));
  }
  async function runLookup(p) {
    if (busy) return;
    busy = true; setStatus(t("ref.loading")); $("#refSearchBtn").disabled = true; $("#refInsertBtn").disabled = true;
    try {
      cur = await lookup(p);
      cur.abstract = null;
      setStatus("");
      renderPreview();
      if (cur.pmid) loadAbstract(cur);
    } catch (e) { cur = null; renderPreview(); setStatus(errMsg(e), "err"); }
    finally { busy = false; $("#refSearchBtn").disabled = false; $("#refInsertBtn").disabled = !cur; }
  }
  async function loadAbstract(ref) {
    try { const a = await fetchAbstract(ref.pmid); if (cur === ref) { ref.abstract = a || ""; renderAbstract(); } }
    catch (e) { if (cur === ref) { ref.abstract = ""; renderAbstract(); } }
  }
  async function doSearch() {
    const raw = $("#refQuery").value;
    const p = parseInput(raw);
    results = []; renderResults();
    if (!p) { setStatus(t("ref.err.empty"), "err"); return; }
    if (p.kind !== "text") return runLookup(p);
    if (busy) return;
    busy = true; setStatus(t("ref.searching")); $("#refSearchBtn").disabled = true; $("#refInsertBtn").disabled = true;
    try {
      results = await searchText(p.q);
      setStatus(results.length ? "" : t("ref.noResults"), results.length ? "" : "err");
      renderResults();
    } catch (e) { setStatus(errMsg(e), "err"); }
    finally { busy = false; $("#refSearchBtn").disabled = false; $("#refInsertBtn").disabled = !cur; }
  }
  async function doImport() {
    const items = parseList($("#refBulk").value);
    const prog = $("#refBulkStatus");
    if (!items.length) { prog.textContent = t("ref.bulk.none"); prog.className = "status ref-err"; return; }
    if (busy) return;
    busy = true; $("#refBulkBtn").disabled = true; prog.className = "status";
    let added = 0, dups = 0; const failed = [];
    for (let i = 0; i < items.length; i++) {
      prog.textContent = t("ref.bulk.progress", { n: i + 1, tot: items.length });
      try {
        const r = await lookup(items[i]);
        const res = addToDoc(r);
        res.dup ? dups++ : added++;
      } catch (e) {
        failed.push(items[i].id + " (" + errMsg(e) + ")");
        if (e.code === "offline") break;
      }
    }
    busy = false; $("#refBulkBtn").disabled = false;
    prog.textContent = t("ref.bulk.done", { added, dups, failed: failed.length }) + (failed.length ? " " + failed.join("; ") : "");
    if (failed.length) prog.className = "status ref-err";
    renderList();
  }

  function badge(cls, text, href) {
    cls = "b-" + cls;
    return href ? `<a class="ref-badge ${cls}" href="${esc(href)}" target="_blank" rel="noopener">${esc(text)}</a>` : `<span class="ref-badge ${cls}">${esc(text)}</span>`;
  }
  function statusBadges(r) {
    const st = r.status || {};
    const out = [];
    const notice = (ty) => { const n = (st.notices || []).find((x) => ty.includes(x.type)); return n ? "https://doi.org/" + n.doi : ""; };
    if (st.retracted) out.push(badge("bad", t("ref.flag.retracted"), notice(RETRACT)));
    if (st.concern) out.push(badge("warn", t("ref.flag.concern"), notice(CONCERN)));
    if (st.corrected) out.push(badge("warn", t("ref.flag.corrected"), notice(CORRECT)));
    if (r.oaUrl) out.push(badge("ok", t("ref.flag.oa"), r.oaUrl));
    return out.join("");
  }
  function renderPreview() {
    const box = $("#refPreview");
    $("#refInsertCard").hidden = !cur; $("#refCiteCard").hidden = !cur;
    if (!cur) { box.innerHTML = `<p class="ref-empty">${esc(t("ref.preview.empty"))}</p>`; return; }
    const r = cur, ex = findSame(r);
    const warn = [];
    if (r.status && r.status.retracted) warn.push(`<div class="warn ref-warn-bad">${esc(t("ref.warn.retracted"))}</div>`);
    else if (r.status && r.status.concern) warn.push(`<div class="warn">${esc(t("ref.warn.concern"))}</div>`);
    else if (r.status && r.status.corrected) warn.push(`<div class="warn">${esc(t("ref.warn.corrected"))}</div>`);
    if (ex) warn.push(`<div class="warn">${esc(t("ref.warn.dup", { n: numOf(ex) }))}</div>`);
    const links = [];
    if (r.doi) links.push(`<a href="https://doi.org/${esc(r.doi)}" target="_blank" rel="noopener">DOI ${esc(r.doi)}</a>`);
    if (r.pmid) links.push(`<a href="https://pubmed.ncbi.nlm.nih.gov/${esc(r.pmid)}/" target="_blank" rel="noopener">PMID ${esc(r.pmid)}</a>`);
    box.innerHTML = `
      <div class="ref-meta">
        <div class="ref-authors">${esc(authorsLine(r)) || "—"}</div>
        <div class="ref-title">${esc(r.title) || "—"}</div>
        <dl class="ref-dl">
          <dt>${esc(t("ref.f.journal"))}</dt><dd>${esc(r.journal || r.journalAbbr || "—")}${r.journalAbbr && r.journal && r.journalAbbr !== r.journal ? ` <span class="ref-dim">(${esc(r.journalAbbr)})</span>` : ""}</dd>
          <dt>${esc(t("ref.f.year"))}</dt><dd>${esc(r.year || "—")}</dd>
          <dt>${esc(t("ref.f.volume"))}</dt><dd>${esc(r.volume || "—")}${r.issue ? ` (${esc(r.issue)})` : ""}</dd>
          <dt>${esc(t("ref.f.pages"))}</dt><dd>${esc(hyph(r.pages) || "—")}</dd>
        </dl>
        <div class="ref-links">${links.join("")}</div>
        <div class="ref-badges">${statusBadges(r)}</div>
      </div>${warn.join("")}
      <details class="ref-abs" id="refAbs" ${r.pmid ? "" : "hidden"}><summary>${esc(t("ref.abstract"))}</summary><div id="refAbsBody"></div></details>`;
    $$("#refPreview a").forEach(linkify);
    renderAbstract();
    renderCite();
  }
  function renderAbstract() {
    const b = $("#refAbsBody");
    if (!b || !cur) return;
    if (cur.abstract == null) b.innerHTML = `<p class="ref-dim">${esc(t("ref.loading"))}</p>`;
    else if (!cur.abstract) b.innerHTML = `<p class="ref-dim">${esc(t("ref.abstract.none"))}</p>`;
    else {
      b.innerHTML = `<p class="ref-abs-text">${esc(cur.abstract).replace(/\n\n/g, "</p><p class=\"ref-abs-text\">")}</p><button type="button" class="inline-btn" id="refCopyAbs">${esc(t("ref.copyAbs"))}</button>`;
      $("#refCopyAbs").addEventListener("click", () => L().toast(copyText(cur.abstract) ? t("ref.toast.absCopied") : t("ref.toast.copyFail")));
    }
    const chk = $("#refNotesWrap");
    if (chk) chk.hidden = !cur.pmid;
  }
  function currentNum(r) { const ex = r && (DOC.refs.includes(r) ? r : findSame(r)); return ex ? numOf(ex) : DOC.refs.length + 1; }
  function renderCite() {
    $$("#refStyle button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === DOC.style));
    $$("#refShort button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === DOC.shortFmt));
    if (!cur) return;
    $("#refShortOut").textContent = formatShort(cur, DOC.shortFmt, currentNum(cur));
    $("#refFullOut").textContent = (NUMBERED(DOC.style) ? currentNum(cur) + ". " : "") + formatFull(cur, DOC.style);
    renderInsertUI();
  }
  function renderInsertUI() {
    $$("#refPresets .tile").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === P.preset));
    $$("#refPos button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === P.pos));
    $$("#refTarget button").forEach((b) => b.setAttribute("aria-pressed", b.dataset.v === P.target));
    const oaBtn = $('#refTarget button[data-v="oa"]');
    if (oaBtn) oaBtn.disabled = !(cur && cur.oaUrl);
    const setV = (el, v) => { if (el && el !== document.activeElement && el.value !== String(v)) el.value = v; };
    setV($("#refFont"), P.font); setV($("#refSize"), P.size); setV($("#refColor"), P.color); setV($("#refColorHex"), P.color);
    $("#refQrOn").checked = P.qr; $("#refQrOpts").hidden = !P.qr;
    setV($("#refQrCm"), P.qrCm); $("#refQrCmVal").textContent = (+P.qrCm).toFixed(1).replace(".", t("ref.decimal")) + " cm";
    $("#refNotes").checked = P.notes;
    $("#refTargetNote").textContent = cur && P.target === "oa" && !cur.oaUrl ? t("ref.noOa") : "";
    // anteprima dal vivo: come apparirà in basso nella slide
    const pv = $("#refSlidePv");
    if (pv && cur) {
      const text = formatShort(cur, DOC.shortFmt, currentNum(cur));
      const qr = P.qr ? refQR(cur) : null;
      const img = qr ? `<img alt="" src="data:image/svg+xml;charset=utf-8,${encodeURIComponent(L().withFont(qr.svg))}" style="height:${Math.round(P.qrCm * 9)}px">` : "";
      const txt = `<span style="font-size:${Math.max(7, P.size * 0.5)}px;color:${esc(P.color)};${P.font ? "font-family:'" + esc(P.font) + "'," + "sans-serif;" : ""}">${esc(text)}</span>`;
      pv.className = "ref-slide pos-" + P.pos + (P.preset === "dark" || hexLum(P.color) > 0.6 ? " dark" : "");
      pv.innerHTML = `<div class="ref-slide-in">${P.pos === "right" ? txt + img : img + txt}</div>`;
      if (P.pos === "center") pv.querySelector(".ref-slide-in").innerHTML = txt + img;
    }
    const ins = $("#refInsertBtn");
    ins.disabled = !cur;
    $("#refInsertLabel").textContent = L().isOffice() ? t("ref.insert") : t("ref.copyShort");
    $("#refDlQr").hidden = L().isOffice();
    $("#refDlQr").disabled = !cur || !refUrl(cur || {});
  }
  function hexLum(h) {
    const m = /^#?([0-9a-f]{6})$/i.exec(h || ""); if (!m) return 0;
    const n = parseInt(m[1], 16); return (0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255;
  }

  async function doInsert() {
    if (!cur || busy) return;
    const btn = $("#refInsertBtn");
    // l'abstract va copiato subito, finché il clic è ancora "fresco" (requisito degli appunti)
    let absCopied = false;
    if (P.notes && cur.abstract) absCopied = copyText(cur.abstract);
    const { ref, dup } = addToDoc(cur);
    if (!L().isOffice()) { // versione web: copia della citazione
      const ok = copyText(formatShort(ref, DOC.shortFmt, numOf(ref)));
      L().toast(ok ? t("ref.toast.copied") : t("ref.toast.copyFail"));
      renderList(); renderCite();
      return;
    }
    busy = true; btn.disabled = true; setStatus(t("ref.inserting"));
    try {
      const mode = await insertOnSlide(ref, formatShort(ref, DOC.shortFmt, numOf(ref)));
      ref.onSlides = true; saveDoc();
      let msg = t(mode === "basic" ? "ref.toast.insertedBasic" : "ref.toast.inserted");
      if (dup) msg += " " + t("ref.toast.reused", { n: numOf(ref) });
      if (P.notes && cur.abstract) msg += " " + t(absCopied ? "ref.toast.absNotes" : "ref.toast.copyFail");
      setStatus(""); L().toast(msg);
    } catch (e) {
      setStatus(e && e.code === "noslide" ? t("ref.err.noslide") : t("ref.err.insert"), "err");
    } finally { busy = false; btn.disabled = !cur; renderList(); renderCite(); }
  }
  async function downloadQr() {
    if (!cur) return;
    const out = refQR(cur);
    if (!out) return;
    const png = await L().svgToPng(out.svg, out.W, out.H, 1200);
    const a = document.createElement("a"); a.href = png; a.download = "qr-reference-" + (cur.doi || cur.pmid || "ref").replace(/[^\w.-]+/g, "_") + ".png";
    document.body.appendChild(a); a.click(); a.remove();
  }

  /* --- elenco dei riferimenti della presentazione --- */
  function renderList() {
    const box = $("#refList");
    const dups = dupIds();
    $("#refCount").textContent = DOC.refs.length ? "(" + DOC.refs.length + ")" : "";
    $("#refBibBtns").hidden = !DOC.refs.length;
    if (!DOC.refs.length) { box.innerHTML = `<p class="ref-empty">${esc(t("ref.list.empty"))}</p>`; return; }
    box.innerHTML = DOC.refs.map((r, i) => {
      if (editing === r.id) return editForm(r, i);
      const flags = [];
      if (r.status && r.status.retracted) flags.push(badge("bad", t("ref.flag.retracted")));
      else if (r.status && (r.status.corrected || r.status.concern)) flags.push(badge("warn", t(r.status.concern ? "ref.flag.concern" : "ref.flag.corrected")));
      if (dups.has(r.id)) flags.push(badge("warn", t("ref.flag.dup")));
      if (L().isOffice() && r.onSlides === false) flags.push(badge("dim", t("ref.flag.notOnSlides")));
      const del = confirmDel === r.id;
      return `<div class="ref-item" data-id="${esc(r.id)}">
        <span class="ref-num">${i + 1}</span>
        <div class="ref-item-body"><div class="ref-item-short">${esc(formatShort(r, "authorJournal"))}</div><div class="ref-item-title">${esc(r.title)}</div>${flags.length ? `<div class="ref-badges">${flags.join("")}</div>` : ""}</div>
        <div class="ref-item-acts">
          <button type="button" data-act="use" title="${esc(t("ref.act.use"))}" aria-label="${esc(t("ref.act.use"))}"><svg viewBox="0 0 24 24"><path d="M4 12h12M12 6l6 6-6 6"/></svg></button>
          <button type="button" data-act="edit" title="${esc(t("ref.act.edit"))}" aria-label="${esc(t("ref.act.edit"))}"><svg viewBox="0 0 24 24"><path d="M4 20h4L19 9l-4-4L4 16z"/></svg></button>
          <button type="button" data-act="del" class="${del ? "confirm" : ""}" title="${esc(t(del ? "ref.act.confirmDel" : "ref.act.del"))}" aria-label="${esc(t(del ? "ref.act.confirmDel" : "ref.act.del"))}">${del ? esc(t("ref.act.confirmShort")) : '<svg viewBox="0 0 24 24"><path d="M5 7h14M10 7V4h4v3M7 7l1 13h8l1-13"/></svg>'}</button>
        </div></div>`;
    }).join("");
    $$(".ref-item [data-act]", box).forEach((b) => b.addEventListener("click", () => onItemAct(b.closest(".ref-item").dataset.id, b.dataset.act)));
    const f = $(".ref-edit", box);
    if (f) bindEditForm(f);
  }
  const FIELDS = [["authorsText", "ref.f.authors"], ["title", "ref.f.title"], ["journal", "ref.f.journal"], ["journalAbbr", "ref.f.journalAbbr"], ["year", "ref.f.year"], ["volume", "ref.f.volume"], ["issue", "ref.f.issue"], ["pages", "ref.f.pages"], ["doi", "DOI"], ["pmid", "PMID"]];
  const authorsToText = (r) => (r.authors || []).map((a) => (a.org ? a.family : a.family + (a.given ? ", " + a.given : ""))).join("; ");
  const textToAuthors = (s) => String(s || "").split(";").map((x) => x.trim()).filter(Boolean).map((x) => { const [f, g] = x.split(",").map((y) => y.trim()); return g ? { family: f, given: g, init: /^[A-Z]{1,5}$/.test(g) } : { family: f, given: "", org: x.split(/\s+/).length > 2 }; });
  function editForm(r, i) {
    const v = Object.assign({}, r, { authorsText: authorsToText(r) });
    return `<form class="ref-edit" data-id="${esc(r.id)}"><div class="ref-edit-h">${esc(t("ref.act.edit"))} · ${i + 1}</div>
      ${FIELDS.map(([k, lab]) => `<div class="field"><label>${esc(lab.startsWith("ref.") ? t(lab) : lab)}</label>${k === "title" || k === "authorsText" ? `<textarea name="${k}" rows="2">${esc(v[k] || "")}</textarea>` : `<input type="text" name="${k}" value="${esc(v[k] || "")}">`}${k === "authorsText" ? `<div class="hint">${esc(t("ref.f.authorsHint"))}</div>` : ""}</div>`).join("")}
      <div class="btns"><button class="btn" type="button" data-cancel>${esc(t("ref.cancel"))}</button><button class="btn" type="submit">${esc(t("ref.save"))}</button></div></form>`;
  }
  function bindEditForm(f) {
    f.addEventListener("submit", (e) => {
      e.preventDefault();
      const r = DOC.refs.find((x) => x.id === f.dataset.id);
      if (r) {
        const fd = new FormData(f);
        FIELDS.forEach(([k]) => { if (k !== "authorsText") r[k] = String(fd.get(k) || "").trim(); });
        r.doi = r.doi.toLowerCase().replace(/^https?:\/\/(dx\.)?doi\.org\//, "");
        r.authors = textToAuthors(fd.get("authorsText"));
        saveDoc();
        if (cur && cur.id === r.id) { Object.assign(cur, r); renderPreview(); }
        L().toast(t(L().isOffice() ? "ref.toast.savedSync" : "ref.toast.saved"));
      }
      editing = null; renderList(); renderCite();
    });
    $("[data-cancel]", f).addEventListener("click", () => { editing = null; renderList(); });
    const first = $("textarea,input", f); if (first) first.focus();
  }
  async function onItemAct(id, act) {
    const r = DOC.refs.find((x) => x.id === id);
    if (!r) return;
    if (act === "use") { cur = JSON.parse(JSON.stringify(r)); cur.abstract = null; renderPreview(); if (cur.pmid) loadAbstract(cur); $("#refPreviewCard").scrollIntoView({ behavior: "smooth", block: "start" }); return; }
    if (act === "edit") { editing = id; confirmDel = null; renderList(); return; }
    if (act === "del") {
      if (confirmDel !== id) { confirmDel = id; renderList(); setTimeout(() => { if (confirmDel === id) { confirmDel = null; renderList(); } }, 4000); return; }
      confirmDel = null;
      DOC.refs = DOC.refs.filter((x) => x !== r); saveDoc();
      if (L().isOffice() && fullApi()) { try { await deleteFromSlides(id); if (NUMBERED(DOC.style) || DOC.shortFmt === "numeric") await renumber(); } catch (e) { /* le forme restano nella slide */ } }
      renderList(); renderCite(); L().toast(t("ref.toast.deleted"));
    }
  }
  async function doRenumber() {
    if (!L().isOffice()) return;
    if (!fullApi()) { L().toast(t("ref.err.oldPpt")); return; }
    const b = $("#refRenumBtn"); b.disabled = true;
    try { const r = await renumber(); L().toast(t("ref.toast.renumbered", { n: r.found })); }
    catch (e) { L().toast(t("ref.err.scan")); }
    finally { b.disabled = false; renderList(); renderCite(); }
  }
  async function doBibliography() {
    const withQr = $("#refBibQr").checked;
    P.bibQr = withQr; savePrefs();
    if (!L().isOffice()) {
      const text = t("ref.bibTitle") + "\n\n" + bibEntries().map((e) => e.text).join("\n");
      L().toast(copyText(text) ? t("ref.toast.bibCopied") : t("ref.toast.copyFail"));
      return;
    }
    if (!fullApi()) { L().toast(t("ref.err.oldPpt")); return; }
    const b = $("#refBibBtn"); b.disabled = true; setListStatus(t("ref.bibWorking"));
    try {
      const r = await generateBibliography(withQr);
      setListStatus("");
      L().toast(t("ref.toast.bib", { n: r.entries, p: r.pages }) + (r.qrFail ? " " + t("ref.toast.bibQrFail") : ""));
    } catch (e) { setListStatus(t("ref.err.bib"), "err"); }
    finally { b.disabled = false; renderList(); }
  }
  function setListStatus(msg, kind) { const el = $("#refListStatus"); el.textContent = msg || ""; el.className = "status" + (kind === "err" ? " ref-err" : ""); }

  function linkify(a) {
    a.addEventListener("click", (e) => {
      if (!L().isOffice()) return;
      e.preventDefault(); L().openUrl(a.href);
    });
  }

  /* --- montaggio dei controlli --- */
  function presetTiles() {
    const box = $("#refPresets");
    box.innerHTML = Object.keys(PRESETS).map((k) => {
      const p = PRESETS[k], dark = k === "dark";
      const bar = `<span class="pt-txt" style="background:${p.color};width:${p.size * 1.3}px"></span>`;
      const q = `<span class="pt-qr" style="width:${p.qrCm * 6}px;height:${p.qrCm * 6}px"></span>`;
      return `<button type="button" class="tile" data-v="${k}" aria-pressed="false"><span class="thumb pt ${dark ? "pt-dark" : ""} pt-${p.pos}">${p.pos === "right" ? bar + q : q + bar}</span><span>${esc(t("ref.preset." + k))}</span></button>`;
    }).join("");
    $$(".tile", box).forEach((b) => b.addEventListener("click", () => { Object.assign(P, PRESETS[b.dataset.v], { preset: b.dataset.v }); savePrefs(); renderInsertUI(); }));
  }
  function custom() { P.preset = "custom"; savePrefs(); renderInsertUI(); }
  function bind() {
    $("#refSearchBtn").addEventListener("click", doSearch);
    $("#refQuery").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); doSearch(); } });
    $("#refBulkBtn").addEventListener("click", doImport);
    $$("#refStyle button").forEach((b) => b.addEventListener("click", () => { DOC.style = b.dataset.v; saveDoc(); renderCite(); }));
    $$("#refShort button").forEach((b) => b.addEventListener("click", () => { DOC.shortFmt = b.dataset.v; saveDoc(); renderCite(); }));
    $$("#refPos button").forEach((b) => b.addEventListener("click", () => { P.pos = b.dataset.v; custom(); }));
    $$("#refTarget button").forEach((b) => b.addEventListener("click", () => { P.target = b.dataset.v; savePrefs(); renderInsertUI(); }));
    $("#refFont").addEventListener("change", (e) => { P.font = e.target.value; custom(); });
    $("#refSize").addEventListener("input", (e) => { const v = +e.target.value; if (v >= 8 && v <= 60) { P.size = v; custom(); } });
    $("#refColor").addEventListener("input", (e) => { P.color = e.target.value; custom(); });
    $("#refColorHex").addEventListener("change", (e) => { let v = e.target.value.trim(); if (!v.startsWith("#")) v = "#" + v; if (/^#[0-9a-f]{6}$/i.test(v)) { P.color = v.toLowerCase(); custom(); } else renderInsertUI(); });
    $("#refQrOn").addEventListener("change", (e) => { P.qr = e.target.checked; savePrefs(); renderInsertUI(); });
    $("#refQrCm").addEventListener("input", (e) => { P.qrCm = +e.target.value; custom(); });
    $("#refNotes").addEventListener("change", (e) => { P.notes = e.target.checked; savePrefs(); });
    $("#refInsertBtn").addEventListener("click", doInsert);
    $("#refDlQr").addEventListener("click", downloadQr);
    $("#refQrDesign").addEventListener("click", () => L().showQrDesign());
    $("#refRenumBtn").addEventListener("click", doRenumber);
    $("#refBibBtn").addEventListener("click", doBibliography);
    $("#refBibQr").checked = !!P.bibQr;
  }
  function applyLang() {
    if (!inited) return;
    presetTiles();
    $("#refBibBtnLabel").textContent = L().isOffice() ? t("ref.bibGen") : t("ref.bibCopy");
    $("#refRenumBtn").hidden = !L().isOffice();
    renderResults(); renderPreview(); renderList(); renderCite();
  }

  let inited = false;
  function init() {
    if (inited) return;
    inited = true;
    P = Object.assign({}, defaultPrefs, L().store.get(PREF_KEY, {}) || {});
    loadDoc();
    bind();
    applyLang();
  }
  // Il pannello si è accorto di essere in PowerPoint solo dopo l'avvio: ricarico dal documento
  function officeReady() { if (!inited) return; loadDoc(); applyLang(); }

  window.QRRef = { init, applyLang, officeReady, refresh: () => { if (inited) renderInsertUI(); },
    _test: { parseInput, parseList, fromCrossref, fromPubmed, formatFull, formatShort, layout, paginate, statusFrom, get DOC() { return DOC; }, get P() { return P; } } };
})();
