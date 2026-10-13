(() => {
  "use strict";

  const state = {
    paymentFile: null,
    baseFile: null,
    payments: [],
    quarters: [],
    paymentOcrText: "",
    baseText: "",
    sourceRate: null,
    sourceRateLabel: "",
    results: null,
    ocrWorker: null
  };

  const $ = (id) => document.getElementById(id);
  const amountFmt = new Intl.NumberFormat("cs-CZ", { maximumFractionDigits: 0 });
  const currencyFmt = (value) => amountFmt.format(Math.round(Number(value) || 0)) + " Kč";
  const dateFmt = (iso) => {
    if (!iso) return "—";
    const d = parseDate(iso);
    return d ? new Intl.DateTimeFormat("cs-CZ").format(d) : String(iso);
  };
  const isoToday = () => {
    const d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  };
  $("asOfDate").value = isoToday();

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toLocaleString("cs-CZ", { maximumFractionDigits: 1 }) + " kB";
    return (bytes / (1024 * 1024)).toLocaleString("cs-CZ", { maximumFractionDigits: 2 }) + " MB";
  }

  function setStatus(title, message, busy) {
    $("statusPanel").hidden = false;
    $("statusTitle").textContent = title;
    $("statusText").textContent = message || "";
    $("spinner").classList.toggle("done", !busy);
    $("spinner").hidden = !busy && !message;
  }

  function clearStatus() { $("statusPanel").hidden = true; }

  function selectFile(kind, file) {
    if (!file) return;
    if (kind === "payments") state.paymentFile = file;
    else state.baseFile = file;
    const meta = $(kind === "payments" ? "paymentsMeta" : "baseMeta");
    meta.textContent = file.name + " · " + formatSize(file.size) + " · pouze lokálně";
    meta.classList.add("selected");
    const drop = $(kind === "payments" ? "paymentsDrop" : "baseDrop");
    drop.classList.add("has-file");
    $("readButton").disabled = !(state.paymentFile && state.baseFile);
    $("reviewPanel").hidden = true;
    $("resultsPanel").hidden = true;
  }

  [["payments", "paymentsFile", "paymentsDrop"], ["base", "baseFile", "baseDrop"]].forEach(([kind, inputId, dropId]) => {
    const input = $(inputId);
    const drop = $(dropId);
    input.addEventListener("change", () => selectFile(kind, input.files && input.files[0]));
    ["dragenter", "dragover"].forEach(type => drop.addEventListener(type, event => {
      event.preventDefault();
      drop.classList.add("dragging");
    }));
    ["dragleave", "drop"].forEach(type => drop.addEventListener(type, event => {
      event.preventDefault();
      drop.classList.remove("dragging");
    }));
    drop.addEventListener("drop", event => {
      const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
      if (!file) return;
      const transfer = new DataTransfer();
      transfer.items.add(file);
      input.files = transfer.files;
      selectFile(kind, file);
    });
  });

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    if (typeof value === "number" && value > 20000 && value < 80000 && window.XLSX) {
      const p = XLSX.SSF.parse_date_code(value);
      if (p && p.y && p.m && p.d) return new Date(p.y, p.m - 1, p.d);
    }
    const s = String(value == null ? "" : value).trim();
    if (!s) return null;
    let m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
    if (m) {
      let year = Number(m[3]);
      if (year < 100) year += year >= 70 ? 1900 : 2000;
      const d = new Date(year, Number(m[2]) - 1, Number(m[1]));
      if (d.getFullYear() === year && d.getMonth() === Number(m[2]) - 1 && d.getDate() === Number(m[1])) return d;
    }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    const parsed = new Date(s);
    return Number.isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  function toISO(value) {
    const d = parseDate(value);
    if (!d) return "";
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }

  function parseAmount(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    let s = String(value == null ? "" : value).trim().replace(/[Kk][čc]\.?/g, "").replace(/CZK/gi, "").replace(/\s+/g, "");
    if (!s) return null;
    s = s.replace(/[^\d,.\-]/g, "");
    if (!s || !/\d/.test(s)) return null;
    if (s.includes(",") && s.includes(".")) {
      if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
      else s = s.replace(/,/g, "");
    } else if (s.includes(",")) {
      const parts = s.split(",");
      if (parts.length === 2 && parts[1].length <= 2) s = parts[0].replace(/\./g, "") + "." + parts[1];
      else s = s.replace(/,/g, "");
    } else if (s.includes(".")) {
      const parts = s.split(".");
      if (!(parts.length === 2 && parts[1].length <= 2)) s = s.replace(/\./g, "");
    }
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  function quarterMatch(value) {
    const s = String(value == null ? "" : value).trim();
    const m = s.match(/(?:^|\b)Q\s*([1-4])\s*[\/.-]\s*(\d{2,4})\b/i) || s.match(/(?:^|\b)([1-4])\s*Q\s*[\/.-]\s*(\d{2,4})\b/i);
    if (!m) return null;
    const q = Number(m[1]);
    let year = Number(m[2]);
    if (year < 100) year += year >= 70 ? 1900 : 2000;
    return { period: "Q" + q + "/" + year, quarter: q, year };
  }

  function dateMatches(line) {
    const re = /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g;
    return Array.from(String(line).matchAll(re), m => ({ text: m[0], index: m.index, end: m.index + m[0].length }));
  }

  function amountTokens(text) {
    const re = /(?:\d{1,3}(?:[ \u00a0\u202f]\d{3})+|\d+)(?:[,.]\d{1,2})?/g;
    return Array.from(String(text).matchAll(re), m => ({ text: m[0], index: m.index, end: m.index + m[0].length, value: parseAmount(m[0]) }))
      .filter(x => x.value != null);
  }

  function removeRanges(text, ranges) {
    let out = String(text);
    [...ranges].sort((a,b) => b.index - a.index).forEach(r => { out = out.slice(0, r.index) + " " + out.slice(r.end); });
    return out;
  }

  function monthsLate(dueValue, paidValue) {
    const due = parseDate(dueValue);
    const paid = parseDate(paidValue);
    if (!due || !paid || paid <= due) return 0;
    return Math.max(0, (paid.getFullYear() - due.getFullYear()) * 12 + paid.getMonth() - due.getMonth());
  }

  function normalizeQuarterLabel(value) {
    const m = quarterMatch(value);
    return m ? m.period : String(value || "").trim();
  }

  function parsePaymentText(text) {
    const lines = String(text || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const parsed = [];
    lines.forEach(line => {
      const dates = dateMatches(line);
      if (!dates.length) return;
      const date = toISO(dates[dates.length - 1].text);
      if (!date) return;
      const tokens = amountTokens(removeRanges(line, dates));
      if (!tokens.length) return;
      const vals = tokens.map(t => t.value);
      const uniqueVals = vals.filter((v, i) => vals.findIndex(x => Math.abs(x - v) < 0.005) === i);
      const amount = uniqueVals.length === 1 ? uniqueVals[0] : tokens[tokens.length - 1].value;
      if (amount == null || amount < 0) return;
      const normalized = { date, amount: Math.round(amount * 100) / 100, raw: line };
      if (!parsed.some(p => p.date === normalized.date && p.amount === normalized.amount && p.raw === normalized.raw)) parsed.push(normalized);
    });
    return parsed.sort((a,b) => a.date.localeCompare(b.date));
  }

  function extractRate(text) {
    const clean = String(text || "").replace(/<[^>]+>/g, " ").replace(/&nbsp;|&#160;/gi, " ");
    const m = clean.match(/sazb[aeu][\s\S]{0,120}?(\d{1,2}(?:[,.]\d+)?)\s*(?:‰|promile|promil)/i);
    return m ? parseAmount(m[1]) : null;
  }

  function normalizeRowText(row) {
    return row.map(v => {
      if (v instanceof Date) return toISO(v);
      if (typeof v === "number" && v > 20000 && v < 80000 && window.XLSX) {
        const date = parseDate(v);
        if (date) return toISO(date);
      }
      return String(v == null ? "" : v).trim();
    }).join(" | ");
  }

  function findQuarterHeaders(grid, rowIndex) {
    const start = Math.max(0, rowIndex - 5);
    let map = { due: 1, source: 2, base: 3, premium: 4 };
    for (let ri = start; ri < rowIndex; ri++) {
      const row = (grid[ri] || []).map(v => String(v == null ? "" : v).trim().toLowerCase());
      row.forEach((v, ci) => {
        if (!v) return;
        if (/splatnost/.test(v)) map.due = ci;
        if (/období pro výpočet/.test(v)) map.source = ci;
        if (/základ|vyměřovacích základů/.test(v)) map.base = ci + ((grid[ri][ci + 1] == null || grid[ri][ci + 1] === "") ? 1 : 0);
        if (/řádné pojistné/.test(v)) map.premium = ci;
      });
      const premiumCandidates = [];
      row.forEach((v, ci) => { if (v === "pojistné" || v === "pojistné ") premiumCandidates.push(ci); });
      if (premiumCandidates.length > 1) map.premium = premiumCandidates[premiumCandidates.length - 1];
      else if (premiumCandidates.length === 1 && row.some(v => /zaplaceno/.test(v))) map.premium = premiumCandidates[0];
      if (row.some(v => /zaplaceno/.test(v)) && row.some(v => /splatnost/.test(v))) {
        const pi = row.findIndex(v => /zaplaceno/.test(v));
        const before = row.map((v, i) => ({v, i})).filter(x => x.i < pi && /pojistné/.test(x.v));
        if (before.length) map.premium = before[before.length - 1].i;
      }
    }
    return map;
  }

  function parseQuartersFromGrid(grid) {
    const quarters = [];
    let rate = null;
    for (let ri = 0; ri < grid.length; ri++) {
      const row = grid[ri] || [];
      const cells = row.map(v => String(v == null ? "" : v).trim());
      const joined = cells.join(" ");
      if (/sazba/i.test(joined) && /promile|‰/i.test(joined)) {
        for (let ci = 0; ci < cells.length - 1; ci++) {
          if (/sazba/i.test(cells[ci])) {
            for (let j = ci + 1; j < Math.min(cells.length, ci + 4); j++) {
              const num = parseAmount(cells[j]);
              if (num != null && num > 0 && num < 100) { rate = num; break; }
            }
          }
        }
      }
      const periodCellIndex = row.findIndex(v => quarterMatch(v));
      if (periodCellIndex < 0) continue;
      const periodInfo = quarterMatch(row[periodCellIndex]);
      const headers = findQuarterHeaders(grid, ri);
      let dueDate = toISO(row[headers.due]);
      if (!dueDate) {
        for (let ci = periodCellIndex + 1; ci < row.length; ci++) {
          const date = toISO(row[ci]);
          if (date) { dueDate = date; break; }
        }
      }
      let sourcePeriod = row[headers.source] && quarterMatch(row[headers.source]) ? quarterMatch(row[headers.source]).period : "";
      let base = parseAmount(row[headers.base]);
      if (!(base > 0)) {
        const largeNums = row.map((v, ci) => ({ ci, n: parseAmount(v) })).filter(x => x.ci > periodCellIndex && x.n > 10000);
        if (largeNums.length) base = largeNums[0].n;
      }
      let premium = parseAmount(row[headers.premium]);
      if (!(premium >= 0 && premium < 10000)) {
        const candidates = row.map((v, ci) => ({ ci, n: parseAmount(v) }))
          .filter(x => x.ci > periodCellIndex && x.n != null && x.n >= 0 && x.n < 10000 && !(x.n > 0 && String(row[x.ci]).match(/^\d{4}$/)));
        if (candidates.length) premium = candidates[candidates.length - 1].n;
      }
      if ((!premium || premium <= 0) && base > 0 && rate > 0) premium = Math.round(base * rate / 1000);
      if (!dueDate && !(base > 0) && !(premium > 0)) continue;
      if (quarters.some(q => q.period === periodInfo.period)) continue;
      quarters.push({ period: periodInfo.period, dueDate, sourcePeriod, base: base > 0 ? base : null, premium: premium > 0 ? Math.round(premium) : null, raw: normalizeRowText(row) });
    }
    return { quarters: quarters.sort((a,b) => (a.dueDate || "").localeCompare(b.dueDate || "")), rate };
  }

  function parseQuartersFromText(text) {
    const lines = String(text || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const quarters = [];
    lines.forEach(line => {
      const matches = Array.from(line.matchAll(/(?:\bQ\s*[1-4]\s*[\/.-]\s*\d{2,4}\b|\b[1-4]\s*Q\s*[\/.-]\s*\d{2,4}\b)/gi));
      if (!matches.length) return;
      const firstInfo = quarterMatch(matches[0][0]);
      if (!firstInfo || quarters.some(q => q.period === firstInfo.period)) return;
      const dates = dateMatches(line);
      const dueDate = dates.length ? toISO(dates[0].text) : "";
      const remove = [...dates, ...matches.map(m => ({ index: m.index, end: m.index + m[0].length }))];
      const nums = amountTokens(removeRanges(line, remove)).filter(x => x.value != null);
      const baseItem = nums.find(x => x.value > 10000);
      const smallNums = nums.filter(x => x.value >= 0 && x.value < 10000);
      const premiumItem = smallNums.length ? smallNums[smallNums.length - 1] : null;
      let sourcePeriod = matches.length > 1 ? normalizeQuarterLabel(matches[1][0]) : "";
      const base = baseItem ? baseItem.value : null;
      const premium = premiumItem && premiumItem.value > 0 ? Math.round(premiumItem.value) : null;
      if (!dueDate && !(base > 0) && !(premium > 0)) return;
      quarters.push({ period: firstInfo.period, dueDate, sourcePeriod, base, premium, raw: line });
    });
    return quarters.sort((a,b) => (a.dueDate || "").localeCompare(b.dueDate || ""));
  }

  function detectRateFromGrid(grid) {
    for (const row of grid) {
      const vals = row.map(v => String(v == null ? "" : v).trim());
      if (vals.some(v => /sazba/i.test(v))) {
        const i = vals.findIndex(v => /sazba/i.test(v));
        for (let j = i + 1; j < Math.min(vals.length, i + 5); j++) {
          const val = parseAmount(vals[j]);
          if (val != null && val > 0 && val < 100) return val;
        }
      }
    }
    return null;
  }

  async function getOcrWorker() {
    if (!window.Tesseract || !Tesseract.createWorker) throw new Error("Knihovna OCR není dostupná. Zkontroluj připojení k internetu a opakuj načtení.");
    if (state.ocrWorker) return state.ocrWorker;
    setStatus("Připravuji OCR", "Stahuji pouze jazykový model češtiny a angličtiny. Obsah screenshotu zůstává v prohlížeči.", true);
    state.ocrWorker = await Tesseract.createWorker("ces+eng", 1, {
      logger: function(m) {
        if (m.status === "recognizing text") {
          const pct = Math.round((m.progress || 0) * 100);
          $("statusText").textContent = "Lokální OCR rozpoznává text: " + pct + " %";
        }
      },
      langPath: "https://tessdata.projectnaptha.com/4.0.0_best",
      gzip: true
    });
    return state.ocrWorker;
  }

  async function ocrImageBlob(blob, label) {
    const worker = await getOcrWorker();
    setStatus("Rozpoznávám text", label || "Probíhá lokální OCR obrázku.", true);
    let input = blob;
    if (window.createImageBitmap) {
      try {
        const bitmap = await createImageBitmap(blob);
        const factor = Math.min(3, Math.max(1, 1800 / Math.max(bitmap.width, bitmap.height)));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(bitmap.width * factor);
        canvas.height = Math.round(bitmap.height * factor);
        canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
        input = canvas;
        bitmap.close();
      } catch (_) {}
    }
    const result = await worker.recognize(input);
    return result && result.data ? result.data.text || "" : "";
  }

  async function parseImageForText(file) {
    return ocrImageBlob(file, "Čtu screenshot v prohlížeči.");
  }

  async function parsePdf(file) {
    if (!window.pdfjsLib) throw new Error("Knihovna PDF není dostupná. Zkontroluj připojení k internetu.");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const allLines = [];
    for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
      setStatus("Čtu PDF", "Zpracovávám stránku " + pageNo + " z " + pdf.numPages + " lokálně.", true);
      const page = await pdf.getPage(pageNo);
      const content = await page.getTextContent();
      const grouped = new Map();
      (content.items || []).forEach(item => {
        const txt = String(item.str || "").trim();
        if (!txt) return;
        const y = Math.round(item.transform && item.transform[5] ? item.transform[5] : 0);
        if (!grouped.has(y)) grouped.set(y, []);
        grouped.get(y).push({ x: item.transform && item.transform[4] ? item.transform[4] : 0, txt });
      });
      const lines = Array.from(grouped.entries()).sort((a,b) => b[0] - a[0]).map(entry => entry[1].sort((a,b) => a.x - b.x).map(x => x.txt).join(" "));
      const pageText = lines.join("\n").trim();
      if (pageText.length > 30) allLines.push(pageText);
      else {
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        allLines.push(await ocrImageBlob(canvas, "PDF nemá čitelnou textovou vrstvu; rozpoznávám stránku " + pageNo + " pomocí OCR."));
      }
    }
    return allLines.join("\n\n");
  }

  async function readTextFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (["png","jpg","jpeg","webp","gif"].includes(ext) || file.type.startsWith("image/")) return parseImageForText(file);
    if (ext === "pdf" || file.type === "application/pdf") return parsePdf(file);
    if (["xlsx","xls","csv"].includes(ext) && window.XLSX) {
      const data = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const textLines = [];
      data.SheetNames.forEach(name => {
        const grid = XLSX.utils.sheet_to_json(data.Sheets[name], { header: 1, raw: true, defval: "" });
        grid.forEach(row => { if (row.some(v => v !== "" && v != null)) textLines.push(normalizeRowText(row)); });
      });
      return textLines.join("\n");
    }
    if (ext === "docx") {
      if (!window.mammoth) throw new Error("Knihovna Word není dostupná.");
      const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
      return result.value || "";
    }
    if (ext === "doc") throw new Error("Starý formát .doc není v prohlížeči bezpečně podporován. Ulož jej jako .docx nebo PDF a vyber znovu.");
    const text = await file.text();
    if (ext === "html" || ext === "htm" || /<html[\s>]/i.test(text.slice(0, 400))) {
      const doc = new DOMParser().parseFromString(text, "text/html");
      return (doc.body ? doc.body.innerText : doc.documentElement.innerText) || text.replace(/<[^>]+>/g, " ");
    }
    return text;
  }

  async function readGridFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext) || !window.XLSX) return null;
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    let allQuarters = [], allLines = [], rate = null;
    workbook.SheetNames.forEach(name => {
      const grid = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" });
      const parsed = parseQuartersFromGrid(grid);
      allQuarters = allQuarters.concat(parsed.quarters);
      if (rate == null) rate = parsed.rate || detectRateFromGrid(grid);
      grid.forEach(row => { if (row.some(v => v !== "" && v != null)) allLines.push(normalizeRowText(row)); });
    });
    const unique = [];
    allQuarters.forEach(q => { if (!unique.some(x => x.period === q.period)) unique.push(q); });
    return { quarters: unique.sort((a,b) => (a.dueDate || "").localeCompare(b.dueDate || "")), text: allLines.join("\n"), rate };
  }

  async function processInputs() {
    if (!state.paymentFile || !state.baseFile) return;
    $("readButton").disabled = true;
    $("reviewPanel").hidden = true; $("resultsPanel").hidden = true;
    state.payments = []; state.quarters = []; state.results = null;
    try {
      setStatus("Čtu screenshot plateb", "Soubor se zpracovává lokálně.", true);
      state.paymentOcrText = await parseImageForText(state.paymentFile);
      state.payments = parsePaymentText(state.paymentOcrText);
      setStatus("Čtu základní dokument", "Extrahuji tabulku čtvrtletí a předpisů.", true);
      const gridResult = await readGridFile(state.baseFile);
      state.baseText = await readTextFile(state.baseFile);
      state.sourceRate = gridResult && gridResult.rate != null ? gridResult.rate : extractRate(state.baseText);
      state.sourceRateLabel = state.sourceRate != null ? "Sazba rozpoznána ze základního dokumentu" : "Sazba nebyla rozpoznána";
      state.quarters = gridResult && gridResult.quarters.length ? gridResult.quarters : parseQuartersFromText(state.baseText);
      const detectedTextRate = extractRate(state.baseText);
      if (state.sourceRate == null && detectedTextRate != null) state.sourceRate = detectedTextRate;
      if (state.sourceRate != null) {
        $("rateInput").value = String(state.sourceRate);
        $("rateSource").textContent = state.sourceRateLabel;
      } else {
        $("rateInput").value = "";
        $("rateSource").textContent = "Sazba nebyla rozpoznána; doplň ji pouze pokud v podkladu chybí pojistné.";
      }
      renderInputTables();
      $("sourceText").textContent = (state.baseText || "Z dokumentu nebyl získán žádný text.").slice(0, 35000);
      $("reviewPanel").hidden = false;
      $("reviewPanel").scrollIntoView({ behavior: "smooth", block: "start" });
      const notes = [];
      if (!state.payments.length) notes.push("Nepodařilo se spolehlivě rozpoznat platby.");
      if (!state.quarters.length) notes.push("Nepodařilo se rozpoznat čtvrtletní předpisy.");
      setStatus(notes.length ? "Podklady načteny s upozorněním" : "Podklady načteny", notes.length ? notes.join(" ") + " Doplň nebo oprav tabulky ručně." : "Zkontroluj částky, data a čtvrtletí před výpočtem.", false);
      $("spinner").hidden = false;
      $("readButton").disabled = false;
    } catch (error) {
      console.error(error);
      setStatus("Zpracování se nepodařilo", error && error.message ? error.message : "Soubor se nepodařilo načíst.", false);
      $("readButton").disabled = false;
    }
  }

  function element(tag, props, text) {
    const node = document.createElement(tag);
    Object.entries(props || {}).forEach(([key, value]) => {
      if (key === "className") node.className = value;
      else if (key === "textContent") node.textContent = value;
      else node.setAttribute(key, value);
    });
    if (text != null) node.textContent = text;
    return node;
  }

  function addInputCell(tr, type, value, label, handler, step) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = type;
    input.value = value == null ? "" : value;
    input.setAttribute("aria-label", label);
    if (type === "number") {
      input.min = "0";
      input.step = step || "1";
    }
    input.addEventListener("change", () => handler(input.value));
    td.appendChild(input);
    tr.appendChild(td);
    return input;
  }

  function renderInputTables() {
    const pBody = $("paymentsBody"); pBody.replaceChildren();
    state.payments.forEach((payment, i) => {
      const tr = document.createElement("tr");
      addInputCell(tr, "date", payment.date, "Datum platby " + (i + 1), value => { payment.date = value; invalidateResults(); });
      addInputCell(tr, "number", payment.amount, "Částka platby " + (i + 1), value => { payment.amount = parseAmount(value); invalidateResults(); }, "0.01");
      const rawTd = document.createElement("td");
      const raw = element("span", { className: "raw-text", title: payment.raw || "" }, payment.raw || "Ruční záznam");
      rawTd.appendChild(raw); tr.appendChild(rawTd);
      const delTd = document.createElement("td");
      const del = element("button", { type: "button", className: "delete-row", "aria-label": "Odstranit platbu" }, "×");
      del.addEventListener("click", () => { state.payments.splice(i,1); renderInputTables(); invalidateResults(); });
      delTd.appendChild(del); tr.appendChild(delTd);
      pBody.appendChild(tr);
    });
    $("paymentsEmpty").hidden = state.payments.length > 0;

    const qBody = $("quartersBody"); qBody.replaceChildren();
    state.quarters.forEach((quarter, i) => {
      const tr = document.createElement("tr");
      addInputCell(tr, "text", quarter.period, "Čtvrtletí " + (i + 1), value => { quarter.period = value; invalidateResults(); });
      addInputCell(tr, "date", quarter.dueDate, "Splatnost " + (i + 1), value => { quarter.dueDate = value; invalidateResults(); });
      addInputCell(tr, "number", quarter.base, "Výpočtový základ " + (i + 1), value => {
        quarter.base = parseAmount(value);
        recalcMissingPremium(quarter);
        renderInputTables();
        invalidateResults();
      }, "1");
      addInputCell(tr, "number", quarter.premium, "Řádné pojistné " + (i + 1), value => { quarter.premium = parseAmount(value); invalidateResults(); }, "1");
      const delTd = document.createElement("td");
      const del = element("button", { type: "button", className: "delete-row", "aria-label": "Odstranit čtvrtletí" }, "×");
      del.addEventListener("click", () => { state.quarters.splice(i,1); renderInputTables(); invalidateResults(); });
      delTd.appendChild(del); tr.appendChild(delTd);
      qBody.appendChild(tr);
    });
    $("quartersEmpty").hidden = state.quarters.length > 0;
  }

  function recalcMissingPremium(quarter) {
    const rate = parseAmount($("rateInput").value);
    if (!(quarter.premium > 0) && quarter.base > 0 && rate > 0) quarter.premium = Math.round(quarter.base * rate / 1000);
  }

  $("rateInput").addEventListener("change", () => {
    state.quarters.forEach(recalcMissingPremium);
    renderInputTables();
    invalidateResults();
  });

  function invalidateResults() {
    $("resultsPanel").hidden = true;
  }

  $("addPayment").addEventListener("click", () => {
    state.payments.push({ date: isoToday(), amount: "", raw: "Ruční záznam" });
    renderInputTables(); invalidateResults();
  });
  $("addQuarter").addEventListener("click", () => {
    state.quarters.push({ period: "", dueDate: "", base: "", premium: "", sourcePeriod: "", raw: "Ruční záznam" });
    renderInputTables(); invalidateResults();
  });
  $("readButton").addEventListener("click", processInputs);
  $("backToFiles").addEventListener("click", () => {
    $("reviewPanel").hidden = true;
    $("resultsPanel").hidden = true;
    $("paymentsFile").focus();
  });
  $("editDataButton").addEventListener("click", () => {
    $("resultsPanel").hidden = true;
    $("reviewPanel").hidden = false;
    $("reviewPanel").scrollIntoView({ behavior: "smooth", block: "start" });
  });

  function validISODate(value) {
    if (!value) return false;
    const d = parseDate(value);
    return !!d && toISO(d) === value;
  }

  function computeResults() {
    const warnings = [];
    const asOf = $("asOfDate").value || isoToday();
    const quarters = state.quarters.map((q, i) => ({
      ...q,
      index: i,
      premium: Math.round(Number(q.premium) || 0),
      base: Number(q.base) || 0,
      remaining: Math.round(Number(q.premium) || 0),
      paidPrincipal: 0,
      assessedFeeOnPaid: 0,
      paidFee: 0,
      unpaidFeeFromPaid: 0,
      feeOnOutstanding: 0,
      feeUnpaid: 0,
      allocations: []
    })).filter(q => q.period || q.dueDate || q.premium > 0);
    const allPayments = state.payments.map((p,i) => ({ ...p, index:i, amount: Number(p.amount), date: p.date }));
    allPayments.filter(p => validISODate(p.date) && Number.isFinite(p.amount) && p.amount === 0).forEach(p => {
      warnings.push("Řádek plateb ze dne " + dateFmt(p.date) + " má ve sloupci Vyřízeno 0 Kč; nezapočítává se jako uhrazená platba.");
    });
    const payments = allPayments.filter(p => p.amount > 0 || !p.date);
    if (!quarters.length) throw new Error("Základní dokument neobsahuje žádné čtvrtletní předpisy. Přidej je do tabulky.");
    if (!payments.length) throw new Error("Nenašel jsem žádné kladné vyřízené platby. Řádky s částkou 0 Kč se nezapočítávají.");
    quarters.forEach(q => {
      if (!q.period) warnings.push("U jednoho předpisu chybí označení čtvrtletí.");
      if (!validISODate(q.dueDate)) warnings.push("U čtvrtletí " + (q.period || "(bez označení)") + " chybí nebo není platné datum splatnosti.");
      if (!(q.premium >= 0) || !Number.isFinite(q.premium) || q.premium <= 0) warnings.push("U čtvrtletí " + (q.period || "(bez označení)") + " chybí řádné pojistné.");
    });
    payments.forEach(p => {
      if (!validISODate(p.date)) warnings.push("Jedna platba nemá platné datum.");
      if (!Number.isFinite(p.amount) || p.amount <= 0) warnings.push("Jedna platba nemá kladnou částku.");
      if (validISODate(p.date) && p.date > asOf) warnings.push("Platba ze dne " + dateFmt(p.date) + " je po zvoleném datu výpočtu.");
    });
    if (warnings.some(w => /chybí|není platné|nemá platné|nemá kladnou/.test(w))) {
      throw new Error("Oprav chybějící hodnoty: " + warnings.filter(w => /chybí|není platné|nemá platné|nemá kladnou/.test(w)).join(" "));
    }
    if (warnings.some(w => /po zvoleném datu výpočtu/.test(w))) {
      throw new Error("Některé platby jsou po zvoleném datu výpočtu. Posuň datum výpočtu na den poslední zahrnované platby nebo později.");
    }

    quarters.sort((a,b) => a.dueDate.localeCompare(b.dueDate));
    payments.sort((a,b) => a.date.localeCompare(b.date) || a.index - b.index);
    const detail = [];
    let overpayment = 0;

    payments.forEach(payment => {
      let remainingPayment = payment.amount;
      let principalAllocated = 0;
      let feesPaidThisPayment = 0;
      for (const q of quarters) {
        if (remainingPayment <= 0.0001) break;
        if (q.remaining <= 0) continue;
        const allocated = Math.min(q.remaining, remainingPayment);
        const delay = monthsLate(q.dueDate, payment.date);
        const fee = Math.round(allocated * delay / 10);
        q.remaining = Math.max(0, q.remaining - allocated);
        q.paidPrincipal += allocated;
        q.assessedFeeOnPaid += fee;
        q.unpaidFeeFromPaid += fee;
        q.feeUnpaid += fee;
        principalAllocated += allocated;
        remainingPayment -= allocated;
        const item = { date: payment.date, period: q.period, type: "Úhrada řádného pojistného", amount: allocated, delayMonths: delay, fee, sourceAmount: payment.amount };
        q.allocations.push(item); detail.push(item);
      }
      if (remainingPayment > 0.0001) {
        for (const q of quarters) {
          if (remainingPayment <= 0.0001) break;
          if (q.feeUnpaid <= 0) continue;
          const paidFee = Math.min(q.feeUnpaid, remainingPayment);
          q.feeUnpaid = Math.max(0, q.feeUnpaid - paidFee);
          q.paidFee += paidFee;
          feesPaidThisPayment += paidFee;
          remainingPayment -= paidFee;
          const item = { date: payment.date, period: q.period, type: "Úhrada navýšení za prodlení", amount: paidFee, delayMonths: null, fee: 0, sourceAmount: payment.amount };
          q.allocations.push(item); detail.push(item);
        }
      }
      if (remainingPayment > 0.0001) {
        overpayment += remainingPayment;
        detail.push({ date: payment.date, period: "—", type: "Nezařazený přeplatek", amount: remainingPayment, delayMonths: null, fee: 0, sourceAmount: payment.amount });
        warnings.push("Z platby ze dne " + dateFmt(payment.date) + " zbylo " + currencyFmt(remainingPayment) + " po započtení pojistného i evidovaného navýšení.");
      }
    });

    quarters.forEach(q => {
      if (q.remaining > 0) {
        const delay = monthsLate(q.dueDate, asOf);
        const fee = Math.round(q.remaining * delay / 10);
        q.feeOnOutstanding = fee;
        q.assessedFeeOnPaid += 0;
        q.feeUnpaid += fee;
        detail.push({ date: asOf, period: q.period, type: "Navýšení z neuhrazeného pojistného k datu výpočtu", amount: q.remaining, delayMonths: delay, fee, sourceAmount: null });
      }
      q.totalAssessedFee = q.assessedFeeOnPaid + q.feeOnOutstanding;
      q.paidPrincipal = Math.round(q.paidPrincipal);
      q.remaining = Math.round(q.remaining);
      q.assessedFeeOnPaid = Math.round(q.assessedFeeOnPaid);
      q.feeOnOutstanding = Math.round(q.feeOnOutstanding);
      q.paidFee = Math.round(q.paidFee);
      q.feeUnpaid = Math.round(q.feeUnpaid);
    });

    const totals = {
      premium: quarters.reduce((sum,q) => sum + q.premium, 0),
      principalPaid: quarters.reduce((sum,q) => sum + q.paidPrincipal, 0),
      principalOpen: quarters.reduce((sum,q) => sum + q.remaining, 0),
      feeAssessed: quarters.reduce((sum,q) => sum + q.totalAssessedFee, 0),
      feePaid: quarters.reduce((sum,q) => sum + q.paidFee, 0),
      feeOpen: quarters.reduce((sum,q) => sum + q.feeUnpaid, 0),
      paymentTotal: payments.reduce((sum,p) => sum + (Number.isFinite(p.amount) ? p.amount : 0), 0),
      overpayment: Math.round(overpayment),
      due: quarters.reduce((sum,q) => sum + q.remaining + q.feeUnpaid, 0)
    };
    const seen = new Map();
    payments.forEach(p => {
      const key = p.date + "|" + p.amount;
      seen.set(key, (seen.get(key) || 0) + 1);
    });
    if (Array.from(seen.values()).some(n => n > 1)) warnings.push("Ve vstupu jsou stejné částky se stejným datem vícekrát. Ověř, že nejde o duplicitně načtenou platbu.");
    const detectedRate = parseAmount($("rateInput").value);
    const impliedRates = quarters.filter(q => q.base > 0 && q.premium > 0).map(q => q.premium / q.base * 1000).filter(x => Number.isFinite(x) && x > 0);
    if (detectedRate != null && impliedRates.length) {
      const meanRate = impliedRates.reduce((sum,x) => sum + x, 0) / impliedRates.length;
      if (Math.abs(detectedRate - meanRate) > 0.4) warnings.push("Sazba " + detectedRate.toLocaleString("cs-CZ") + " ‰ se výrazně liší od poměru základu a předpisů v tabulce (orientačně " + meanRate.toLocaleString("cs-CZ", { maximumFractionDigits: 2 }) + " ‰). Ověř, zda sazba i čtvrtletní předpisy patří ke stejnému případu.");
    }
    if (state.sourceRate == null && !($("rateInput").value.trim())) {
      const anyMissingPremium = quarters.some(q => !(q.premium > 0));
      if (anyMissingPremium) warnings.push("Chybí sazba nebo předpis pojistného; výpočet nepoužívej, dokud údaje nedoplníš.");
    }

    return { asOf, quarters, payments, detail, totals, warnings };
  }

  function renderResults(result) {
    $("summaryGrid").replaceChildren();
    const kpis = [
      ["Předpis řádného pojistného", result.totals.premium, "Součet načtených čtvrtletí"],
      ["Uhrazené řádné pojistné", result.totals.principalPaid, "Platby přiřazené k pojistnému"],
      ["Zbývající pojistné", result.totals.principalOpen, "Neuhrazené jistiny"],
      ["Zbývající navýšení", result.totals.feeOpen, "Po odečtení uhrazeného navýšení"]
    ];
    kpis.forEach(k => {
      const card = element("div", { className: "kpi-card" });
      card.append(element("span", {}, k[0]), element("strong", {}, currencyFmt(k[1])), element("small", {}, k[2]));
      $("summaryGrid").append(card);
    });
    const qb = $("resultQuartersBody"); qb.replaceChildren();
    result.quarters.forEach(q => {
      const tr = document.createElement("tr");
      [
        q.period, dateFmt(q.dueDate), currencyFmt(q.premium), currencyFmt(q.paidPrincipal),
        currencyFmt(q.remaining), currencyFmt(q.totalAssessedFee), currencyFmt(q.paidFee), currencyFmt(q.feeUnpaid)
      ].forEach(v => tr.appendChild(element("td", {}, v)));
      qb.append(tr);
    });
    const ab = $("allocationsBody"); ab.replaceChildren();
    result.detail.forEach(item => {
      const tr = document.createElement("tr");
      [
        dateFmt(item.date), item.period, item.type,
        item.type === "Navýšení z neuhrazeného pojistného k datu výpočtu" ? currencyFmt(item.amount) : currencyFmt(item.amount),
        item.delayMonths == null ? "—" : String(item.delayMonths),
        currencyFmt(item.fee)
      ].forEach(v => tr.appendChild(element("td", {}, v)));
      ab.append(tr);
    });
    $("warningList").replaceChildren();
    result.warnings.forEach(w => $("warningList").append(element("li", {}, w)));
    $("resultWarnings").hidden = result.warnings.length === 0;
    $("resultsPanel").hidden = false;
    $("resultsPanel").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  $("calculateButton").addEventListener("click", () => {
    try {
      state.results = computeResults();
      renderResults(state.results);
      setStatus("Výpočet dokončen", "Zkontroluj souhrn, varování a rozpis započtení. Excel stáhni až po ověření výsledků.", false);
      $("spinner").hidden = false;
    } catch (error) {
      setStatus("Výpočet nelze dokončit", error.message || "Zkontroluj vstupní hodnoty.", false);
      $("spinner").hidden = true;
    }
  });

  function styleHeader(row) {
    row.height = 24;
    row.eachCell(cell => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF102E46" } };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: "FFCAD8DF" } } };
    });
  }

  function styleBodyRows(sheet, startRow) {
    for (let r = startRow; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      row.alignment = { vertical: "top", wrapText: true };
      row.eachCell(cell => {
        cell.border = { bottom: { style: "hair", color: { argb: "FFE3EAEE" } } };
        if (r % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7FAFB" } };
        if (typeof cell.value === "number") cell.numFmt = "#,##0;[Red]-#,##0";
      });
    }
  }

  async function downloadWorkbook() {
    if (!state.results) return;
    if (!window.ExcelJS) {
      setStatus("Exportní knihovna není dostupná", "Zkontroluj připojení k internetu a zkus export znovu.", false);
      return;
    }
    const r = state.results;
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Výpočet zákonného pojištění";
    workbook.subject = "Lokálně vypočtené přiřazení plateb a navýšení";
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.calcProperties.fullCalcOnLoad = true;

    const summary = workbook.addWorksheet("Souhrn", { views: [{ state: "frozen", ySplit: 3 }] });
    summary.columns = [{width:28},{width:19},{width:19},{width:20},{width:19},{width:19}];
    summary.mergeCells("A1:F1");
    summary.getCell("A1").value = "VÝPOČET ZÁKONNÉHO POJIŠTĚNÍ";
    summary.getCell("A1").font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    summary.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF102E46" } };
    summary.getRow(1).height = 32;
    summary.mergeCells("A2:F2");
    summary.getCell("A2").value = "Datum výpočtu: " + dateFmt(r.asOf) + " | Způsob započtení: nejstarší neuhrazené řádné pojistné jako první";
    summary.getCell("A2").font = { color: { argb: "FF526775" }, italic: true, size: 10 };
    summary.addRow([]);
    const sumHeader = summary.addRow(["Ukazatel","Částka (Kč)","Poznámka"]);
    styleHeader(sumHeader);
    [
      ["Předepsané řádné pojistné", r.totals.premium, "Součet čtvrtletních předpisů"],
      ["Celkem skutečně zaplaceno", r.totals.paymentTotal, "Součet částek ze screenshotu"],
      ["Započteno na řádné pojistné", r.totals.principalPaid, "Přiřazení nejstaršímu nedoplatku"],
      ["Zbývající řádné pojistné", r.totals.principalOpen, "Jistina po započtení plateb"],
      ["Navýšení za prodlení celkem", r.totals.feeAssessed, "Navýšení z uhrazených částek a otevřené jistiny"],
      ["Uhrazené navýšení", r.totals.feePaid, "Platby použité až po úhradě řádného pojistného"],
      ["Zbývající navýšení", r.totals.feeOpen, "Neuhrazené navýšení za prodlení"],
      ["Celkem zbývá uhradit", r.totals.due, "Zbývající pojistné a neuhrazené navýšení"],
      ["Nezařazený přeplatek", r.totals.overpayment, "Platby nad rámec všech předpisů a evidovaného navýšení"]
    ].forEach(row => summary.addRow(row));
    styleBodyRows(summary, 5);
    summary.getColumn(2).numFmt = "#,##0;[Red]-#,##0";
    summary.addRow([]);
    const notice = summary.addRow(["Poznámka", "Navýšení za prodlení je počítáno podle vzorce ze vzorového sešitu: 10 % za každý kalendářní měsíc mezi měsícem splatnosti a měsícem platby (pouze pokud je platba po splatnosti), s zaokrouhlením na celé Kč."]);
    notice.getCell(1).font = { bold: true, color: { argb: "FF77521B" } };
    notice.getCell(2).alignment = { wrapText: true, vertical: "top" };
    summary.mergeCells("B" + notice.number + ":F" + notice.number);
    notice.height = 37;

    const legacy = workbook.addWorksheet("List1", { views: [{ state: "frozen", ySplit: 4 }] });
    legacy.columns = [{width:13},{width:14},{width:15},{width:18},{width:14},{width:14},{width:14},{width:13},{width:12},{width:14},{width:14}];
    const topRow = legacy.addRow(["IČ", "", "", "Vznik pojištění", "", "", "", "sazba:", Number($("rateInput").value) || "", "promile", ""]);
    topRow.getCell(1).font = { bold: true, color: { argb: "FF102E46" } };
    topRow.getCell(8).font = { bold: true, color: { argb: "FF102E46" } };
    topRow.getCell(9).numFmt = "0.0";
    legacy.addRow([]);
    const legacyHeader = legacy.addRow(["Pojistné", "Splatnost", "Z vyměřovacích základů", "", "Pojistné", "Zaplaceno", "částka", "Zpoždění", "Zvýšení", "částka", "Celkem"]);
    styleHeader(legacyHeader);
    const legacySub = legacy.addRow(["na kvartál", "", "", "", "", "dne", "", "v měsících", "%", "částka", ""]);
    legacySub.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FF526775" }, size: 9 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF0F3" } };
    });
    const legacyRows = [];
    const shortQuarter = value => {
      const m = String(value || "").match(/^Q([1-4])\/(\d{4})$/i);
      return m ? m[1] + "Q/" + m[2].slice(-2) : String(value || "");
    };
    r.quarters.forEach(q => {
      const paymentAllocations = q.allocations.filter(a => a.type === "Úhrada řádného pojistného");
      const openFeeMonths = monthsLate(q.dueDate, r.asOf);
      const makeRow = (first, allocation, openFee) => {
        const delay = allocation ? allocation.delayMonths : (openFee ? openFeeMonths : null);
        const fee = allocation ? allocation.fee : (openFee ? q.feeOnOutstanding : 0);
        const row = [
          first ? shortQuarter(q.period) : "",
          first ? (parseDate(q.dueDate) || "") : "",
          first ? shortQuarter(q.sourcePeriod) : "",
          first ? (q.base || "") : "",
          first ? q.premium : "",
          allocation ? (parseDate(allocation.date) || "") : "",
          allocation ? Math.round(allocation.amount) : "",
          delay == null ? "" : String(delay),
          delay == null ? "" : delay / 10,
          Math.round(fee || 0),
          first ? Math.round(q.premium + (fee || 0)) : Math.round(fee || 0)
        ];
        legacyRows.push(row);
      };
      if (paymentAllocations.length) {
        paymentAllocations.forEach((allocation, i) => makeRow(i === 0, allocation, false));
        if (q.remaining > 0) makeRow(false, null, true);
      } else {
        makeRow(true, null, q.remaining > 0);
      }
      legacyRows.push(Array(11).fill(""));
    });
    legacyRows.forEach(values => {
      const row = legacy.addRow(values);
      row.alignment = { vertical: "middle", wrapText: true };
      row.eachCell(cell => {
        cell.border = { bottom: { style: "hair", color: { argb: "FFE3EAEE" } } };
        if (row.number % 2 === 0) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7FAFB" } };
      });
      row.getCell(2).numFmt = "dd.mm.yyyy";
      row.getCell(6).numFmt = "dd.mm.yyyy";
      row.getCell(9).numFmt = "0%";
      [4,5,7,10,11].forEach(ci => { row.getCell(ci).numFmt = "#,##0;[Red]-#,##0"; });
    });
    const legacyTotal = legacy.addRow(["Celkem", "", "", "", r.totals.premium, "", r.totals.principalPaid, "", "", r.totals.feeAssessed, r.totals.premium + r.totals.feeAssessed]);
    legacyTotal.eachCell(cell => {
      cell.font = { bold: true, color: { argb: "FF102E46" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAF0F3" } };
      cell.border = { top: { style: "thin", color: { argb: "FF102E46" } } };
    });
    [5,7,10,11].forEach(ci => { legacyTotal.getCell(ci).numFmt = "#,##0;[Red]-#,##0"; });
    legacy.autoFilter = { from: "A3", to: "K" + (legacy.rowCount - 1) };

    const calc = workbook.addWorksheet("Přehled čtvrtletí", { views: [{ state: "frozen", ySplit: 1 }] });
    calc.columns = [
      {width:14},{width:14},{width:19},{width:21},{width:20},
      {width:22},{width:23},{width:20},{width:22},{width:20}
    ];
    const qhead = calc.addRow(["Čtvrtletí","Splatnost","Období pro výpočet","Výpočtový základ (Kč)","Řádné pojistné (Kč)","Uhrazeno na pojistné (Kč)","Nedoplatek pojistného (Kč)","Navýšení celkem (Kč)","Navýšení uhrazeno (Kč)","Navýšení zbývá (Kč)"]); styleHeader(qhead);
    r.quarters.forEach(q => calc.addRow([
      q.period, dateFmt(q.dueDate), q.sourcePeriod || "", q.base || 0, q.premium, q.paidPrincipal, q.remaining,
      q.totalAssessedFee, q.paidFee, q.feeUnpaid
    ]));
    styleBodyRows(calc, 2);
    calc.autoFilter = { from: "A1", to: "J" + calc.rowCount };

    const alloc = workbook.addWorksheet("Započtení plateb", { views: [{ state: "frozen", ySplit: 1 }] });
    const allocHeads = ["Datum", "Čtvrtletí", "Typ započtení", "Částka (Kč)", "Prodlení (měsíce)", "Navýšení (Kč)", "Původní částka platby (Kč)"];
    alloc.addRow(allocHeads); styleHeader(alloc.getRow(1));
    alloc.columns = [{width:14},{width:14},{width:44},{width:18},{width:20},{width:18},{width:26}];
    r.detail.forEach(d => alloc.addRow([dateFmt(d.date), d.period, d.type, Math.round(d.amount), d.delayMonths == null ? "" : d.delayMonths, Math.round(d.fee), d.sourceAmount == null ? "" : Math.round(d.sourceAmount)]));
    styleBodyRows(alloc, 2);
    alloc.autoFilter = { from: "A1", to: "G" + alloc.rowCount };

    const inputs = workbook.addWorksheet("Vstupní data", { views: [{ state: "frozen", ySplit: 1 }] });
    inputs.addRow(["PŮVODNÍ PLATBY"]); styleHeader(inputs.getRow(1));
    inputs.addRow(["Datum platby","Částka (Kč)","Rozpoznaný řádek"]);
    styleHeader(inputs.getRow(2));
    r.payments.forEach(p => inputs.addRow([dateFmt(p.date), Number(p.amount), p.raw || "Ruční záznam"]));
    const qStart = inputs.rowCount + 2;
    inputs.addRow([]);
    inputs.addRow(["PŘEDPISY POJISTNÉHO"]); styleHeader(inputs.getRow(inputs.rowCount));
    inputs.addRow(["Čtvrtletí","Splatnost","Období pro výpočet","Základ (Kč)","Řádné pojistné (Kč)","Původní řádek"]);
    styleHeader(inputs.getRow(inputs.rowCount));
    r.quarters.forEach(q => inputs.addRow([q.period, dateFmt(q.dueDate), q.sourcePeriod || "", q.base || "", q.premium, q.raw || "Ruční záznam"]));
    inputs.columns = [{width:18},{width:20},{width:22},{width:20},{width:22},{width:72}];
    styleBodyRows(inputs, 3);
    inputs.autoFilter = { from: "A2", to: "C" + (2 + r.payments.length) };

    const warnings = workbook.addWorksheet("Kontrola");
    warnings.columns = [{width:22},{width:100}];
    warnings.addRow(["KONTROLA VÝPOČTU","VÝSLEDEK"]); styleHeader(warnings.getRow(1));
    warnings.addRow(["Datum výpočtu",dateFmt(r.asOf)]);
    warnings.addRow(["Sazba za prodlení","10 % z příslušné částky za každý kalendářní měsíc prodlení; podle vzorců ve vzorovém sešitu."]);
    warnings.addRow(["Pravidlo přiřazení","Platby se započítávají nejprve na nejstarší neuhrazené řádné pojistné. Navýšení se hradí až po úhradě řádného pojistného."]);
    warnings.addRow(["Zdrojové soubory","Obsah vstupů se zpracoval v prohlížeči; do repozitáře ani na aplikační server nebyl poslán."]);
    if (r.warnings.length) {
      warnings.addRow([]);
      warnings.addRow(["UPOZORNĚNÍ","Ověř před použitím"]);
      styleHeader(warnings.getRow(warnings.rowCount));
      r.warnings.forEach(w => warnings.addRow(["Upozornění", w]));
    } else {
      warnings.addRow(["Upozornění","Bez automaticky zjištěných nesrovnalostí."]);
    }
    styleBodyRows(warnings, 2);

    setStatus("Připravuji Excel", "Vytvářím výstupní sešit pouze v prohlížeči.", true);
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "vypocet-zakonneho-pojisteni-" + r.asOf + ".xlsx";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    setStatus("Excel vytvořen", "Soubor byl vygenerován lokálně a stažen do tvého počítače.", false);
    $("spinner").hidden = false;
  }

  $("downloadButton").addEventListener("click", () => downloadWorkbook().catch(error => {
    console.error(error);
    setStatus("Export se nepodařil", error && error.message ? error.message : "Excel se nepodařilo vytvořit.", false);
    $("spinner").hidden = true;
  }));

  if (!("createImageBitmap" in window)) {
    // The OCR flow has a canvas fallback for browsers without createImageBitmap.
  }
})();