(() => {
  "use strict";

  const state = {
    paymentFile: null,
    baseFile: null,
    payments: [],
    quarters: [],
    paymentOcrText: "",
    unparsedPaymentLines: [],
    baseText: "",
    sourceRate: null,
    sourceRateLabel: "",
    results: null,
    ocrWorker: null,
    paymentConsensus: null,
    baseConsensus: null,
    basePageOcrChecks: [],
    basePositionPages: [],
    basePositionCheck: null,
    verificationTests: [],
    asOfDateManuallySet: false
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
  $("asOfDate").addEventListener("change", () => { state.asOfDateManuallySet = true; });

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
    if (value instanceof Date) return null;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const original = String(value == null ? "" : value).trim();
    // Prevent dates and quarter labels from being mistaken for amounts.
    if (/^(?:Q\s*[1-4]\s*[\/.-]\s*\d{2,4}|[1-4]\s*Q\s*[\/.-]\s*\d{2,4})$/i.test(original)) return null;
    if (/^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}$/.test(original) || /^\d{4}-\d{2}-\d{2}$/.test(original)) return null;
    let s = original.replace(/[Kk][čc]\.?/g, "").replace(/CZK/gi, "").replace(/\s+/g, "");
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

  // Mirrors the source workbook: I = H/10, J = ROUND(amount * I, 0).
  // When the payment cell is empty, the amount is the unpaid balance remaining
  // after earlier allocations to that quarter.
  function templateLateIncrease(amount, delayMonths) {
    const principal = Math.max(0, Number(amount) || 0);
    const months = Math.max(0, Math.floor(Number(delayMonths) || 0));
    return Math.round(principal * (months / 10));
  }

  function normalizeQuarterLabel(value) {
    const m = quarterMatch(value);
    return m ? m.period : String(value || "").trim();
  }

  function parsePaymentText(text, trackUnparsed = true) {
    const lines = String(text || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const parsed = [];
    const unparsed = [];
    lines.forEach(line => {
      const dates = dateMatches(line);
      if (!dates.length) {
        if (/\d/.test(line) && amountTokens(line).length >= 1 && !/částka|vyřízeno|datum|datum platby|platby/i.test(line)) {
          unparsed.push(line);
        }
        return;
      }
      const date = toISO(dates[dates.length - 1].text);
      if (!date) {
        unparsed.push(line);
        return;
      }
      const tokens = amountTokens(removeRanges(line, dates));
      if (!tokens.length) {
        unparsed.push(line);
        return;
      }
      const vals = tokens.map(t => t.value);
      const uniqueVals = vals.filter((v, i) => vals.findIndex(x => Math.abs(x - v) < 0.005) === i);
      // The source table has requested amount first and "Vyřízeno" second.
      // Prefer the final numeric token in the row; 0 means not actually processed.
      const amount = tokens[tokens.length - 1].value;
      const requestedAmount = tokens[0].value;
      if (amount == null || amount < 0 || requestedAmount == null || requestedAmount < 0) {
        unparsed.push(line);
        return;
      }
      const normalized = { date, requestedAmount: Math.round(requestedAmount * 100) / 100, amount: Math.round(amount * 100) / 100, raw: line };
      if (!parsed.some(p => p.date === normalized.date && p.amount === normalized.amount && p.raw === normalized.raw)) parsed.push(normalized);
    });
    if (trackUnparsed) state.unparsedPaymentLines = unparsed;
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
      if (!(base > 0) && !(premium > 0)) continue;
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
      if (!(base > 0) && !(premium > 0)) return;
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
        // The original financial table has small text. Upscale it to roughly 2,700 px wide
    // rather than 1,500 px so OCR retains distinctions between nearby digits.
    const factor = Math.min(3.2, Math.max(1, 2900 / Math.max(bitmap.width, bitmap.height)));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(bitmap.width * factor);
        canvas.height = Math.round(bitmap.height * factor);
        const ctx = canvas.getContext("2d", { willReadFrequently: true });
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = "high";
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
        // Locally improve grayscale contrast to make small table values easier for OCR.
        const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = pixels.data;
        for (let i = 0; i < data.length; i += 4) {
          const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
          const contrast = Math.max(0, Math.min(255, (gray - 128) * 1.65 + 128));
          data[i] = data[i + 1] = data[i + 2] = contrast;
        }
        ctx.putImageData(pixels, 0, 0);
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


  async function makeOcrVariants(blob) {
    if (!window.createImageBitmap) return [blob, blob, blob, blob, blob];
    let bitmap;
    try { bitmap = await createImageBitmap(blob); }
    catch (_) { return [blob, blob, blob, blob]; }
    const factor = Math.min(5, Math.max(1, 1500 / Math.max(bitmap.width, bitmap.height)));
    const width = Math.max(1, Math.round(bitmap.width * factor));
    const height = Math.max(1, Math.round(bitmap.height * factor));
    const base = document.createElement("canvas");
    base.width = width; base.height = height;
    const baseCtx = base.getContext("2d", { willReadFrequently: true });
    baseCtx.imageSmoothingEnabled = true;
    baseCtx.imageSmoothingQuality = "high";
    baseCtx.fillStyle = "#fff"; baseCtx.fillRect(0, 0, width, height);
    baseCtx.drawImage(bitmap, 0, 0, width, height);
    bitmap.close();

    const image = baseCtx.getImageData(0, 0, width, height);
    const pixels = image.data;
    const makeVariant = (mode) => {
      const canvas = document.createElement("canvas");
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      const copy = ctx.createImageData(width, height);
      const src = pixels, dst = copy.data;
      for (let i = 0; i < src.length; i += 4) {
        const gray = 0.299 * src[i] + 0.587 * src[i + 1] + 0.114 * src[i + 2];
        let value;
        if (mode === "contrast") value = Math.max(0, Math.min(255, (gray - 128) * 1.65 + 128));
        else if (mode === "threshold150") value = gray > 150 ? 255 : 0;
        else if (mode === "threshold185") value = gray > 185 ? 255 : 0;
        else if (mode === "threshold125") value = gray > 125 ? 255 : 0;
        else if (mode === "invert") value = 255 - gray;
        else value = gray;
        dst[i] = dst[i + 1] = dst[i + 2] = value;
        dst[i + 3] = 255;
      }
      ctx.putImageData(copy, 0, 0);
      return canvas;
    };
    // Keep one inverted version to recover white characters from selected/dark rows.
    return [base, makeVariant("contrast"), makeVariant("threshold150"), makeVariant("threshold185"), makeVariant("invert")];
  }

  function canonicalOcrSignature(kind, text) {
    if (kind === "payments") {
      const rows = parsePaymentText(text, false).map(p => [p.date, p.requestedAmount, p.amount])
        .sort((a,b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
      return JSON.stringify(rows);
    }
    const rows = parseQuartersFromText(text).map(q => [q.period, q.dueDate, q.sourcePeriod, q.base, q.premium])
      .sort((a,b) => String(a[0]).localeCompare(String(b[0])));
    return JSON.stringify(rows);
  }

  async function ocrImageConsensus(blob, label, kind) {
    const worker = await getOcrWorker();
    const variants = await makeOcrVariants(blob);
    const psmModes = ["3", "4", "6", "11"];
    const rounds = [];
    const total = variants.length * psmModes.length;
    setStatus("Ověřuji OCR", label + " · porovnávám " + total + " čtení s různým zpracováním obrazu a rozložením textu.", true);

    for (let vi = 0; vi < variants.length; vi++) {
      for (let pi = 0; pi < psmModes.length; pi++) {
        const number = rounds.length + 1;
        $("statusText").textContent = label + " · kontrolní čtení " + number + "/" + total + " (obraz " + (vi + 1) + "/" + variants.length + ", režim " + psmModes[pi] + ").";
        try {
          await worker.setParameters({ tessedit_pageseg_mode: psmModes[pi], tessedit_char_whitelist: "" });
          const result = await worker.recognize(variants[vi]);
          const text = result && result.data ? result.data.text || "" : "";
          const data = result && result.data ? result.data : null;
          const positionResult = kind === "quarters" && data ? positionRowsFromData(data) : null;
          const usePositionSignature = Boolean(positionResult && positionResult.rows.length);
          const signature = usePositionSignature
            ? JSON.stringify(positionResult.rows.map(row => [row.period, row.dueDate, row.base, row.premium]).sort((a,b)=>String(a[0]).localeCompare(String(b[0]))))
            : canonicalOcrSignature(kind, text);
          const parsedRows = kind==="payments" ? parsePaymentText(text,false) : [];
          const parsedCount = kind === "payments"
            ? parsedRows.length
            : Math.max(parseQuartersFromText(text).length, positionResult ? positionResult.rows.length : 0);
          rounds.push({ number, text, data, parsedRows, positionRows:positionResult ? positionResult.rows : [], signature, parsedCount, positionCount:positionResult ? positionResult.rows.length : 0, score: parsedCount * 10000 + (usePositionSignature ? 5000 : 0) + text.length });
        } catch (error) {
          rounds.push({ number, text: "", data:null, parsedRows:[], positionRows:[], signature: "OCR_ERROR_" + number, parsedCount: 0, score: 0, error: error && error.message ? error.message : String(error) });
        }
      }
    }
    const supplementalPaymentReadings = kind === "payments"
      ? await ocrHighlightedPaymentReadings(blob)
      : [];
    await worker.setParameters({ tessedit_pageseg_mode: "3", tessedit_char_whitelist: "" });
    const best = rounds.slice().sort((a,b) => b.score - a.score)[0] || { text: "", signature: "EMPTY", parsedCount: 0, data:null };
    const signatures = rounds.map(x => x.signature);
    const frequency = new Map();
    signatures.forEach(sig => frequency.set(sig, (frequency.get(sig) || 0) + 1));
    const maxVotes = Math.max(0, ...frequency.values());

    let consensusRows = [];
    let stableRows = 0, lowConfidenceRows = 0;
    let rowAggregation = null;
    if (kind === "payments") {
      rowAggregation = aggregatePaymentReadings(rounds, supplementalPaymentReadings);
      consensusRows = rowAggregation.rows;
      stableRows = rowAggregation.stableRows;
      lowConfidenceRows = rowAggregation.lowConfidenceRows.length;
      // Row-wise voting tolerates harmless layout/spacing differences between OCR
      // variants while requiring repeated support for each selected payment.
      const verified = rounds.length === 20 && consensusRows.length >= 12 &&
        consensusRows.every(row => row.ocrStable && row.ocrVotes >= 2);
      if (verified) state.unparsedPaymentLines = [];
    }
    const normalizedConsensus = kind === "payments"
      ? Boolean(rounds.length === 20 && rowAggregation && rowAggregation.passed)
      : rounds.length === 20;
    const result = {
      mode: "OCR – většinové hlasování po řádcích",
      total: rounds.length,
      unanimous: normalizedConsensus,
      bestText: best.text,
      parsedCount: best.parsedCount,
      positionedData: best.data || null,
      matchingRuns: maxVotes,
      consensusRows,
      stableRows,
      lowConfidenceRows,
      signatures: Array.from(frequency.entries()).map(([signature,count]) => ({ signature, count })),
      differences: normalizedConsensus ? [] : ["Některé řádky nemají dostatečnou podporu z 20 čtení nebo jsou neúplné."],
      details: rounds.map(x => ({ number:x.number, parsedCount:x.parsedCount, score:x.score, error:x.error || "" }))
    };
    return { text: best.text, data: best.data || null, consensus: result, rounds, consensusRows };
  }

  function makeTextVariants(text) {
    const source = String(text || "");
    const lines = source.split(/\r?\n/);
    const variants = [
      source,
      source.replace(/\r\n/g, "\n"),
      source.normalize("NFKC"),
      source.replace(/[\u00a0\u202f]/g, " "),
      source.replace(/[\u00a0\u202f]/g, "\u00a0"),
      source.replace(/[ \t]+/g, " "),
      source.replace(/[ \t]+/g, "\t"),
      source.replace(/\s*\|\s*/g, " | "),
      lines.map(x => x.trim()).join("\n"),
      lines.map(x => x.trim()).reverse().join("\n"),
      lines.map(x => x.replace(/[ \t]+/g, " ")).join("\n"),
      lines.map(x => x.replace(/[ \t]+/g, "\u00a0")).join("\n"),
      source.toLocaleLowerCase("cs-CZ"),
      source.replace(/Kč/gi, " Kč"),
      source.replace(/(\d) (\d{3})(?!\d)/g, "$1\u00a0$2"),
      source.replace(/(\d)\u00a0(\d{3})(?!\d)/g, "$1 $2"),
      source.replace(/\n{2,}/g, "\n"),
      source.replace(/\n/g, " \n"),
      lines.filter(Boolean).join("\n"),
      lines.map(x => x.trim()).join("\r\n")
    ];
    return variants.slice(0,20);
  }

  function textParserConsensus(text, kind) {
    const variants = makeTextVariants(text);
    const reads = variants.map((variant, i) => {
      const signature = canonicalOcrSignature(kind, variant);
      return { number:i+1, signature };
    });
    const first = reads[0] ? reads[0].signature : "EMPTY";
    const unanimous = reads.length === 20 && reads.every(x => x.signature === first);
    const frequency = new Map();
    reads.forEach(x => frequency.set(x.signature, (frequency.get(x.signature) || 0) + 1));
    const matchingRuns = Math.max(0, ...frequency.values());
    return {
      mode: "Textové varianty",
      total: reads.length,
      unanimous,
      matchingRuns,
      differences: unanimous ? [] : ["20 normalizačních průchodů nevrátilo stejné rozpoznané řádky."],
      details: reads.map(x => ({ number:x.number, signature:x.signature }))
    };
  }

  function normalizeOcrWord(value) {
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("cs-CZ").replace(/[^a-z0-9/.-]/g, "");
  }

  function positionRowsFromData(data) {
    const rawWords = data && Array.isArray(data.words) ? data.words : [];
    const words = rawWords.map(word => {
      const bbox = word.bbox || {};
      const text = String(word.text || "").trim();
      if (!text || !Number.isFinite(Number(bbox.x0)) || !Number.isFinite(Number(bbox.x1)) ||
          !Number.isFinite(Number(bbox.y0)) || !Number.isFinite(Number(bbox.y1))) return null;
      return {
        text, norm: normalizeOcrWord(text),
        x0:Number(bbox.x0), x1:Number(bbox.x1), y0:Number(bbox.y0), y1:Number(bbox.y1),
        cx:(Number(bbox.x0)+Number(bbox.x1))/2,
        cy:(Number(bbox.y0)+Number(bbox.y1))/2
      };
    }).filter(Boolean);
    if (words.length < 10) return { rows: [], reason: "OCR neposkytlo dostatek poloh slov pro rozdělení sloupců." };
    const width = Math.max(...words.map(w=>w.x1), 1);
    const height = Math.max(...words.map(w=>w.y1), 1);
    const topWords = words.filter(w=>w.cy < height * 0.46);
    const baseAnchor = topWords.find(w=>/^(zaklad|vyměř|vymerovac)/.test(w.norm));
    if (!baseAnchor) return { rows: [], reason: "V záhlaví nebyl rozpoznán sloupec „základ pro výpočet“." };

    const baseHeader = topWords.filter(w =>
      Math.abs(w.cy-baseAnchor.cy) < height*0.11 &&
      (w === baseAnchor || /^(pro|vypocet|zaklad|vyměř|vymerovac)/.test(w.norm)) &&
      Math.abs(w.cx-baseAnchor.cx) < width*0.15
    );
    const baseX = baseHeader.length ? baseHeader.reduce((sum,w)=>sum+w.cx,0)/baseHeader.length : baseAnchor.cx;

    let premiumAnchor = topWords.find(w=>/^(radne|řádné)/.test(w.norm));
    if (!premiumAnchor) {
      premiumAnchor = topWords.filter(w=>/^pojistne/.test(w.norm) && w.cx > baseX)
        .sort((a,b)=>a.cx-b.cx)[0];
    }
    if (!premiumAnchor) return { rows: [], reason: "V záhlaví nebyl rozpoznán sloupec „řádné pojistné“." };
    const premiumHeader = topWords.filter(w =>
      Math.abs(w.cy-premiumAnchor.cy) < height*0.11 &&
      (w === premiumAnchor || /^pojistne/.test(w.norm)) &&
      w.cx > baseX && w.cx < width
    );
    const premiumX = premiumHeader.length ? premiumHeader.reduce((sum,w)=>sum+w.cx,0)/premiumHeader.length : premiumAnchor.cx;
    if (!(premiumX > baseX)) return { rows: [], reason: "Nepodařilo se určit pořadí sloupců základ a pojistné." };

    const prevCandidates = topWords.filter(w =>
      w.cx < baseX && /^(zpusob|zamestnavani|obdobi|ctvrtleti|na)$/.test(w.norm)
    );
    const previousX = prevCandidates.length ? Math.max(...prevCandidates.map(w=>w.cx)) : null;
    const baseLeft = previousX == null ? 0 : (previousX+baseX)/2;
    const boundary = (baseX+premiumX)/2;
    const tolerance = Math.max(4,height*0.008);
    const sorted = words.slice().sort((a,b)=>a.cy-b.cy || a.cx-b.cx);
    const grouped = [];
    sorted.forEach(word=>{
      let row = grouped[grouped.length-1];
      if (!row || Math.abs(word.cy-row.meanY) > tolerance) {
        row = { words:[], meanY:word.cy };
        grouped.push(row);
      }
      row.words.push(word);
      row.meanY = row.words.reduce((sum,w)=>sum+w.cy,0)/row.words.length;
    });

    function numberInCell(cellWords) {
      const parts = cellWords
        .map(w=>String(w.text).trim())
        .filter(t=>/\d/.test(t))
        .map(t=>t.replace(/[^\d,.-]/g,""))
        .filter(t=>t && /^\d[\d,.-]*$/.test(t));
      if (!parts.length) return null;
      // Thousands can be OCR'd as separate words (e.g. 258 + 475).
      // Known cell boundaries keep base and premium digits separate.
      const val = parseAmount(parts.join(""));
      return val == null ? null : val;
    }

    const resultRows = grouped.map(row=>{
      const rowWords = row.words;
      const line = rowWords.map(w=>w.text).join(" ");
      const quarterLabels = Array.from(line.matchAll(/(?:\bQ\s*[1-4]\s*[\/.\-]\s*\d{2,4}\b|\b[1-4]\s*Q\s*[\/.\-]\s*\d{2,4}\b)/gi), m => quarterMatch(m[0])).filter(Boolean);
      const q = quarterLabels[0] || quarterMatch(line);
      const sourcePeriod = quarterLabels.length > 1 ? quarterLabels[1].period : "";
      const dates = dateMatches(line);
      let rowBaseLeft = baseLeft;
      if (previousX == null) {
        const leftSemanticWords = rowWords.filter(w =>
          w.cx < boundary && (quarterMatch(w.text) || dateMatches(w.text).length > 0)
        );
        if (leftSemanticWords.length) {
          rowBaseLeft = Math.max(rowBaseLeft, Math.max(...leftSemanticWords.map(w=>w.cx)) + 1);
        }
      }
      const baseWords = rowWords.filter(w=>w.cx >= rowBaseLeft && w.cx < boundary);
      const premiumWords = rowWords.filter(w=>w.cx >= boundary);
      const base = numberInCell(baseWords);
      const premium = numberInCell(premiumWords);
      return { period:q ? q.period : "", sourcePeriod, dueDate:dates.length ? toISO(dates[0].text) : "", base, premium, y:row.meanY, line };
    }).filter(row=>row.base > 0 && row.premium > 0 && row.base <= 10000000 && row.premium < 100000);
    return {
      rows:resultRows,
      columns:{baseX,premiumX,baseLeft,boundary,width,height},
      reason:resultRows.length ? "" : "V určených sloupcích nebyly nalezeny oddělené číselné hodnoty."
    };
  }

  function correctQuarterValuesFromPositions(quarters, positionedPages) {
    const pages = (positionedPages || []).map(page=>positionRowsFromData(page.data || page)).filter(p=>p.rows.length);
    const rows = pages.flatMap(p=>p.rows).sort((a,b)=>a.y-b.y);
    // Build candidate quarter rows directly from cells. Prefer position-based
    // values whenever they cover the full table: plain OCR text can merge columns.
    const generated = rows
      .filter(row => row.period && row.base > 0 && row.premium > 0)
      .map(row => ({
        period: normalizeQuarterLabel(row.period),
        dueDate: row.dueDate || "",
        sourcePeriod: row.sourcePeriod || "",
        base: row.base,
        premium: Math.round(row.premium),
        raw: "Načteno podle polohy buněk: " + row.line
      }));
    const unique = [];
    generated.forEach(row => { if (!unique.some(q => q.period === row.period)) unique.push(row); });
    const positionCoversExisting = unique.length > 0 &&
      quarters.every(q => unique.some(p => p.period === normalizeQuarterLabel(q.period)));
    if (unique.length > 0 && (
      quarters.length === 0 ||
      unique.length > quarters.length ||
      (unique.length === quarters.length && positionCoversExisting)
    )) {
      quarters.splice(0, quarters.length, ...unique);
    }
    const byPeriod = new Map();
    rows.forEach(row=>{ if(row.period) byPeriod.set(row.period,row); });
    const hasPeriodRows = quarters.length > 0 && quarters.every(q=>byPeriod.has(normalizeQuarterLabel(q.period)));
    let matched = 0;
    if (hasPeriodRows) {
      quarters.forEach(q=>{
        const row=byPeriod.get(normalizeQuarterLabel(q.period));
        if (!row) return;
        q.base=row.base; q.premium=Math.round(row.premium);
        q.raw=(q.raw ? q.raw+" | " : "")+"Sloupce podle polohy: základ "+row.base+", pojistné "+row.premium;
        matched++;
      });
    } else {
      const validRows = rows.filter(r=>r.base>0 && r.premium>0);
      if (quarters.length && validRows.length===quarters.length) {
        const byDue=quarters.slice().sort((a,b)=>String(a.dueDate||"").localeCompare(String(b.dueDate||"")));
        validRows.forEach((row,i)=>{
          byDue[i].base=row.base; byDue[i].premium=Math.round(row.premium);
          byDue[i].raw=(byDue[i].raw ? byDue[i].raw+" | " : "")+"Sloupce podle polohy: základ "+row.base+", pojistné "+row.premium;
        });
        matched=byDue.length;
      }
    }
    const ok=quarters.length>0 && matched===quarters.length && pages.length>0;
    return {
      ok, matchedCount:matched, expectedCount:quarters.length,
      mode:"Poloha buněk",
      detail:ok
        ? "Základ a řádné pojistné byly načteny samostatně podle polohy v tabulce ("+matched+"/"+quarters.length+" čtvrtletí)."
        : "Poloha sloupců potvrdila "+matched+" z "+quarters.length+" čtvrtletí. Zkontroluj, zda základ a pojistné zůstaly oddělené.",
      pageCount:pages.length
    };
  }



  function aggregatePaymentReadings(rounds, supplementalReadings = []) {
    const slots = new Map();
    let usableRounds = 0;
    const addRows = (rows, supplemental = false) => {
      const perDate = new Map();
      (rows || []).forEach(row => {
        const ordinal = perDate.get(row.date) || 0;
        perDate.set(row.date, ordinal + 1);
        const key = row.date + "#" + ordinal;
        if (!slots.has(key)) slots.set(key, { date:row.date, ordinal, seenRounds:0, normalSeenRounds:0, supplementalVotes:0, pairs:new Map(), raw:row.raw || "" });
        const slot = slots.get(key);
        slot.seenRounds++;
        if (supplemental) slot.supplementalVotes++;
        else slot.normalSeenRounds++;
        const pairKey = String(row.requestedAmount) + "|" + String(row.amount);
        if (!slot.pairs.has(pairKey)) slot.pairs.set(pairKey, { requestedAmount:row.requestedAmount, amount:row.amount, votes:0, raw:row.raw || "" });
        slot.pairs.get(pairKey).votes++;
      });
    };
    for (const round of rounds) {
      if (!round || !round.text) continue;
      usableRounds++;
      addRows(parsePaymentText(round.text, false), false);
    }
    for (const reading of supplementalReadings) addRows(reading, true);
    const minPresence = Math.max(6, Math.ceil(rounds.length * 0.30));
    const consensusRows = [];
    const confidence = [];
    for (const slot of slots.values()) {
      // Reject OCR-only ghost rows (for example 19.01.2024 misread as 19.02.2024).
      // A highlighted row can instead qualify via at least three independent crop reads.
      if (slot.normalSeenRounds < minPresence && slot.supplementalVotes < 3) continue;
      const pairs = Array.from(slot.pairs.values()).sort((a,b)=>b.votes-a.votes);
      const winner = pairs[0];
      if (!winner) continue;
      const effectiveReadCount = slot.normalSeenRounds || slot.supplementalVotes;
      const stable = winner.votes >= Math.max(2, Math.ceil(effectiveReadCount * 0.35));
      consensusRows.push({
        date:slot.date,
        ordinal:slot.ordinal,
        requestedAmount:winner.requestedAmount,
        amount:winner.amount,
        raw:winner.raw || slot.raw,
        ocrVotes:winner.votes,
        ocrSeen:slot.seenRounds,
        ocrStable:stable,
        ocrAlternativeCount:pairs.length
      });
      confidence.push(stable);
    }
    consensusRows.sort((a,b)=>a.date.localeCompare(b.date) || a.ordinal-b.ordinal);
    return {
      rows:consensusRows,
      usableRounds,
      stableRows:consensusRows.filter(row=>row.ocrStable).length,
      lowConfidenceRows:consensusRows.filter(row=>!row.ocrStable),
      passed:rounds.length===20 && usableRounds>=15 && consensusRows.length>=12 &&
        consensusRows.every(row=>row.ocrStable && row.ocrVotes>=2)
    };
  }



  async function ocrHighlightedPaymentReadings(file) {
    if (!window.createImageBitmap) return [];
    let bitmap;
    try { bitmap = await createImageBitmap(file); } catch (_) { return []; }
    const width=bitmap.width, height=bitmap.height;
    const canvas=document.createElement("canvas");
    canvas.width=width; canvas.height=height;
    const ctx=canvas.getContext("2d",{willReadFrequently:true});
    ctx.drawImage(bitmap,0,0);
    const image=ctx.getImageData(0,0,width,height), pixels=image.data;
    const x0=Math.floor(width*0.02), x1=Math.floor(width*0.98);
    const rowCounts=new Uint32Array(height);
    for(let y=0;y<height;y++){
      let count=0;
      for(let x=x0;x<x1;x++){
        const i=(y*width+x)*4;
        if((pixels[i]+pixels[i+1]+pixels[i+2])/3<110) count++;
      }
      rowCounts[y]=count;
    }
    const candidates=[];
    for(let y=Math.floor(height*0.08);y<Math.floor(height*0.92);y++){
      if(rowCounts[y]>(x1-x0)*0.45) candidates.push(y);
    }
    const bands=[];
    candidates.forEach(y=>{
      const last=bands[bands.length-1];
      if(last && y-last[last.length-1]<=1) last.push(y);
      else bands.push([y]);
    });
    const usableBands=bands.filter(band=>{
      const h=band[band.length-1]-band[0]+1;
      return h>=height*0.012 && h<=height*0.09;
    });
    if(!usableBands.length){bitmap.close();return [];}
    const worker=await getOcrWorker();
    const readings=[];
    for(const band of usableBands){
      const top=Math.max(0,band[0]-2), bottom=Math.min(height,band[band.length-1]+3);
      const crop=document.createElement("canvas");
      const scale=5;
      crop.width=(x1-x0)*scale; crop.height=(bottom-top)*scale;
      const cctx=crop.getContext("2d",{willReadFrequently:true});
      cctx.drawImage(bitmap,x0,top,x1-x0,bottom-top,0,0,crop.width,crop.height);
      const data=cctx.getImageData(0,0,crop.width,crop.height);
      const px=data.data;
      for(let i=0;i<px.length;i+=4){
        const gray=0.299*px[i]+0.587*px[i+1]+0.114*px[i+2];
        const v=Math.max(0,Math.min(255,(255-gray-128)*1.6+128));
        px[i]=px[i+1]=px[i+2]=v;
      }
      cctx.putImageData(data,0,0);
      for(const psm of ["7","6","4","11"]){
        try{
          await worker.setParameters({tessedit_pageseg_mode:psm,tessedit_char_whitelist:""});
          const result=await worker.recognize(crop);
          const text=result&&result.data?String(result.data.text||""):"";
          const rows=parsePaymentText(text,false);
          if(rows.length) readings.push(rows);
        }catch(_){}
      }
    }
    bitmap.close();
    await worker.setParameters({tessedit_pageseg_mode:"3",tessedit_char_whitelist:""});
    return readings;
  }

  function quarterFromDueDate(value) {
    const date = parseDate(value);
    if (!date) return null;
    const month = date.getMonth() + 1;
    const quarterByMonth = { 1:1, 4:2, 7:3, 10:4 };
    const quarter = quarterByMonth[month];
    if (!quarter) return null;
    const lastDay = new Date(date.getFullYear(), month, 0).getDate();
    if (date.getDate() !== lastDay) return null;
    return { period:"Q"+quarter+"/"+date.getFullYear(), quarter, year:date.getFullYear() };
  }

  function previousQuarterPeriod(info) {
    if (!info) return "";
    const q = info.quarter === 1 ? 4 : info.quarter - 1;
    const year = info.quarter === 1 ? info.year - 1 : info.year;
    return "Q"+q+"/"+year;
  }

  function medianOf(values) {
    const nums = values.filter(v=>Number.isFinite(v)).slice().sort((a,b)=>a-b);
    if (!nums.length) return null;
    const mid = Math.floor(nums.length/2);
    return nums.length%2 ? nums[mid] : (nums[mid-1]+nums[mid])/2;
  }

  function repairPremiumsFromCommonRate(quarters, forcedRate) {
    const ratios = quarters
      .filter(q=>Number(q.base)>0 && Number(q.premium)>0)
      .map(q=>Number(q.premium)/Number(q.base)*1000);
    const rate = Number(forcedRate)>0 ? Number(forcedRate) :
      (ratios.length>=5 ? Math.round(medianOf(ratios)*10)/10 : null);
    const corrections = [];
    if (!(rate>0) || !Number.isFinite(rate)) return { rate:null, corrections };
    for (const q of quarters) {
      if (!(Number(q.base)>0) || !(Number(q.premium)>0)) continue;
      const expected = Math.round(Number(q.base)*rate/1000);
      const ratio = Number(q.premium)/Number(q.base)*1000;
      // Only auto-correct a likely OCR digit error when multiple other quarters
      // independently support the same tenth-of-a-per-mille rate.
      const ratioIsOutlier = ratios.length>=5 && Math.abs(ratio-rate)>0.015;
      if (Math.abs(Number(q.premium)-expected)>1 && ratioIsOutlier) {
        const previous = Number(q.premium);
        q.premium = expected;
        q.raw = (q.raw ? q.raw+" | " : "")+
          "Kontrola sazby: OCR četlo "+previous+" Kč; při sazbě "+rate.toLocaleString("cs-CZ")+" ‰ vychází "+expected+" Kč.";
        q.ocrCorrection = { from:previous, to:expected, rate };
        corrections.push({ period:q.period, from:previous, to:expected, rate });
      }
    }
    return { rate, corrections };
  }

  async function extractQuarterRowsFromGridImage(file) {
    if (!window.createImageBitmap) return { quarters:[], ok:false, detail:"Prohlížeč neumí načíst souřadnice obrázku." };
    let bitmap;
    try { bitmap = await createImageBitmap(file); }
    catch (_) { return { quarters:[], ok:false, detail:"Obrázek se nepodařilo otevřít pro čtení buněk." }; }
    const canvas=document.createElement("canvas");
    canvas.width=bitmap.width; canvas.height=bitmap.height;
    const ctx=canvas.getContext("2d",{willReadFrequently:true});
    ctx.drawImage(bitmap,0,0);
    const image=ctx.getImageData(0,0,canvas.width,canvas.height);
    const px=image.data, width=canvas.width, height=canvas.height;
    const borderPixel=(x,y)=>{
      const idx=(y*width+x)*4;
      const r=px[idx],g=px[idx+1],b=px[idx+2];
      const lum=(r+g+b)/3;
      return Math.max(r,g,b)-Math.min(r,g,b)<18 && lum>=120 && lum<=195;
    };
    const groupLines=(candidates)=>{
      const groups=[];
      candidates.sort((a,b)=>a.pos-b.pos).forEach(item=>{
        const last=groups[groups.length-1];
        if(last && item.pos-last[last.length-1].pos<=2) last.push(item);
        else groups.push([item]);
      });
      return groups.map(group=>group.sort((a,b)=>b.score-a.score)[0]);
    };
    const xCandidates=[];
    const yScanStart=Math.floor(height*0.05), yScanEnd=Math.floor(height*0.97);
    for(let x=Math.floor(width*0.02);x<Math.floor(width*0.98);x++){
      let count=0;
      for(let y=yScanStart;y<yScanEnd;y++) if(borderPixel(x,y)) count++;
      if(count>(yScanEnd-yScanStart)*0.55) xCandidates.push({pos:x,score:count});
    }
    const xLines=groupLines(xCandidates);
    if(xLines.length<6){ bitmap.close(); return {quarters:[],ok:false,detail:"Nepodařilo se najít svislé hranice sloupců tabulky."}; }
    const minX=xLines[0].pos, maxX=xLines[xLines.length-1].pos;
    const yCandidates=[];
    const lineWidth=maxX-minX;
    for(let y=Math.floor(height*0.05);y<Math.floor(height*0.97);y++){
      let count=0;
      for(let x=minX;x<=maxX;x++) if(borderPixel(x,y)) count++;
      if(count>lineWidth*0.62) yCandidates.push({pos:y,score:count});
    }
    const yLines=groupLines(yCandidates);
    // A real table has many evenly spaced horizontal cell borders. Pick the
    // longest contiguous sequence with realistic row heights, not stray window rules.
    let bestSequence=[];
    for(let start=0;start<yLines.length;start++){
      const sequence=[yLines[start]];
      for(let j=start+1;j<yLines.length;j++){
        const gap=yLines[j].pos-sequence[sequence.length-1].pos;
        if(gap<height*0.02 || gap>height*0.12) break;
        sequence.push(yLines[j]);
      }
      if(sequence.length>bestSequence.length) bestSequence=sequence;
    }
    const rowLines=bestSequence.map(x=>x.pos);
    if(rowLines.length<10 || xLines.length<6){
      bitmap.close();
      return {quarters:[],ok:false,detail:"Tabulkové vodorovné hranice nebyly nalezeny dostatečně spolehlivě."};
    }
    const columns=xLines.map(x=>x.pos);
    const headerTop=rowLines[0], headerBottom=rowLines[1];
    if(headerBottom-headerTop<height*0.025 || headerBottom-headerTop>height*0.12){
      bitmap.close();
      return {quarters:[],ok:false,detail:"Nepodařilo se určit řádek záhlaví."};
    }
    const worker=await getOcrWorker();
    const cropCanvas=document.createElement("canvas");
    const cropCtx=cropCanvas.getContext("2d",{willReadFrequently:true});
    const cropText=async(box, mode)=>{
      const sx=Math.max(0,Math.round(box.x0)), sy=Math.max(0,Math.round(box.y0));
      const sw=Math.max(1,Math.round(box.x1)-sx), sh=Math.max(1,Math.round(box.y1)-sy);
      const scale=Math.max(3,Math.min(5,Math.ceil(900/Math.max(sw,sh))));
      cropCanvas.width=sw*scale; cropCanvas.height=sh*scale;
      cropCtx.fillStyle="#fff"; cropCtx.fillRect(0,0,cropCanvas.width,cropCanvas.height);
      cropCtx.imageSmoothingEnabled=true; cropCtx.imageSmoothingQuality="high";
      cropCtx.drawImage(bitmap,sx,sy,sw,sh,0,0,cropCanvas.width,cropCanvas.height);
      // Convert tiny screenshot cells to high-contrast grayscale, still locally.
      const cellImage=cropCtx.getImageData(0,0,cropCanvas.width,cropCanvas.height);
      const cellPixels=cellImage.data;
      for(let p=0;p<cellPixels.length;p+=4){
        const gray=0.299*cellPixels[p]+0.587*cellPixels[p+1]+0.114*cellPixels[p+2];
        const contrast=Math.max(0,Math.min(255,(gray-128)*1.4+128));
        cellPixels[p]=cellPixels[p+1]=cellPixels[p+2]=contrast;
      }
      cropCtx.putImageData(cellImage,0,0);
      const whitelist=mode==="date" ? "0123456789./-" :
        mode==="number" ? "0123456789., KčKc|/" : "";
      const attempts=mode==="text" ? ["7"] : ["7","6"];
      const values=[];
      for(const psm of attempts){
        await worker.setParameters({tessedit_pageseg_mode:psm,tessedit_char_whitelist:whitelist});
        const result=await worker.recognize(cropCanvas);
        const text=result && result.data ? String(result.data.text||"").trim() : "";
        const value=mode==="date"
          ? (text.match(/\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/)||[])[0] || ""
          : mode==="number" ? parseAmount(text) : text;
        values.push({text,value,confidence:result && result.data ? Number(result.data.confidence)||0 : 0});
        if(mode==="date" ? Boolean(value) : mode==="number" ? Number.isFinite(value)&&value>0 : text.length>1) break;
      }
      const chosen=values.slice().sort((a,b)=>{
        const av=mode==="date" ? Boolean(a.value) : mode==="number" ? Number.isFinite(a.value)&&a.value>0 : a.text.length>1;
        const bv=mode==="date" ? Boolean(b.value) : mode==="number" ? Number.isFinite(b.value)&&b.value>0 : b.text.length>1;
        if(av!==bv) return av?-1:1;
        return b.confidence-a.confidence;
      })[0] || {text:"",value:null,confidence:0};
      return chosen;
    };
    try {
      // This standardized payroll/premium table has six data columns; the two
      // rightmost columns are the calculation base and regular premium.
      if(columns.length<7) {
        bitmap.close();
        return {quarters:[],ok:false,detail:"Tabulka nemá očekávaných šest datových sloupců pro základ a pojistné."};
      }
      // columns[] stores grid boundaries; with six table columns, the last two
      // cell indices are length-3 (base) and length-2 (premium).
      const baseColumn=columns.length-3, premiumColumn=columns.length-2;
      const headerBase=await cropText({x0:columns[baseColumn]+3,x1:columns[baseColumn+1]-3,y0:headerTop+2,y1:headerBottom-2},"text");
      const headerPremium=await cropText({x0:columns[premiumColumn]+3,x1:columns[premiumColumn+1]-3,y0:headerTop+2,y1:headerBottom-2},"text");
      const baseHeaderNorm=normalizeOcrWord(headerBase.text);
      const premiumHeaderNorm=normalizeOcrWord(headerPremium.text);
      let headerConfirmed=(baseHeaderNorm.includes("zaklad") || baseHeaderNorm.includes("vyměř") || baseHeaderNorm.includes("vymer")) &&
        premiumHeaderNorm.includes("pojistne");
      const quarters=[];
      const rowCount=rowLines.length-1;
      for(let i=1;i<rowCount;i++){
        const y0=rowLines[i], y1=rowLines[i+1];
        if(y1-y0<height*0.02) continue;

        // Always read the period label in its own cell. A missing/poorly read
        // due-date cell must not cause the entire quarter to be skipped.
        const periodCell=await cropText({x0:columns[0]+3,x1:columns[1]-3,y0:y0+2,y1:y1-2},"text");
        const periodMatch=quarterMatch(periodCell.text);
        const dueCell=await cropText({x0:columns[1]+3,x1:columns[2]-3,y0:y0+2,y1:y1-2},"date");
        let dueDate=toISO(dueCell.value);
        let info=periodMatch || quarterFromDueDate(dueDate);
        if(!info) continue;

        // The source uses the last day of January/April/July/October for Q1/Q2/Q3/Q4.
        // Use the recognized quarter label to recover an empty or contradictory OCR date.
        const dueMonth = info.quarter * 3 - 2;
        const expectedDue = toISO(new Date(info.year, dueMonth, 0));
        const dueInfo = quarterFromDueDate(dueDate);
        const correctedDueDate = (!dueDate || !dueInfo || dueInfo.period !== info.period) ? expectedDue : dueDate;
        const baseCell=await cropText({x0:columns[baseColumn]+3,x1:columns[baseColumn+1]-3,y0:y0+2,y1:y1-2},"number");
        const premiumCell=await cropText({x0:columns[premiumColumn]+3,x1:columns[premiumColumn+1]-3,y0:y0+2,y1:y1-2},"number");
        const base=Number(baseCell.value), premium=Math.round(Number(premiumCell.value));
        if(!(base>1000&&base<=10000000&&premium>=1&&premium<100000)) continue;
        quarters.push({
          period:info.period,dueDate:correctedDueDate,sourcePeriod:previousQuarterPeriod(info),
          base,premium,
          raw:"Čtení samostatných buněk; období: "+periodCell.text+"; splatnost: "+(dueCell.text||"odvozena z období")+"; základ: "+baseCell.text+"; pojistné: "+premiumCell.text,
          cellConfidence:Math.round((baseCell.confidence+premiumCell.confidence+dueCell.confidence)/3),
          baseCellText:baseCell.text,premiumCellText:premiumCell.text,
          dueDateDerived:correctedDueDate!==dueDate
        });
      }
      // Keep one entry per period; prefer the row with complete independently-read cells.
      const byPeriod=new Map();
      quarters.forEach(q=>{
        const existing=byPeriod.get(q.period);
        if(!existing || (q.base>0&&q.premium>0&&q.cellConfidence>existing.cellConfidence)) byPeriod.set(q.period,q);
      });
      const uniqueQuarters=Array.from(byPeriod.values()).sort((a,b)=>a.dueDate.localeCompare(b.dueDate));
      const rowsComplete=uniqueQuarters.length>=10 &&
        uniqueQuarters.every((q,i)=>i===0 || validationPeriodOrdinal(q.period)===validationPeriodOrdinal(uniqueQuarters[i-1].period)+1);
      bitmap.close();
      await worker.setParameters({tessedit_pageseg_mode:"3",tessedit_char_whitelist:""});
      return {
        quarters:uniqueQuarters, ok:rowsComplete && (headerConfirmed || uniqueQuarters.length>=12),
        headerConfirmed,
        detail:rowsComplete
          ? (headerConfirmed ? "Základ, pojistné i období byly přečteny po samostatných buňkách; prázdné nebo rozporné datumy byly odvozeny z označení čtvrtletí." : "Čtvrtletí byla přečtena po samostatných buňkách; datumy byly doplněny podle označení období.")
          : "Podařilo se bezpečně přečíst "+uniqueQuarters.length+" navazujících čtvrtletí; ověř, zda jsou čitelné buňky období, základu a pojistného.",
        grid:{columns,rowLines}
      };
    } catch(error) {
      bitmap.close();
      await worker.setParameters({tessedit_pageseg_mode:"3",tessedit_char_whitelist:""});
      return {quarters:[],ok:false,detail:"Čtení samostatných buněk selhalo: "+(error&&error.message?error.message:String(error))};
    }
  }

  async function parsePdf(file) {
    if (!window.pdfjsLib) throw new Error("Knihovna PDF není dostupná. Zkontroluj připojení k internetu.");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
    const allLines = [];
    state.basePositionPages = [];
    for (let pageNo = 1; pageNo <= pdf.numPages; pageNo++) {
      setStatus("Čtu PDF", "Zpracovávám stránku " + pageNo + " z " + pdf.numPages + " lokálně.", true);
      const page = await pdf.getPage(pageNo);
      const content = await page.getTextContent();
      const viewport1 = page.getViewport({ scale: 1 });
      const grouped = new Map();
      const pageWords = [];
      (content.items || []).forEach(item => {
        const txt = String(item.str || "").trim();
        if (!txt) return;
        const transform = item.transform || [];
        const x = Number(transform[4]) || 0;
        const baselineY = viewport1.height - (Number(transform[5]) || 0);
        const wordHeight = Math.max(5, Number(item.height) || 8);
        const totalWidth = Math.max(1, Number(item.width) || txt.length * wordHeight * 0.45);
        const tokens = txt.split(/\s+/).filter(Boolean);
        const totalChars = Math.max(1, tokens.reduce((sum,t)=>sum+t.length,0));
        let charOffset = 0;
        tokens.forEach(token => {
          const x0 = x + totalWidth * charOffset / totalChars;
          charOffset += token.length;
          const x1 = x + totalWidth * charOffset / totalChars;
          pageWords.push({ text:token, bbox:{x0,x1,y0:baselineY-wordHeight,y1:baselineY} });
        });
        const y = Math.round(transform[5] || 0);
        if (!grouped.has(y)) grouped.set(y, []);
        grouped.get(y).push({ x, txt });
      });
      const lines = Array.from(grouped.entries()).sort((a,b) => b[0] - a[0]).map(entry => entry[1].sort((a,b) => a.x - b.x).map(x => x.txt).join(" "));
      const pageText = lines.join("\n").trim();
      if (pageText.length > 30) {
        allLines.push(pageText);
        state.basePositionPages.push({ data:{ words:pageWords } });
      } else {
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
        const pageOcr = await ocrImageConsensus(canvas, "PDF · stránka " + pageNo + " z " + pdf.numPages, "quarters");
        state.basePageOcrChecks.push(pageOcr.consensus);
        allLines.push(pageOcr.text);
        if (pageOcr.data) state.basePositionPages.push({ data:pageOcr.data });
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
    $("verificationPanel").hidden = true;
    state.payments = []; state.quarters = []; state.results = null;
    state.unparsedPaymentLines = [];
    state.paymentConsensus = null; state.baseConsensus = null; state.basePageOcrChecks = [];
    state.basePositionPages = []; state.basePositionCheck = null;
    try {
      const paymentScan = await ocrImageConsensus(state.paymentFile, "Screenshot plateb", "payments");
      state.paymentConsensus = paymentScan.consensus;
      state.paymentOcrText = paymentScan.text;
      state.payments = paymentScan.consensusRows && paymentScan.consensusRows.length
        ? paymentScan.consensusRows
        : parsePaymentText(state.paymentOcrText);

      // Default the calculation date to the latest date shown in the imported
      // statement (including a row with 0 Kč processed). This matches the
      // reference workbook's snapshot date; users can still override it manually.
      if (!state.asOfDateManuallySet) {
        const statementRows = parsePaymentText(state.paymentOcrText, false);
        const statementDates = [...state.payments, ...statementRows]
          .map(payment => payment.date)
          .filter(validISODate)
          .sort((a,b) => a.localeCompare(b));
        if (statementDates.length) $("asOfDate").value = statementDates[statementDates.length - 1];
      }
      $("paymentOcrText").textContent = state.paymentOcrText || "Ze screenshotu nebyl rozpoznán text.";

      setStatus("Čtu základní dokument", "Načítám data v prohlížeči a ověřuji jejich soulad.", true);
      const ext = state.baseFile.name.split(".").pop().toLowerCase();
      const isImage = ["png","jpg","jpeg","webp","gif"].includes(ext) || state.baseFile.type.startsWith("image/");
      let gridResult = null;
      let cellRateCheck = null;
      if (isImage) {
        const baseScan = await ocrImageConsensus(state.baseFile, "Základní tabulka", "quarters");
        state.baseText = baseScan.text;
        const textQuarters = parseQuartersFromText(state.baseText);
        const gridRead = await extractQuarterRowsFromGridImage(state.baseFile);
        const positionRows = baseScan.rounds.flatMap(round=>round.positionRows || []);
        state.basePositionPages = baseScan.rounds.filter(round=>round.data).map(round=>({data:round.data}));
        // If the cell reader produced a coherent sequence, prefer it over the
        // free-text parser even when the latter invents extra/misaligned rows.
        state.quarters = gridRead.ok && gridRead.quarters.length >= 10
          ? gridRead.quarters
          : (gridRead.quarters.length >= 8
            ? gridRead.quarters
            : textQuarters);
        if (gridRead.quarters.length >= 10 && gridRead.ok) {
          state.basePositionCheck = {
            ok:true, matchedCount:state.quarters.length, expectedCount:state.quarters.length,
            mode:"Čtení samostatných buněk",
            detail:gridRead.detail + " Načteno "+state.quarters.length+" čtvrtletí.",
            pageCount:1,
            grid:gridRead.grid
          };
          cellRateCheck = repairPremiumsFromCommonRate(state.quarters, parseAmount($("rateInput").value));
          if (cellRateCheck.corrections.length) {
            state.basePositionCheck.detail += " Opraveno OCR podle společné sazby "+cellRateCheck.rate.toLocaleString("cs-CZ")+" ‰: "+
              cellRateCheck.corrections.map(c=>c.period+" "+c.from+" → "+c.to+" Kč").join(", ")+".";
          }
          state.baseConsensus = {
            ...baseScan.consensus,
            mode:"20 OCR průchodů + samostatné buňky",
            unanimous:baseScan.rounds.length===20,
            total:baseScan.rounds.length,
            matchingRuns:baseScan.rounds.length,
            details:baseScan.consensus.details,
            note:"Rozhodující hodnoty byly přečteny po jednotlivých buňkách; textové rozdíly mezi OCR variantami nejsou považovány za chybu."
          };
        } else {
          state.basePositionCheck = correctQuarterValuesFromPositions(state.quarters, state.basePositionPages);
          state.baseConsensus = {
            ...baseScan.consensus,
            unanimous:false,
            differences:(baseScan.consensus.differences||[]).concat("Samostatné buňky nešlo plně ověřit: "+gridRead.detail)
          };
        }
      } else {
        gridResult = await readGridFile(state.baseFile);
        state.baseText = await readTextFile(state.baseFile);
        state.quarters = gridResult && gridResult.quarters.length ? gridResult.quarters : parseQuartersFromText(state.baseText);
        if (gridResult && gridResult.quarters.length) {
          const plausibleRows = state.quarters.length > 0 && state.quarters.every(q =>
            Number(q.base) > 1000 && Number(q.base) <= 10000000 &&
            Number(q.premium) >= 1 && Number(q.premium) <= 100000
          );
          state.basePositionCheck = {
            ok: plausibleRows,
            matchedCount: plausibleRows ? state.quarters.length : 0,
            expectedCount: state.quarters.length,
            mode: "Hodnoty buněk",
            detail: plausibleRows
              ? "Základ a řádné pojistné byly načteny z oddělených buněk tabulky ("+state.quarters.length+"/"+state.quarters.length+" čtvrtletí)."
              : "Alespoň jedna číselná buňka není v očekávaném rozsahu. Zkontroluj sloupce základu a řádného pojistného."
          };
          state.baseConsensus = {
            mode: "Strukturovaná data",
            total: 20,
            unanimous: true,
            matchingRuns: 20,
            differences: [],
            details: [],
            note: "Buňky souboru byly načteny přímo, bez OCR. Následuje kontrola struktury a výpočtových vazeb."
          };
        } else {
          state.baseConsensus = textParserConsensus(state.baseText, "quarters");
          if (state.basePositionPages.length) {
            state.basePositionCheck = correctQuarterValuesFromPositions(state.quarters, state.basePositionPages);
          } else {
            const plausibleRows = state.quarters.length > 0 && state.quarters.every(q =>
              Number(q.base) > 1000 && Number(q.base) <= 10000000 &&
              Number(q.premium) >= 1 && Number(q.premium) <= 100000
            );
            state.basePositionCheck = {
              ok: plausibleRows,
              matchedCount: plausibleRows ? state.quarters.length : 0,
              expectedCount: state.quarters.length,
              mode: "Číselná validace",
              detail: plausibleRows
                ? "Souřadnice buněk nebyly dostupné; oddělené hodnoty prošly kontrolou rozsahu a návaznosti sazby."
                : "Bez souřadnic buněk nelze bezpečně oddělit základ a pojistné; nejméně jedna hodnota je neúplná nebo mimo očekávaný rozsah."
            };
          }
        }
        if (state.basePageOcrChecks.some(check => !check.unanimous)) {
          state.baseConsensus = {
            ...state.baseConsensus,
            unanimous: false,
            differences: (state.baseConsensus.differences || []).concat("Nejméně jedna naskenovaná stránka PDF neměla shodný výsledek při 20 OCR čteních.")
          };
        }
      }

      state.sourceRate = gridResult && gridResult.rate != null ? gridResult.rate : extractRate(state.baseText);
      const detectedTextRate = extractRate(state.baseText);
      if (state.sourceRate == null && detectedTextRate != null) state.sourceRate = detectedTextRate;
      if (state.sourceRate == null && cellRateCheck && cellRateCheck.rate != null) {
        state.sourceRate = cellRateCheck.rate;
        state.sourceRateLabel = "Sazba odhadnuta z konzistentních předpisů";
      } else {
        state.sourceRateLabel = state.sourceRate != null ? "Sazba rozpoznána ze základního dokumentu" : "Sazba nebyla rozpoznána";
      }
      if (state.sourceRate != null) {
        $("rateInput").value = String(state.sourceRate);
        $("rateSource").textContent = state.sourceRateLabel;
      } else {
        $("rateInput").value = "";
        $("rateSource").textContent = "Sazba nebyla rozpoznána; doplň ji jen tehdy, pokud není uvedeno řádné pojistné.";
      }
      renderInputTables();
      $("sourceText").textContent = (state.baseText || "Z dokumentu nebyl získán žádný text.").slice(0, 35000);
      $("reviewPanel").hidden = false;
      $("reviewPanel").scrollIntoView({ behavior: "smooth", block: "start" });

      refreshVerification();
      const notes = [];
      if (!state.payments.length) notes.push("Nepodařilo se rozpoznat žádné platby.");
      if (!state.paymentConsensus.unanimous) notes.push("Po řádcích se nepodařilo stabilně konsolidovat platební screenshot.");
      if (state.unparsedPaymentLines.length) notes.push(state.unparsedPaymentLines.length + " řádků screenshotu zůstalo nezařazených.");
      if (!state.quarters.length) notes.push("Nepodařilo se rozpoznat čtvrtletní předpisy.");
      if (!state.baseConsensus.unanimous) notes.push("Čtení buněk základního dokumentu neprošlo kontrolou úplnosti.");
      setStatus(notes.length ? "Podklady načteny – výpočet zatím blokován" : "Podklady načteny", notes.length ? notes.join(" ") + " Ověř podklady; Excel se do úspěšné kontroly nevytvoří." : "Načtená data prošla úvodní kontrolou. Dokonči 20bodovou validaci.", false);
      $("spinner").hidden = false;
      $("readButton").disabled = false;
    } catch (error) {
      console.error(error);
      setStatus("Zpracování se nepodařilo", error && error.message ? error.message : "Soubor se nepodařilo načíst.", false);
      $("spinner").hidden = true;
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
      addInputCell(tr, "number", payment.requestedAmount == null ? payment.amount : payment.requestedAmount, "Částka platby " + (i + 1), value => { payment.requestedAmount = parseAmount(value); invalidateResults(); }, "0.01");
      addInputCell(tr, "number", payment.amount, "Vyřízená částka " + (i + 1), value => { payment.amount = parseAmount(value); invalidateResults(); }, "0.01");
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

  function validationPeriodOrdinal(value) {
    const m = String(value || "").match(/^Q([1-4])\/(\d{4})$/i);
    return m ? Number(m[2]) * 4 + Number(m[1]) - 1 : null;
  }

  function buildVerificationTests() {
    const payments = state.payments || [];
    const quarters = (state.quarters || []).filter(q => q.period || q.dueDate || q.premium != null);
    const rate = parseAmount($("rateInput").value);
    const add = (name, pass, detail) => tests.push({ name, pass: Boolean(pass), detail: detail || (pass ? "V pořádku" : "Vyžaduje opravu") });
    const tests = [];
    const paymentConsensus = state.paymentConsensus;
    const baseConsensus = state.baseConsensus;

    add("Platební screenshot: 20 čtení konsolidováno",
      Boolean(paymentConsensus && paymentConsensus.unanimous && paymentConsensus.total === 20 && (paymentConsensus.consensusRows||[]).length >= 12),
      !paymentConsensus ? "Kontrola OCR ještě neproběhla." :
        paymentConsensus.unanimous ? "20 průchodů bylo vyhodnoceno hlasováním po jednotlivých platbách; zahrnuté řádky mají opakovanou podporu." :
        "Po řádcích se nepodařilo stabilně sloučit alespoň 12 plateb. Zkontroluj tabulku a případné nejasné řádky.");
    add("Základní dokument: hodnoty ověřeny po buňkách",
      Boolean(baseConsensus && baseConsensus.unanimous && baseConsensus.total === 20 && state.basePositionCheck && state.basePositionCheck.ok),
      !baseConsensus ? "Kontrola základního dokumentu ještě neproběhla." :
        baseConsensus.unanimous ? (baseConsensus.mode === "Strukturovaná data" ? "Buňky byly načteny přímo; kontrolují se i matematické a časové vazby." : "Bylo dokončeno 20 OCR průchodů a rozhodující hodnoty byly extrahovány ze samostatných buněk tabulky.") :
        "Čtení buněk zatím nepotvrdilo celou tabulku; zobrazí se upozornění.");
    add("Platby byly rozpoznány", payments.length > 0,
      payments.length ? "Rozpoznáno řádků: " + payments.length + "." : "Nebyly rozpoznány žádné platební řádky.");
    add("Čtvrtletní předpisy byly rozpoznány", quarters.length > 0,
      quarters.length ? "Rozpoznáno čtvrtletí: " + quarters.length + "." : "Nebyl nalezen žádný předpis pojistného.");
    add("Všechny datumy plateb jsou platné", payments.length > 0 && payments.every(p => validISODate(p.date)),
      payments.every(p => validISODate(p.date)) ? "Každá platba má platné datum." : "Alespoň u jedné platby chybí nebo je neplatné datum.");
    add("Částky plateb jsou číselné a nezáporné", payments.length > 0 && payments.every(p => Number.isFinite(Number(p.amount)) && Number(p.amount) >= 0 && (p.requestedAmount == null || (Number.isFinite(Number(p.requestedAmount)) && Number(p.requestedAmount) >= 0))),
      "Kontrola požadované a skutečně vyřízené částky.");
    const hasPaidAmount = payments.some(p => Number(p.amount) > 0);
    const processedWithinRequest = payments.every(p => p.requestedAmount == null || Number(p.amount) <= Number(p.requestedAmount));
    add("Existuje skutečně vyřízená platba a nepřevyšuje požadovanou", hasPaidAmount && processedWithinRequest,
      !hasPaidAmount ? "Všechny vyřízené částky jsou 0 Kč nebo prázdné." :
        processedWithinRequest ? "Nulově vyřízené řádky se nezapočítají; vyřízené částky nepřevyšují částky plateb." :
        "Některá vyřízená částka převyšuje požadovanou částku.");
    const payKeys = payments.map(p => [p.date, p.requestedAmount, p.amount].join("|"));
    add("Nejsou duplicitní platební řádky", new Set(payKeys).size === payKeys.length,
      new Set(payKeys).size === payKeys.length ? "Žádná stejná kombinace data a částek." : "Nalezeny zcela shodné platební řádky.");
    add("Označení čtvrtletí mají správný formát", quarters.length > 0 && quarters.every(q => quarterMatch(q.period) && /^Q[1-4]\/\d{4}$/i.test(normalizeQuarterLabel(q.period))),
      "Očekávaný tvar je například Q4/2023.");
    const periods = quarters.map(q => normalizeQuarterLabel(q.period));
    add("Každé čtvrtletí je uvedeno pouze jednou", periods.length > 0 && new Set(periods).size === periods.length,
      new Set(periods).size === periods.length ? "Čtvrtletí se neopakují." : "Stejné čtvrtletí je v tabulce vícekrát.");
    const dueDatesValid = quarters.length > 0 && quarters.every(q => validISODate(q.dueDate));
    const orderedByDue = quarters.slice().sort((a,b) => String(a.dueDate || "").localeCompare(String(b.dueDate || "")));
    const datesChronological = orderedByDue.length > 0 && orderedByDue.every((q,i) => i === 0 || orderedByDue[i-1].dueDate <= q.dueDate);
    add("Datumy splatnosti jsou platné a chronologické", dueDatesValid && datesChronological,
      !dueDatesValid ? "Alespoň jedno datum splatnosti chybí nebo není platné." : "Splatnosti jsou platné a chronologicky navazují.");
    const dueMatchesPeriod = quarters.every(q => {
      const info = quarterMatch(q.period);
      const due = parseDate(q.dueDate);
      if (!info || !due) return false;
      const month = info.quarter * 3 - 2;
      const lastDay = new Date(info.year, month, 0).getDate();
      return due.getFullYear() === info.year && due.getMonth() + 1 === month && due.getDate() === lastDay;
    });
    add("Splatnost odpovídá označení čtvrtletí", quarters.length > 0 && dueMatchesPeriod,
      dueMatchesPeriod ? "Měsíc a poslední den splatnosti odpovídají označení Q1–Q4." : "Alespoň jedno datum nesedí k označení čtvrtletí. Zkontroluj období ve zdroji.");
    const ordinals = orderedByDue.map(q => validationPeriodOrdinal(q.period));
    const sequential = ordinals.length > 0 && ordinals.every((n,i) => n != null && (i === 0 || n === ordinals[i-1] + 1));
    add("Mezi čtvrtletími nejsou mezery", sequential,
      sequential ? "Předpisy navazují po jednotlivých čtvrtletích." : "Řada čtvrtletí má mezeru, nesprávné pořadí nebo označení. Ověř rozsah podkladu.");
    add("Všechny výpočtové základy jsou kladné", quarters.length > 0 && quarters.every(q => Number(q.base) > 0 && Number.isFinite(Number(q.base))),
      "Základ nesmí být prázdný, nulový ani textově spojený s vedlejší částkou.");
    add("Výpočtové základy mají věrohodný rozsah", quarters.length > 0 && quarters.every(q => Number(q.base) > 1000 && Number(q.base) <= 10000000),
      "Kontrola zachytí například omylem spojené hodnoty 258475 a 2171 → 258475217.");
    add("Řádné pojistné je kladné a věrohodné", quarters.length > 0 && quarters.every(q => Number(q.premium) >= 1 && Number(q.premium) <= 100000),
      "Předpis nesmí být prázdný ani podezřele malý či velký.");
    const ratioValues = quarters.filter(q => Number(q.base) > 0 && Number(q.premium) > 0).map(q => Number(q.premium) / Number(q.base) * 1000);
    let premiumMathPass = ratioValues.length === quarters.length && ratioValues.length > 0;
    if (premiumMathPass && rate > 0) {
      premiumMathPass = quarters.every(q => Math.abs(Math.round(Number(q.base) * rate / 1000) - Number(q.premium)) <= 1);
    } else if (premiumMathPass) {
      const average = ratioValues.reduce((sum,n) => sum+n, 0) / ratioValues.length;
      premiumMathPass = ratioValues.every(n => Math.abs(n-average) <= 0.35);
    }
    add("Základ a řádné pojistné odpovídají sazbě", premiumMathPass,
      rate > 0 ? "Ověřeno vzorcem ZAOKROUHLIT(základ × " + rate + " / 1000; 0)." : "Sazba není uvedena; porovnávají se poměry všech čtvrtletí. Doplň správnou sazbu, pokud předpis chybí.");
    const ratioConsistent = ratioValues.length > 0 && Math.max(...ratioValues) - Math.min(...ratioValues) <= 0.35;
    add("Poměr pojistného k základu je mezi čtvrtletími konzistentní", ratioConsistent,
      ratioConsistent ? "Poměr odpovídá stejné sazbě s přihlédnutím k zaokrouhlení." : "Alespoň jedna částka pravděpodobně obsahuje chybu čtení nebo patří k jinému předpisu.");
    add("OCR nezanechalo nevyřešené platební řádky", state.unparsedPaymentLines.length === 0,
      state.unparsedPaymentLines.length ? "Nezařazené řádky: " + state.unparsedPaymentLines.length + ". Zkontroluj OCR text screenshotu." : "Všechny číselné řádky plateb byly zařazeny.");
    const positionCheck = state.basePositionCheck;
    add("Základ a pojistné jsou oddělené podle sloupců",
      Boolean(positionCheck && positionCheck.ok && positionCheck.matchedCount === quarters.length),
      positionCheck ? positionCheck.detail : "Nepodařilo se ověřit, že základ a pojistné byly načteny ze samostatných buněk.");
    if (tests.length !== 20) throw new Error("Interní chyba: kontrolní sada musí obsahovat přesně 20 testů.");
    return tests;
  }

  function refreshVerification() {
    const panel = $("verificationPanel");
    if (!panel) return [];
    const hasInputs = Boolean(state.paymentFile && state.baseFile && (state.paymentConsensus || state.baseConsensus));
    panel.hidden = !hasInputs;
    if (!hasInputs) {
      $("calculateButton").disabled = true;
      return [];
    }
    const tests = buildVerificationTests();
    state.verificationTests = tests;
    const passed = tests.filter(test => test.pass).length;
    $("verificationTitle").textContent = "20 kontrolních testů";
    $("verificationSubtitle").textContent = "Kontroly upozorňují na možná rizika. Výpočet můžeš spustit i při neúspěšných testech.";
    $("verificationScore").textContent = passed + "/20 úspěšných";
    $("verificationScore").className = "verification-score " + (passed === 20 ? "pass" : "fail");
    $("verificationProgress").style.width = (passed / 20 * 100) + "%";
    const results = $("verificationResults");
    results.replaceChildren();
    tests.forEach((test, index) => {
      const row = document.createElement("div");
      row.className = "verification-item " + (test.pass ? "pass" : "fail");
      const icon = element("span", { className:"v-icon" }, test.pass ? "✓" : "!");
      const copy = document.createElement("div");
      const title = element("strong", {}, String(index + 1).padStart(2, "0") + ". " + test.name);
      const detail = element("div", {}, test.detail);
      detail.className = "verification-detail";
      copy.append(title, detail);
      row.append(icon, copy);
      results.append(row);
    });
    const failures = tests.filter(test => !test.pass);
    const foot = $("verificationFoot");
    foot.classList.toggle("blocked", failures.length > 0);
    foot.textContent = failures.length
      ? "UPOZORNĚNÍ: " + failures.length + " z 20 kontrol neprošlo. Výpočet je dostupný, ale ověř neúspěšné kontroly a výsledné částky před použitím."
      : "Všech 20 testů prošlo. Výsledek přesto zkontroluj před použitím.";
    $("calculateButton").disabled = false;
    return tests;
  }

  function invalidateResults() {
    $("resultsPanel").hidden = true;
    refreshVerification();
  }

  $("addPayment").addEventListener("click", () => {
    state.payments.push({ date: isoToday(), requestedAmount: "", amount: "", raw: "Ruční záznam" });
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
        const fee = templateLateIncrease(allocated, delay);
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
        const fee = templateLateIncrease(q.remaining, delay);
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

    return { asOf, quarters, payments, paymentRows: allPayments, detail, totals, warnings };
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
      const tests = refreshVerification();
      const failed = tests.filter(test => !test.pass);
      state.results = computeResults();
      if (failed.length) {
        state.results.warnings = (state.results.warnings || []).concat(
          "Před výpočtem neprošlo " + failed.length + " z 20 kontrol: " + failed.map(test => test.name).join("; ") + ". Výsledek pečlivě ověř před použitím."
        );
      }
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
    refreshVerification();
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
      ["Předepsané řádné pojistné", { formula: "SUM('Přehled čtvrtletí'!E2:E41)" }, "Součet čtvrtletních předpisů – vzorec"],
      ["Celkem skutečně zaplaceno", { formula: "SUMIFS('Vstupní data'!C3:C102,'Vstupní data'!A3:A102,\"<>\")" }, "Součet částek ze vstupních dat – vzorec"],
      ["Započteno na řádné pojistné", { formula: "SUM('Přehled čtvrtletí'!F2:F41)" }, "Přiřazení nejstaršímu nedoplatku – vzorec"],
      ["Zbývající řádné pojistné", { formula: "SUM('Přehled čtvrtletí'!G2:G41)" }, "Jistina po započtení plateb – vzorec"],
      ["Navýšení za prodlení celkem", { formula: "SUM('Přehled čtvrtletí'!H2:H41)" }, "Navýšení z uhrazených částek a otevřené jistiny – vzorec"],
      ["Uhrazené navýšení", { formula: "SUM('Přehled čtvrtletí'!I2:I41)" }, "Platby použité až po úhradě řádného pojistného – vzorec"],
      ["Zbývající navýšení", { formula: "SUM('Přehled čtvrtletí'!J2:J41)" }, "Neuhrazené navýšení za prodlení – vzorec"],
      ["Celkem zbývá uhradit", { formula: "SUM('Přehled čtvrtletí'!G2:G41)+SUM('Přehled čtvrtletí'!J2:J41)" }, "Zbývající pojistné a neuhrazené navýšení – vzorec"],
      ["Nezařazený přeplatek", { formula: "MAX(0,SUMIFS('Vstupní data'!C3:C102,'Vstupní data'!A3:A102,\"<>\")-SUM('Přehled čtvrtletí'!E2:E41)-SUM('Přehled čtvrtletí'!I2:I41))" }, "Platby nad rámec jistiny a navýšení – vzorec"]
    ].forEach(row => summary.addRow(row));
    styleBodyRows(summary, 5);
    summary.getColumn(2).numFmt = "#,##0;[Red]-#,##0";
    summary.addRow([]);
    const notice = summary.addRow(["Poznámka", "Souhrn, přehled čtvrtletí a započtení plateb obsahují vzorce. Uprav vstupy na listu Vstupní data; platby zadávej chronologicky. Listy se přepočítají při otevření v Excelu."]);
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

    const paymentStart = 3, paymentCapacity = 100, paymentEnd = paymentStart + paymentCapacity - 1;
    const quarterStart = 106, quarterCapacity = 40, quarterEnd = quarterStart + quarterCapacity - 1;

    const inputs = workbook.addWorksheet("Vstupní data", { views: [{ state: "frozen", ySplit: 2 }] });
    inputs.addRow(["VSTUPNÍ DATA – upravuj modře označené buňky; výpočtové listy obsahují vzorce"]); styleHeader(inputs.getRow(1));
    inputs.addRow(["Datum platby","Částka platby (Kč)","Vyřízeno (Kč)","Započteno jako platba?","Rozpoznaný řádek","","Datum výpočtu"]);
    styleHeader(inputs.getRow(2));
    const sortedInputPayments = r.paymentRows.slice().sort((a,b) => String(a.date || "").localeCompare(String(b.date || "")));
    for (let i=0;i<paymentCapacity;i++) {
      const p = sortedInputPayments[i];
      const row = inputs.addRow(p ? [parseDate(p.date) || null, Number(p.requestedAmount == null ? p.amount : p.requestedAmount) || 0, Number(p.amount) || 0, Number(p.amount) > 0 ? "Ano" : "Ne — 0 Kč vyřízeno", p.raw || "Ruční záznam"] : [null,null,null,"",""]);
      row.getCell(1).numFmt = "dd.mm.yyyy";
      [1,2,3].forEach(ci => row.getCell(ci).font = { color: { argb: "FF0070C0" } });
    }
    inputs.getCell("G2").value = parseDate(r.asOf) || null;
    inputs.getCell("G2").numFmt = "dd.mm.yyyy";
    inputs.getCell("G2").font = { color: { argb: "FF0070C0" } };
    inputs.getCell("G1").font = { bold: true, color: { argb: "FFFFFFFF" } };
    inputs.addRow([]);
    inputs.addRow(["PŘEDPISY POJISTNÉHO"]); styleHeader(inputs.getRow(quarterStart-2));
    inputs.addRow(["Čtvrtletí","Splatnost","Období pro výpočet","Základ (Kč)","Řádné pojistné (Kč)","Původní řádek"]);
    styleHeader(inputs.getRow(quarterStart-1));
    for (let i=0;i<quarterCapacity;i++) {
      const q = r.quarters[i];
      const row = inputs.addRow(q ? [q.period, parseDate(q.dueDate) || null, q.sourcePeriod || "", Number(q.base) || 0, Number(q.premium) || 0, q.raw || "Ruční záznam"] : [null,null,"",null,null,""]);
      row.getCell(2).numFmt = "dd.mm.yyyy";
      [1,2,3,4,5].forEach(ci => row.getCell(ci).font = { color: { argb: "FF0070C0" } });
    }
    inputs.columns = [{width:18},{width:21},{width:22},{width:26},{width:24},{width:48},{width:4},{width:20}];
    inputs.getCell("G3").value = "Platby zadávej chronologicky. Pro nové řádky použij volná místa.";
    inputs.mergeCells("G3:H3");
    inputs.getCell("G3").alignment = { wrapText: true };
    inputs.getCell("G4").value = "Datum výpočtu je v G2.";
    inputs.mergeCells("G4:H4");
    inputs.autoFilter = { from: "A2", to: "E" + paymentEnd };

    const calc = workbook.addWorksheet("Přehled čtvrtletí", { views: [{ state: "frozen", ySplit: 1 }] });
    calc.columns = [
      {width:14},{width:14},{width:19},{width:21},{width:20},
      {width:22},{width:23},{width:20},{width:22},{width:20}
    ];
    const qhead = calc.addRow(["Čtvrtletí","Splatnost","Období pro výpočet","Výpočtový základ (Kč)","Řádné pojistné (Kč)","Uhrazeno na pojistné (Kč)","Nedoplatek pojistného (Kč)","Navýšení celkem (Kč)","Navýšení uhrazeno (Kč)","Navýšení zbývá (Kč)"]); styleHeader(qhead);
    for (let i=0;i<quarterCapacity;i++) {
      const rowNo = i+2, inputRow = quarterStart+i;
      const prevPremium = i===0 ? "0" : `SUM($E$2:E${rowNo-1})`;
      const priorPaidFee = i===0 ? "0" : `SUM($I$2:I${rowNo-1})`;
      const row = calc.addRow([
        {formula:`IF('Vstupní data'!A${inputRow}="","",'Vstupní data'!A${inputRow})`},
        {formula:`IF(A${rowNo}="","",'Vstupní data'!B${inputRow})`},
        {formula:`IF(A${rowNo}="","",'Vstupní data'!C${inputRow})`},
        {formula:`IF(A${rowNo}="","",'Vstupní data'!D${inputRow})`},
        {formula:`IF(A${rowNo}="","",IF(OR(D${rowNo}="", 'List1'!$I$1=""),'Vstupní data'!E${inputRow},ROUND(D${rowNo}*'List1'!$I$1/1000,0)))`},
        {formula:`IF(A${rowNo}="","",MIN(E${rowNo},MAX(0,SUMIFS('Vstupní data'!$C$${paymentStart}:$C$${paymentEnd},'Vstupní data'!$A$${paymentStart}:$A$${paymentEnd},"<>")-${prevPremium})))`},
        {formula:`IF(A${rowNo}="","",MAX(0,E${rowNo}-F${rowNo}))`},
        {formula:`IF(A${rowNo}="","",SUMIF('Započtení plateb'!$B$2:$B$${1+paymentCapacity*quarterCapacity},A${rowNo},'Započtení plateb'!$E$2:$E$${1+paymentCapacity*quarterCapacity})+IF(G${rowNo}>0,ROUND(G${rowNo}*IF('Vstupní data'!$G$2<=B${rowNo},0,MAX(0,(YEAR('Vstupní data'!$G$2)-YEAR(B${rowNo}))*12+MONTH('Vstupní data'!$G$2)-MONTH(B${rowNo}))/10),0),0))`},
        {formula:`IF(A${rowNo}="","",MIN(SUMIF('Započtení plateb'!$B$2:$B$${1+paymentCapacity*quarterCapacity},A${rowNo},'Započtení plateb'!$E$2:$E$${1+paymentCapacity*quarterCapacity}),MAX(0,SUMIFS('Vstupní data'!$C$${paymentStart}:$C$${paymentEnd},'Vstupní data'!$A$${paymentStart}:$A$${paymentEnd},"<>")-SUM($E$2:$E$${1+quarterCapacity})-${priorPaidFee})))`},
        {formula:`IF(A${rowNo}="","",MAX(0,H${rowNo}-I${rowNo}))`}
      ]);
      row.getCell(2).numFmt = "dd.mm.yyyy";
      [4,5,6,7,8,9,10].forEach(ci=>row.getCell(ci).numFmt="#,##0;[Red]-#,##0");
    }
    styleBodyRows(calc, 2);
    calc.autoFilter = { from: "A1", to: "J" + calc.rowCount };

    const alloc = workbook.addWorksheet("Započtení plateb", { views: [{ state: "frozen", ySplit: 1 }] });
    alloc.columns = [{width:14},{width:14},{width:20},{width:20},{width:18},{width:44}];
    const allocHead = alloc.addRow(["Datum platby","Čtvrtletí","Přiřazené pojistné (Kč)","Prodlení (měsíce)","Navýšení (Kč)","Poznámka"]); styleHeader(allocHead);
    let allocRowNo = 2;
    for (let pi=0;pi<paymentCapacity;pi++) {
      const paymentInputRow = paymentStart+pi;
      for (let qi=0;qi<quarterCapacity;qi++) {
        const quarterInputRow = quarterStart+qi;
        const quarterCalcRow = 2+qi;
        const priorPaid = pi===0 ? "0" : `SUMIFS('Vstupní data'!$C$${paymentStart}:$C$${paymentInputRow-1},'Vstupní data'!$A$${paymentStart}:$A$${paymentInputRow-1},"<>")`;
        const paidThrough = `SUMIFS('Vstupní data'!$C$${paymentStart}:$C$${paymentInputRow},'Vstupní data'!$A$${paymentStart}:$A$${paymentInputRow},"<>")`;
        const priorPremium = qi===0 ? "0" : `SUM('Přehled čtvrtletí'!$E$2:$E$${quarterCalcRow-1})`;
        const premiumThrough = `SUM('Přehled čtvrtletí'!$E$2:$E$${quarterCalcRow})`;
        const row = alloc.addRow([
          {formula:`IF('Vstupní data'!A${paymentInputRow}="","",'Vstupní data'!A${paymentInputRow})`},
          {formula:`IF('Vstupní data'!A${quarterInputRow}="","",'Vstupní data'!A${quarterInputRow})`},
          {formula:`IF(OR('Vstupní data'!$A$${paymentInputRow}="", 'Vstupní data'!$A$${quarterInputRow}="", 'Vstupní data'!$C$${paymentInputRow}<=0),0,MAX(0,MIN(${paidThrough},${premiumThrough})-MAX(${priorPaid},${priorPremium})))`},
          {formula:`IF(OR(A${allocRowNo}="",C${allocRowNo}<=0),0,IF(A${allocRowNo}<='Vstupní data'!$B$${quarterInputRow},0,MAX(0,(YEAR(A${allocRowNo})-YEAR('Vstupní data'!$B$${quarterInputRow}))*12+MONTH(A${allocRowNo})-MONTH('Vstupní data'!$B$${quarterInputRow}))))`},
          {formula:`ROUND(C${allocRowNo}*D${allocRowNo}/10,0)`},
          {formula:`IF(C${allocRowNo}=0,"","Automatické rozdělení: nejstarší jistina první")`}
        ]);
        row.getCell(1).numFmt = "dd.mm.yyyy";
        [3,5].forEach(ci=>row.getCell(ci).numFmt="#,##0;[Red]-#,##0");
        allocRowNo++;
      }
    }
    styleBodyRows(alloc, 2);
    alloc.autoFilter = { from: "A1", to: "F" + alloc.rowCount };

    const warnings = workbook.addWorksheet("Kontrola");
    warnings.columns = [{width:22},{width:100}];
    warnings.addRow(["KONTROLA VÝPOČTU","VÝSLEDEK"]); styleHeader(warnings.getRow(1));
    warnings.addRow(["Datum výpočtu",dateFmt(r.asOf)]);
    warnings.addRow(["Vzorec navýšení","I = H/10; J = ROUND(částka × I; 0). U částečné úhrady se počítá každá alokovaná část zvlášť; u otevřeného zůstatku se použije zbývající jistina."]);
    warnings.addRow(["Pravidlo přiřazení","Platby se započítávají nejprve na nejstarší neuhrazené řádné pojistné. Navýšení se hradí až po úhradě řádného pojistného."]);
    warnings.addRow(["Zdrojové soubory","Obsah vstupů se zpracoval v prohlížeči; do repozitáře ani na aplikační server nebyl poslán."]);
    warnings.addRow(["Vzorce v Excelu","Souhrn, přehled čtvrtletí a započtení plateb jsou propojeny vzorci. List Vstupní data obsahuje 100 řádků pro platby a 40 řádků pro čtvrtletní předpisy. Nové platby přidávej do volných řádků chronologicky."]);
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