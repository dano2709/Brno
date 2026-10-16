const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source=fs.readFileSync(path.join(__dirname,"..","app.js"),"utf8");
const parseStart=source.indexOf("  function parseDate(value)");
const parseEnd=source.indexOf("\n  function toISO(value)",parseStart);
const serialStart=source.indexOf("  function excelDateSerial(value)");
const serialEnd=source.indexOf("\n  function parseAmount(value)",serialStart);
assert.ok(parseStart>=0&&parseEnd>parseStart,"parseDate helper must exist");
assert.ok(serialStart>=0&&serialEnd>serialStart,"excelDateSerial helper must exist");

const parseDateSource=source.slice(parseStart,parseEnd);
const serialSource=source.slice(serialStart,serialEnd);
const helpers=new Function("window",parseDateSource+"\n"+serialSource+"\nreturn {excelDateSerial};")({XLSX:null});

const tests=[
  ["ISO string", "2024-10-18", 45583],
  ["Czech date", "18.10.2024", 45583],
  ["January date", "31.01.2025", 45688],
  ["start of statement", "24.10.2023", 45223],
  ["zero-padded ISO", "2026-09-30", 46295]
];
for(const [label,input,expected] of tests){
  assert.equal(helpers.excelDateSerial(input),expected,label+" must map to the correct Excel serial date");
}

// Verify every exported date field uses timezone-safe serial numbers, not JS Date objects.
const exportStart=source.indexOf("  async function downloadWorkbook()");
const exportEnd=source.indexOf('  $("downloadButton")',exportStart);
const exportSource=source.slice(exportStart,exportEnd);
assert.ok(exportSource.includes("excelDateSerial(q.dueDate)"),"due dates in exported worksheets must use Excel serial numbers");
assert.ok(exportSource.includes("excelDateSerial(allocation.date)"),"payment dates in legacy worksheet must use Excel serial numbers");
assert.ok(exportSource.includes("excelDateSerial(p.date)"),"input payment dates must use Excel serial numbers");
assert.ok(exportSource.includes("excelDateSerial(r.asOf)"),"calculation cutoff date must use Excel serial number");
assert.ok(!/parseDate\((?:p\.date|q\.dueDate|allocation\.date|r\.asOf)\)/.test(exportSource),"no exported input dates may be written as timezone-sensitive Date objects");
console.log("PASS: Excel dates retain the original calendar day in all exported sheets.");
