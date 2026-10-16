const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const begin = source.indexOf("  function normalizeOcrWord(");
const finish = source.indexOf("  function aggregatePaymentReadings(", begin);
assert.ok(begin >= 0 && finish > begin, "OCR position parser must exist");
const parserSource = source.slice(begin, finish);

function parseAmount(value) {
  const normalized = String(value ?? "").replace(/\s/g, "").replace(/[^\d,.-]/g, "").replace(",", ".");
  const number = Number(normalized);
  return normalized && Number.isFinite(number) ? number : null;
}
function quarterMatch(value) {
  const text = String(value ?? "").trim();
  const match = text.match(/(?:^|\b)Q\s*([1-4])\s*[\/.-]\s*(\d{2,4})\b/i)
    || text.match(/(?:^|\b)([1-4])\s*Q\s*[\/.-]\s*(\d{2,4})\b/i);
  if (!match) return null;
  const quarter = Number(match[1]);
  let year = Number(match[2]);
  if (year < 100) year += year >= 70 ? 1900 : 2000;
  return {period:"Q"+quarter+"/"+year,quarter,year};
}
function dateMatches(line) {
  const re = /\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g;
  return Array.from(String(line).matchAll(re),m=>({text:m[0],index:m.index,end:m.index+m[0].length}));
}
function toISO(value) {
  if (!value) return "";
  const match=String(value).match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (!match) return "";
  let year=Number(match[3]); if(year<100)year+=year>=70?1900:2000;
  const month=Number(match[2]),day=Number(match[1]);
  const d=new Date(year,month-1,day);
  if(d.getFullYear()!==year||d.getMonth()!==month-1||d.getDate()!==day)return "";
  return year+"-"+String(month).padStart(2,"0")+"-"+String(day).padStart(2,"0");
}
function quarterFromDueDate(value) {
  const d=value ? new Date(value+"T00:00:00") : null;
  if(!d||Number.isNaN(d.getTime()))return null;
  const month=d.getMonth()+1, quarter=month===1?1:month===4?2:month===7?3:month===10?4:null;
  return quarter?{period:"Q"+quarter+"/"+d.getFullYear(),quarter,year:d.getFullYear()}:null;
}
function validationPeriodOrdinal(value) {
  const q=quarterMatch(value);return q?q.year*4+q.quarter-1:null;
}
function previousQuarterPeriod(info) {
  const prevQuarter=info.quarter===1?4:info.quarter-1;
  const prevYear=info.quarter===1?info.year-1:info.year;
  return "Q"+prevQuarter+"/"+prevYear;
}
const parsePositionRows = new Function(
  "parseAmount","quarterMatch","dateMatches","toISO","quarterFromDueDate","validationPeriodOrdinal","previousQuarterPeriod",
  parserSource+"\nreturn positionRowsFromData;"
)(parseAmount,quarterMatch,dateMatches,toISO,quarterFromDueDate,validationPeriodOrdinal,previousQuarterPeriod);

const expected = [
  ["Q4/2023","2023-10-31","Q3/2023",258475,2171],
  ["Q1/2024","2024-01-31","Q4/2023",213833,1796],
  ["Q2/2024","2024-04-30","Q1/2024",260420,2188],
  ["Q3/2024","2024-07-31","Q2/2024",322745,2711],
  ["Q4/2024","2024-10-31","Q3/2024",277675,2332],
  ["Q1/2025","2025-01-31","Q4/2024",222154,1866],
  ["Q2/2025","2025-04-30","Q1/2025",174811,1468],
  ["Q3/2025","2025-07-31","Q2/2025",304593,2559],
  ["Q4/2025","2025-10-31","Q3/2025",296812,2493],
  ["Q1/2026","2026-01-31","Q4/2025",218732,1837],
  ["Q2/2026","2026-04-30","Q1/2026",150430,1264],
  ["Q3/2026","2026-07-31","Q2/2026",271714,2282]
];

// Build OCR word boxes from the source table. Rows are close together relative to the
// overall screenshot height; some number tokens have slightly different vertical centers.
const words=[];
function word(text,cx,cy,width=text.length*7,wordHeight=12,confidence=95) {
  words.push({text,confidence,bbox:{x0:cx-width/2,x1:cx+width/2,y0:cy-wordHeight/2,y1:cy+wordHeight/2}});
}
const headerY=45;
word("na",38,headerY,18);word("obdobi",60,headerY,44);
word("splatnost",145,headerY,60);
word("obdobi",220,headerY,44);word("pro",250,headerY,18);word("vypocet",282,headerY,50);
word("zpusob",405,headerY,44);word("zamestnavani",463,headerY,86);
word("zaklad",625,headerY,48);word("pro",660,headerY,18);word("vypocet",690,headerY,50);
word("radne",835,headerY,42);word("pojistne",886,headerY,55);

// A helper Q3/2023 row has no base/premium; it must not duplicate the next row.
word("Q3/2023",48,85,60);
const dateFor = value => {
 const [year,month,day]=value.split("-");
 return day+"."+month+"."+year;
};
expected.forEach((row,index)=>{
  const [period,dueDate,sourcePeriod,base,premium]=row;
  const cy=120+index*30;
  word(period,48,cy,60);
  word(dateFor(dueDate),145,cy,72);
  word(sourcePeriod,240,cy,60);
  word("bezne",400,cy,40);
  // Deliberately make some thousands groups vertically offset and reverse their
  // vertical order. The parser must still concatenate the base digits left-to-right.
  const baseText=String(base);
  const baseSplit=baseText.slice(0,3);
  const baseRest=baseText.slice(3);
  word(baseSplit,610,cy+(index%3===0?3:-3),baseSplit.length*8);
  word(baseRest,680,cy+(index%3===0?-3:3),baseRest.length*8);
  const premiumText=String(premium);
  const premiumSplit=premiumText.slice(0,premiumText.length-3);
  const premiumRest=premiumText.slice(-3);
  word(premiumSplit,800,cy+(index%2===0?2:-2),premiumSplit.length*8);
  word(premiumRest,845,cy+(index%2===0?-2:2),premiumRest.length*8);
  word("Kc",895,cy,18);
});
// A footer makes the whole OCR canvas tall, the exact condition that exposed the old
// screenshot-relative row tolerance.
word("F7-Detail",80,2480,65,12);

for(let run=1;run<=10;run++){
  // Simulate small OCR bounding-box jitter from different image preprocessing / PSM modes.
  const jittered=words.map((item,index)=>{
    const offset=((index*7+run*3)%7)-3;
    return {...item,bbox:{...item.bbox,y0:item.bbox.y0+offset,y1:item.bbox.y1+offset}};
  });
  const result=parsePositionRows({words:jittered});
  const actual=result.rows.map(row=>[row.period,row.dueDate,row.sourcePeriod,row.base,row.premium]);
  assert.deepEqual(actual,expected,"OCR values must match on synthetic OCR pass "+run);
  const q42024=actual.find(row=>row[0]==="Q4/2024");
  const q12025=actual.find(row=>row[0]==="Q1/2025");
  assert.deepEqual(q42024,["Q4/2024","2024-10-31","Q3/2024",277675,2332],"Q4/2024 must never reuse Q1/2025 values");
  assert.deepEqual(q12025,["Q1/2025","2025-01-31","Q4/2024",222154,1866],"Q1/2025 values must stay on their own row");
  assert.equal(actual.some(row=>row[0]==="Q3/2023"),false,"the blank helper Q3/2023 row must not be treated as a premium assessment");
}
console.log("PASS: 10/10 OCR-position simulations recover all 12 quarters and exclude the blank helper row.");
