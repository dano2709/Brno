const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const start = source.indexOf("  function computeResults() {");
const end = source.indexOf("  function renderResults(result) {", start);
assert.ok(start >= 0 && end > start, "computeResults() must exist in app.js");
const computeSource = source.slice(start, end);

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}
function validISODate(value) {
  const d = parseDate(value);
  return !!d && d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0") === value;
}
function monthsLate(dueValue, paidValue) {
  const due = parseDate(dueValue), paid = parseDate(paidValue);
  if (!due || !paid || paid <= due) return 0;
  return Math.max(0, (paid.getFullYear()-due.getFullYear())*12 + paid.getMonth()-due.getMonth());
}
function templateLateIncrease(amount, delayMonths) {
  return Math.round(Math.max(0, Number(amount)||0) * Math.max(0, Math.floor(Number(delayMonths)||0)) / 10);
}
function parseAmount(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const raw = String(value ?? "").trim().replace(/\s/g, "").replace(/[^\d,.-]/g, "").replace(",", ".");
  const n = Number(raw);
  return raw && Number.isFinite(n) ? n : null;
}
const dateFmt = value => value || "";
const currencyFmt = value => Math.round(Number(value)||0).toLocaleString("cs-CZ") + " Kč";
const engineFactory = new Function(
  "state","$","isoToday","validISODate","monthsLate","dateFmt","currencyFmt","parseAmount","templateLateIncrease",
  computeSource + "\nreturn computeResults;"
);

const quarterFixture = [
  ["Q1/2024","2024-01-31",1000],
  ["Q2/2024","2024-04-30",2000],
  ["Q3/2024","2024-07-31",1500]
].map(([period,dueDate,premium]) => ({ period,dueDate,sourcePeriod:"",base:0,premium,raw:"test fixture" }));
const paymentFixture = [
  {date:"2024-01-15",requestedAmount:500,amount:500,raw:"500.00 500.00 15.01.2024"},
  {date:"2024-05-10",requestedAmount:800,amount:800,raw:"800.00 800.00 10.05.2024"},
  {date:"2024-08-20",requestedAmount:1800,amount:1800,raw:"1800.00 1800.00 20.08.2024"},
  {date:"2024-09-30",requestedAmount:5000,amount:0,raw:"5000.00 0.00 30.09.2024"}
];
const expected = {
  premium:4500,
  principalPaid:3100,
  principalOpen:1400,
  feeAssessed:1200,
  feePaid:0,
  feeOpen:1200,
  paymentTotal:3100,
  overpayment:0,
  due:2600,
  quarters:[
    {period:"Q1/2024",paidPrincipal:1000,remaining:0,totalAssessedFee:200,feeUnpaid:200},
    {period:"Q2/2024",paidPrincipal:2000,remaining:0,totalAssessedFee:710,feeUnpaid:710},
    {period:"Q3/2024",paidPrincipal:100,remaining:1400,totalAssessedFee:290,feeUnpaid:290}
  ]
};

function seededShuffle(items, seed) {
  const out = items.slice();
  let x = seed >>> 0;
  for (let i=out.length-1;i>0;i--) {
    x = (1664525*x + 1013904223) >>> 0;
    const j = x % (i+1);
    [out[i],out[j]] = [out[j],out[i]];
  }
  return out;
}
function runApplicationEngine(quarters, payments) {
  const state = {
    quarters: quarters.map(q=>({...q})),
    payments: payments.map(p=>({...p})),
    unparsedPaymentLines:[],
    sourceRate:null
  };
  const controls = id => ({value: id==="asOfDate" ? "2024-09-30" : ""});
  const calculate = engineFactory(state,controls,()=>"2024-09-30",validISODate,monthsLate,dateFmt,currencyFmt,parseAmount,templateLateIncrease);
  return calculate();
}

// Independent model of the exported Excel payment-allocation matrix.
function formulaSheetModel(quartersInput, paymentsInput, asOf) {
  const quarters = quartersInput.map(q=>({...q})).sort((a,b)=>a.dueDate.localeCompare(b.dueDate));
  const payments = paymentsInput.map((p,index)=>({...p,index})).filter(p=>validISODate(p.date) && Number(p.amount)>0 && p.date<=asOf)
    .sort((a,b)=>a.date.localeCompare(b.date)||a.index-b.index);
  let cashBefore = 0;
  const premiumBefore = [];
  let runningPremium = 0;
  for (const q of quarters) { premiumBefore.push(runningPremium); runningPremium += Number(q.premium); }
  const allocations = [];
  for (const p of payments) {
    const paidBefore = cashBefore;
    const paidThrough = cashBefore + Number(p.amount);
    for (let i=0;i<quarters.length;i++) {
      const q = quarters[i];
      const premiumThrough = premiumBefore[i] + Number(q.premium);
      const amount = Math.max(0, Math.min(paidThrough,premiumThrough)-Math.max(paidBefore,premiumBefore[i]));
      if (amount>0) allocations.push({date:p.date,period:q.period,amount,delayMonths:monthsLate(q.dueDate,p.date)});
    }
    cashBefore = paidThrough;
  }
  const groupedFees = new Map();
  const paidByQuarter = new Map();
  for (const a of allocations) {
    const key=a.date+"|"+a.period;
    groupedFees.set(key,{...a,amount:(groupedFees.get(key)?.amount||0)+a.amount});
    paidByQuarter.set(a.period,(paidByQuarter.get(a.period)||0)+a.amount);
  }
  const byQuarter = quarters.map(q=>{
    let paidFee=0;
    for (const [,a] of groupedFees) if (a.period===q.period) paidFee += templateLateIncrease(a.amount,a.delayMonths);
    const paid=paidByQuarter.get(q.period)||0;
    const remaining=Math.max(0,Number(q.premium)-paid);
    const openFee=remaining ? templateLateIncrease(remaining,monthsLate(q.dueDate,asOf)) : 0;
    return {period:q.period,paidPrincipal:Math.round(paid),remaining:Math.round(remaining),totalAssessedFee:paidFee+openFee,feeUnpaid:paidFee+openFee};
  });
  const totals={
    premium:quarters.reduce((sum,q)=>sum+Number(q.premium),0),
    principalPaid:byQuarter.reduce((sum,q)=>sum+q.paidPrincipal,0),
    principalOpen:byQuarter.reduce((sum,q)=>sum+q.remaining,0),
    feeAssessed:byQuarter.reduce((sum,q)=>sum+q.totalAssessedFee,0),
    feePaid:0,
    feeOpen:byQuarter.reduce((sum,q)=>sum+q.feeUnpaid,0),
    paymentTotal:payments.reduce((sum,p)=>sum+Number(p.amount),0),
    overpayment:0,
    due:byQuarter.reduce((sum,q)=>sum+q.remaining+q.feeUnpaid,0)
  };
  return {totals,quarters:byQuarter};
}

for (let run=1;run<=10;run++) {
  const quarters = run%2 ? quarterFixture : seededShuffle(quarterFixture,run);
  const payments = seededShuffle(paymentFixture,run*173);
  const appResult = runApplicationEngine(quarters,payments);
  const formulaResult = formulaSheetModel(quarters,payments,"2024-09-30");
  assert.deepEqual(appResult.totals, expected, "application totals, run "+run);
  assert.deepEqual(
    appResult.quarters.map(q=>({period:q.period,paidPrincipal:q.paidPrincipal,remaining:q.remaining,totalAssessedFee:q.totalAssessedFee,feeUnpaid:q.feeUnpaid})),
    expected.quarters,
    "application quarterly results, run "+run
  );
  assert.deepEqual(formulaResult.totals, expected, "independent Excel-formula model totals, run "+run);
  assert.deepEqual(formulaResult.quarters, expected.quarters, "independent Excel-formula model quarterly results, run "+run);
}
assert.ok(source.includes('"<= " &'), "export formulas must support an as-of-date limit");
assert.ok(source.includes("paymentInputRow}>'Vstupní data'!$G$2"), "payment allocation formulas must exclude payments after the selected calculation date");
console.log("PASS: 10/10 runs matched application and independent Excel-formula model.");
