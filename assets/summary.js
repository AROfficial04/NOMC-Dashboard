// summary.js
// Dynamic Summary + Charts from meter_sample_filled.xlsx

let workbookData = [];
let timelineChart;
let currentWeekIndex = 0; // track which week is shown
let allDays = []; // sorted list of all days with data

// --- helpers ---
const norm = v => (v ?? "").toString().trim();
const isAffirmative = v => {
  const s = norm(v).toLowerCase();
  return s === "yes" || s === "approved";
};

// Parse many date formats + Excel serials
function parseDateAny(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = norm(v);
  let m = s.match(/^(\d{2})[-/](\d{2})[-/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// classify comm
function classifyComm(lastComm) {
  const s = norm(lastComm).toLowerCase();
  if (!s) return "non";
  if (s.includes("never")) return "never";
  const d = parseDateAny(lastComm);
  if (!d) return "non";
  const diffDays = (Date.now() - d.getTime()) / 86400000;
  return diffDays > 3 ? "non" : "comm";
}

// --- load + build ---
async function loadExcel() {
  const response = await fetch("assets/meter_sample_filled.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { defval: null });

  // global
  window.data = data;
  workbookData = data;
  kpiData = data;

  // Now call all renderers
  renderKpis();
  const grouped = buildHierarchy(data);
  renderSummaryTable(grouped);
  prepareTimelineData();
  renderTimeline();
  renderCharts();
}



function buildHierarchy(rows) {
  const regions = {};
  rows.forEach(r => {
    const region = r["Region Name"] || "Unknown";
    const type = r["Meter Type"] || "Unknown";
    const comm = r["Comm Medium"] || "Unknown";
    regions[region] ??= {};
    regions[region][type] ??= {};
    regions[region][type][comm] ??= {
      total: 0, comm: 0, nonComm: 0, neverComm: 0,
      l1: 0, l2: 0, de: 0, mdm: 0, sap: 0, sat: 0
    };
    const bucket = regions[region][type][comm];
    bucket.total++;
    const status = classifyComm(r["LastComm"]);
    if (status === "comm") bucket.comm++;
    else if (status === "non") bucket.nonComm++;
    else bucket.neverComm++;
    if (isAffirmative(r["L1"])) bucket.l1++;
    if (isAffirmative(r["L2"])) bucket.l2++;
    if (isAffirmative(r["Daily Energy"])) bucket.de++;
    if (isAffirmative(r["MDM"])) bucket.mdm++;
    if (isAffirmative(r["SAP"])) bucket.sap++;
    if (isAffirmative(r["SAT"])) bucket.sat++;
  });
  return Object.keys(regions).map(regionName => ({
    name: regionName,
    meterTypes: Object.keys(regions[regionName]).map(type => ({
      type,
      commMediums: Object.keys(regions[regionName][type]).map(comm => ({
        name: comm,
        ...regions[regionName][type][comm]
      }))
    }))
  }));
}

// --- render summary table (same as before, kept) ---
function renderSummaryTable(summaryData) {
  const tbody = document.getElementById("summaryTableBody");
  tbody.innerHTML = "";
  const sumKeys = ["total","comm","nonComm","neverComm","l1","l2","de","mdm","sap","sat"];
  const addMetricsRow = (tr, obj, bold = false) => {
    sumKeys.forEach(k => {
      const td = document.createElement("td");
      td.textContent = obj[k] ?? 0;
      if (bold) td.style.fontWeight = "600";
      tr.appendChild(td);
    });
  };
  const grand = { total:0, comm:0, nonComm:0, neverComm:0, l1:0, l2:0, de:0, mdm:0, sap:0, sat:0 };
  summaryData.forEach(region => {
    const regionRowCount = region.meterTypes.reduce((a,mt)=>a+mt.commMediums.length,0);
    let regionPrinted=false;
    const regionTotals={ total:0,comm:0,nonComm:0,neverComm:0,l1:0,l2:0,de:0,mdm:0,sap:0,sat:0 };
    region.meterTypes.forEach(mt=>{
      let mtPrinted=false;
      mt.commMediums.forEach(cm=>{
        const tr=document.createElement("tr");
        if(!regionPrinted){
          const tdRegion=document.createElement("td");
          tdRegion.rowSpan=regionRowCount+1;
          tdRegion.textContent=region.name;
          tr.appendChild(tdRegion);
          regionPrinted=true;
        }
        if(!mtPrinted){
          const tdType=document.createElement("td");
          tdType.rowSpan=mt.commMediums.length;
          tdType.textContent=mt.type;
          tr.appendChild(tdType);
          mtPrinted=true;
        }
        const tdComm=document.createElement("td");
        tdComm.textContent=cm.name;
        tr.appendChild(tdComm);
        addMetricsRow(tr,cm);
        tbody.appendChild(tr);
        Object.keys(regionTotals).forEach(k=>{
          regionTotals[k]+=cm[k]||0;
          grand[k]+=cm[k]||0;
        });
      });
    });
    const totalTr=document.createElement("tr");
    totalTr.classList.add("region-total");
    const tdLabel=document.createElement("td");
    tdLabel.colSpan=2; tdLabel.textContent="Total"; tdLabel.style.fontWeight="700";
    totalTr.appendChild(tdLabel);
    addMetricsRow(totalTr,regionTotals,true);
    tbody.appendChild(totalTr);
  });
  const grandTr=document.createElement("tr");
  grandTr.classList.add("grand-total");
  const tdGrand=document.createElement("td");
  tdGrand.colSpan=3; tdGrand.textContent="Grand Total"; tdGrand.style.fontWeight="700";
  grandTr.appendChild(tdGrand);
  ["total","comm","nonComm","neverComm","l1","l2","de","mdm","sap","sat"].forEach(k=>{
    const td=document.createElement("td");
    td.textContent=grand[k]; td.style.fontWeight="700";
    grandTr.appendChild(td);
  });
  tbody.appendChild(grandTr);
}

// --- timeline data prep ---
let installDaily={}, commDaily={};
function prepareTimelineData() {
  installDaily={}; commDaily={};
  workbookData.forEach(r=>{
    const d1=parseDateAny(r["Installation Date"]);
    if(d1){const k=d1.toISOString().slice(0,10); installDaily[k]=(installDaily[k]||0)+1;}
    const d2=parseDateAny(r["First Comm"]);
    if(d2){const k=d2.toISOString().slice(0,10); commDaily[k]=(commDaily[k]||0)+1;}
  });
  allDays = Array.from(new Set([...Object.keys(installDaily),...Object.keys(commDaily)])).sort();
  currentWeekIndex = Math.floor((allDays.length-1)/7); // start at last week
}

function renderTimeline() {
  const start=currentWeekIndex*7;
  const end=start+7;
  const weekDays=allDays.slice(start,end);
  const installSeries=weekDays.map(d=>installDaily[d]||0);
  const commSeries=weekDays.map(d=>commDaily[d]||0);
  if(timelineChart) timelineChart.destroy();
  timelineChart=new Chart(document.getElementById("timelineChart"),{
    type:"line",
    data:{
      labels:weekDays,
      datasets:[
        {label:"Installations",data:installSeries,borderColor:"#3b82f6",tension:0.2,fill:false},
        {label:"First Communication",data:commSeries,borderColor:"#22c55e",tension:0.2,fill:false}
      ]
    },
    options:{responsive:true,interaction:{mode:"index",intersect:false},scales:{y:{beginAtZero:true}}}
  });
}

function prevWeek(){ if(currentWeekIndex>0){currentWeekIndex--; renderTimeline();} }
function nextWeek(){ if((currentWeekIndex+1)*7<allDays.length){currentWeekIndex++; renderTimeline();} }

// --- other charts (status + region pie, unchanged) ---
function renderCharts() {
  let comm=0,non=0,never=0;
  workbookData.forEach(r=>{
    const st=classifyComm(r["LastComm"]);
    if(st==="comm")comm++; else if(st==="non")non++; else never++;
  });
  new Chart(document.getElementById("commStatusChart"),{
    type:"doughnut",
    data:{labels:["Communicating","Non-Communicating","Never Comm"],datasets:[{data:[comm,non,never]}]}
  });
  const regionCounts={};
  workbookData.forEach(r=>{
    const region=r["Region Name"]||"Unknown";
    regionCounts[region]=(regionCounts[region]||0)+1;
  });
  new Chart(document.getElementById("regionPieChart"),{
    type:"pie",
    data:{labels:Object.keys(regionCounts),datasets:[{data:Object.values(regionCounts)}]}
  });
}

// Export raw
document.getElementById("exportSummaryBtn").addEventListener("click",()=>{
  const ws=XLSX.utils.json_to_sheet(workbookData);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Summary");
  XLSX.writeFile(wb,"Meter_Summary.xlsx");
});

//kpi addon

// --- Helpers ---
// function fmtPct(val) {
//   return `${val.toFixed(1)}%`;
// }
// // Load Excel
// // async function loadExcelData() {
// //   const response = await fetch("assets/meter_sample_filled.xlsx"); // place file inside /assets
// //   const arrayBuffer = await response.arrayBuffer();
// //   const workbook = XLSX.read(arrayBuffer, { type: "array" });
// //   const sheet = workbook.Sheets[workbook.SheetNames[0]];
// //   const data = XLSX.utils.sheet_to_json(sheet);
// //   window.data = data; // store globally
// //   return data;
// // }

// // --- Utility: Parse Excel Date ---
// function parseExcelDate(value) {
//   if (!value) return null;

//   // If it's already a Date
//   if (value instanceof Date) return value;

//   // If it's Excel serial number
//   if (typeof value === "number") {
//     return new Date((value - 25569) * 86400 * 1000); // Excel -> JS
//   }

//   // Try parsing string
//   const d = new Date(value);
//   if (!isNaN(d)) return d;

//   return null; // fallback
// }

// // --- Global Totals ---
// function computeGlobalTotals(data) {
//   const rows = data || [];

//   let feeders = 0, dts = 0, meters = 0;
//   let l1Approved = 0, l2Approved = 0;
//   let mdm = 0, sap = 0;
//   let sat = 0, dailyEnergy = 0;

//   let communicating = 0, nonCommunicating = 0, neverComm = 0;
//   let unmapped = 0;

//   const today = new Date();
//   const threeDaysAgo = new Date();
//   threeDaysAgo.setDate(today.getDate() - 3);

//   rows.forEach(r => {
//     const type = (r["MeterType"] || "").toString().trim().toLowerCase();
//     if (type === "feeder") feeders++;
//     else if (type === "dt") dts++;
//     else if (type === "wc") meters++;

//     // Approval
//     if ((r["L1Approved"] || "").toString().trim().toLowerCase() === "yes") l1Approved++;
//     if ((r["L2Approved"] || "").toString().trim().toLowerCase() === "yes") l2Approved++;

//     // Integration
//     if ((r["MDM"] || "").toString().trim().toLowerCase() === "yes") mdm++;
//     if ((r["SAP"] || "").toString().trim().toLowerCase() === "yes") sap++;

//     if ((r["SAT"] || "").toString().trim().toLowerCase() === "yes") sat++;
//     if ((r["DailyEnergy"] || "").toString().trim().toLowerCase() === "yes") dailyEnergy++;

//     // Communication classification
//     const lastCommRaw = (r["LastComm"] || "").toString().trim().toLowerCase();

//     if (lastCommRaw.includes("nevercomm")) {
//       neverComm++;
//     } else {
//       const dt = parseExcelDate(r["LastComm"]);
//       if (dt) {
//         if (dt < threeDaysAgo) {
//           nonCommunicating++;
//         } else {
//           communicating++;
//         }
//       } else {
//         neverComm++; // fallback if invalid date
//       }
//     }

//     // Unmapped
//     if (
//       ((r["Feeder"] || "").toString().trim() === "") &&
//       ((r["DT"] || "").toString().trim() === "") &&
//       ((r["WC"] || "").toString().trim() === "")
//     ) {
//       unmapped++;
//     }
//   });

//   const totalMeters = feeders + dts + meters;

//   const pct = (val) => totalMeters ? (val / totalMeters * 100).toFixed(2) : 0;

//   return {
//     feeders, dts, meters, totalMeters,
//     feedersPct: pct(feeders),
//     dtsPct: pct(dts),
//     metersPct: pct(meters),

//     l1Approved, l2Approved,
//     l1Pct: pct(l1Approved),
//     l2Pct: pct(l2Approved),

//     mdm, sap,
//     mdmPct: pct(mdm),
//     sapPct: pct(sap),

//     sat, satPct: pct(sat),
//     dailyEnergy, dailyEnergyPct: pct(dailyEnergy),

//     communicating, communicatingPct: pct(communicating),
//     nonCommunicating, nonCommunicatingPct: pct(nonCommunicating),
//     neverComm, neverCommPct: pct(neverComm),

//     unmapped, unmappedPct: pct(unmapped)
//   };
// }

// // --- Breakdowns ---
// function computeBreakdowns() {
//   const rows = window.data || [];

//   const commMedium = {
//     GPRS_comm: 0, GPRS_non: 0,
//     RF_comm: 0, RF_non: 0
//   };
//   const meterType = {
//     Feeder_never: 0, DT_never: 0, WC_never: 0
//   };
//   const unmapped = { Feeder: 0, DT: 0, WC: 0 };

//   rows.forEach(r => {
//     const medium = (r["Comm Medium"] || "").toString().trim().toUpperCase();
//     const type = (r["Meter Type"] || "").toString().trim().toUpperCase();

//     const status = classifyComm(r["LastComm"]);

//     if (medium === "GPRS") {
//       if (status === "comm") commMedium.GPRS_comm++;
//       else commMedium.GPRS_non++;
//     }
//     if (medium === "RF") {
//       if (status === "comm") commMedium.RF_comm++;
//       else commMedium.RF_non++;
//     }

//     if (status === "never") {
//       if (type === "FEEDER") meterType.Feeder_never++;
//       if (type === "DT") meterType.DT_never++;
//       if (type === "WC") meterType.WC_never++;
//     }

//     if (!r["Region Name"]) {
//       if (type === "FEEDER") unmapped.Feeder++;
//       if (type === "DT") unmapped.DT++;
//       if (type === "WC") unmapped.WC++;
//     }
//   });

//   return { commMedium, meterType, unmapped };
// }




// --- Render KPI Cards ---
// function renderKpis() {
//   const kpiGrid = document.getElementById("kpiGrid");
//   const totals = computeGlobalTotals(window.data);
//   const breakdowns = computeBreakdowns();

//   const totalMeters = totals.totalMeters;
//   const pct = (num) => totalMeters ? ((num / totalMeters) * 100).toFixed(1) + "%" : "0%";

//   const tiles = [
//     { label: "Total Feeders", value: `${pct(totals.feeders)} (${totals.feeders})`, badge: "Static" },
//     { label: "Total DTs", value: `${pct(totals.dts)} (${totals.dts})`, badge: "Static" },
//     { label: "Total Consumer Meters", value: `${pct(totals.meters)} (${totals.meters})`, badge: "Static" },
//     { label: "Total Meters", value: `${totals.totalMeters}`, badge: "Static" },

//     { label: "L1 / L2 Approved", value: `L1: ${pct(totals.l1Approved)} (${totals.l1Approved}) • L2: ${pct(totals.l2Approved)} (${totals.l2Approved})`, badge: "Quality" },
//     { label: "MDM / SAP", value: `MDM: ${pct(totals.mdm)} (${totals.mdm}) • SAP: ${pct(totals.sap)} (${totals.sap})`, badge: "Integration" },
//     { label: "SAT", value: `${pct(totals.sat)} (${totals.sat})`, badge: "Static" },
//     { label: "Daily Energy", value: `${pct(totals.dailyEnergy)} (${totals.dailyEnergy})`, badge: "Static" },

//     { label: "Communicating", value: `${pct(totals.communicating)} (${totals.communicating})`, badge: "Comm",
//       subtitle: `GPRS: ${breakdowns.commMedium.GPRS_comm}, RF: ${breakdowns.commMedium.RF_comm}` },
//     { label: "Non-Communicating", value: `${pct(totals.nonCommunicating)} (${totals.nonCommunicating})`, badge: "Non-Comm",
//       subtitle: `GPRS: ${breakdowns.commMedium.GPRS_non}, RF: ${breakdowns.commMedium.RF_non}` },
//     { label: "NeverComm", value: `${pct(totals.neverComm)} (${totals.neverComm})`, badge: "NeverComm",
//       subtitle: `Feeder: ${breakdowns.meterType.Feeder_never}, DT: ${breakdowns.meterType.DT_never}, WC: ${breakdowns.meterType.WC_never}` },
//     { label: "Unmapped", value: `${pct(totals.unmapped)} (${totals.unmapped})`, badge: "Unmapped",
//       subtitle: `Feeder: ${breakdowns.unmapped.Feeder}, DT: ${breakdowns.unmapped.DT}, WC: ${breakdowns.unmapped.WC}` }
//   ];

//   kpiGrid.innerHTML = tiles.map(t => `
//     <div class="kpi kpi--${t.badge.toLowerCase()}">
//       <div class="label">${t.label}</div>
//       <div class="value">${t.value}</div>
//       ${t.subtitle ? `<div class="subtitle">${t.subtitle}</div>` : ""}
//       <div class="badge">${t.badge}</div>
//     </div>
//   `).join("");
// }

// kpi.js
const EXCEL_PATH = "assets/meter_sample_filled.xlsx"; // adjust if needed
let kpiData = [];

// --- Helpers ---
// const norm = v => (v ?? "").toString().trim();
const isYes = v => ["yes", "approved"].includes(norm(v).toLowerCase());

// robust field getter (handles multiple possible column names)
function getField(row, candidates) {
  for (const c of candidates) {
    if (Object.prototype.hasOwnProperty.call(row, c) && row[c] != null && String(row[c]).trim() !== "") {
      return row[c];
    }
  }
  return null;
}

// Accept many column name variants
const COLS = {
  meterType: ["Meter Type", "MeterType", "Type"],
  region: ["Region Name", "RegionName", "Region"],
  commMedium: ["Comm Medium", "CommMedium", "Comm"],
  lastComm: ["LastComm", "Last Comm", "Last Communication", "Last Communication Date"],
  firstComm: ["First Comm", "FirstComm", "First Communication"],
  l1: ["L1", "L1Approved", "L1 Approved", "L1 Status"],
  l2: ["L2", "L2Approved", "L2 Approved", "L2 Status"],
  mdm: ["MDM"],
  sap: ["SAP"],
  sat: ["SAT"],
  de: ["Daily Energy", "DE", "DailyEnergy"],
};

// parse Excel serial or common string formats
function parseExcelDate(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") {
    // Excel serial -> JS date
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(value).trim();
  // dd-mm-yyyy or dd/mm/yyyy
  let m = s.match(/^(\d{2})[-/](\d{2})[-/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  // yyyy-mm-dd
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// classify communication
function classifyComm(lastCommRaw) {
  const s = norm(lastCommRaw).toLowerCase();
  if (!s) return "non";
  if (s.includes("never")) return "never";
  const d = parseExcelDate(lastCommRaw);
  if (!d) return "non";
  const diffDays = (Date.now() - d.getTime()) / 86400000;
  return diffDays > 3 ? "non" : "comm";
}

// --- Core compute function ---
// rows: array of objects (sheet_to_json)
function computeKpiTotals(rows) {
  // Installed breakdown (exclude unmapped region)
  let installedFeeder = 0, installedDT = 0, installedWC = 0;

  // approvals / integration / energy / sat
  let l1 = 0, l2 = 0, mdm = 0, sap = 0, sat = 0, de = 0;

  // comm buckets
  let comm = 0, non = 0, never = 0;

  // region counts (regionBreak defines total meters)
  const regionBreak = {};
  const unmappedBreak = { Feeder: 0, DT: 0, WC: 0 };

  // aging counters (30+ days)
  const now = new Date();
  let agingNonComm = 0, agingUnmapped = 0;

  rows.forEach(row => {
    const typeRaw = getField(row, COLS.meterType) ?? "";
    const type = norm(typeRaw).toUpperCase();
    const regionRaw = getField(row, COLS.region);
    const regionNorm = regionRaw == null ? "" : norm(regionRaw);
    const regionLower = regionNorm.toLowerCase();
    // consider unmapped when region column explicitly contains 'unmapped' OR is blank
    const isUnmapped = regionNorm === "" || regionLower.includes("unmapped");

    // region key used for totals (map blanks to "Unmapped")
    const regionKey = isUnmapped ? "Unmapped" : regionNorm || "(Blank)";

    // increment region total
    regionBreak[regionKey] = (regionBreak[regionKey] ?? 0) + 1;

    // installed breakdown excludes Unmapped
    if (!isUnmapped) {
      if (type === "FEEDER") installedFeeder++;
      else if (type === "DT") installedDT++;
      else if (type === "WC") installedWC++;
    } else {
      // unmapped break down by type
      if (type === "FEEDER") unmappedBreak.Feeder++;
      else if (type === "DT") unmappedBreak.DT++;
      else if (type === "WC") unmappedBreak.WC++;
    }

    // approvals / integration / energy / sat
    const vL1 = getField(row, COLS.l1);
    const vL2 = getField(row, COLS.l2);
    const vMdm = getField(row, COLS.mdm);
    const vSap = getField(row, COLS.sap);
    const vSat = getField(row, COLS.sat);
    const vDe = getField(row, COLS.de);

    if (isYes(vL1)) l1++;
    if (isYes(vL2)) l2++;
    if (isYes(vMdm)) mdm++;
    if (isYes(vSap)) sap++;
    if (isYes(vSat)) sat++;
    if (isYes(vDe)) de++;

    // communication
    const lc = getField(row, COLS.lastComm);
    const commClass = classifyComm(lc);
    if (commClass === "comm") comm++;
    else if (commClass === "never") never++;
    else non++;

    // aging checks (30+ days)
    const firstCommVal = getField(row, COLS.firstComm);
    const firstCommDate = parseExcelDate(firstCommVal);
    if (commClass === "non" && firstCommDate) {
      const diffDays = (now - firstCommDate.getTime()) / 86400000;
      if (diffDays > 30) agingNonComm++;
    }
    if (isUnmapped && firstCommDate) {
      const diffDays = (now - firstCommDate.getTime()) / 86400000;
      if (diffDays > 30) agingUnmapped++;
    }
  });

  const totalMeters = Object.values(regionBreak).reduce((a, b) => a + b, 0);
  // unmapped count normalized (case-insensitive)
  const unmapped = (regionBreak["Unmapped"] ?? 0);

  // Installed = totalMeters - unmapped (as requested)
  const installed = Math.max(0, totalMeters - unmapped);

  const pctTotal = (val) => totalMeters ? ((val / totalMeters) * 100).toFixed(1) + "%" : "0%";
  const pctInstalled = (val) => installed ? ((val / installed) * 100).toFixed(1) + "%" : "0%";

  return {
    // installed by type
    installedFeeder, installedDT, installedWC, installed,

    // approvals / integration / energy / sat
    l1, l2, mdm, sap, sat, de,

    // communication
    comm, non, never,

    // totals
    totalMeters, regionBreak, unmapped, unmappedBreak,

    // aging
    agingNonComm, agingUnmapped,

    // helpers
    pctTotal, pctInstalled
  };
}

// --- Render KPI cards ---
function renderKpis() {
  const t = computeKpiTotals(kpiData);

  const gapMdm = t.installed - t.mdm;
  const gapSap = t.installed - t.sap;

  const tiles = [
    {
      label: "Total Meters",
      value: `${t.totalMeters} (${t.pctTotal(t.totalMeters)})`,
      badge: "Total",
      wide: true,
      subtitle: Object.entries(t.regionBreak)
        .map(([r, c]) => `${r}: ${c} (${t.pctTotal(c)})`)
        .join(", ")
    },
    {
      label: "Installed Meters",
      value: `${t.installed} (${t.pctTotal(t.installed)})`,
      badge: "Installed",
      subtitle:
        `Feeder: ${t.installedFeeder} (${t.pctInstalled(t.installedFeeder)}), ` +
        `DT: ${t.installedDT} (${t.pctInstalled(t.installedDT)}), ` +
        `WC: ${t.installedWC} (${t.pctInstalled(t.installedWC)})`
    },
    {
      label: "L1 / L2 Approved",
      value: `L1: ${t.l1} (${t.pctTotal(t.l1)}) • L2: ${t.l2} (${t.pctTotal(t.l2)})`,
      badge: "Quality"
    },
    {
      label: "MDM / SAP",
      value: `MDM: ${t.mdm} (${t.pctTotal(t.mdm)}) • SAP: ${t.sap} (${t.pctTotal(t.sap)})`,
      badge: "Integration"
    },
    {
      label: "Daily Energy",
      value: `${t.de} (${t.pctTotal(t.de)})`,
      badge: "Energy"
    },
    {
      label: "SAT",
      value: `${t.sat} (${t.pctTotal(t.sat)})`,
      badge: "SAT"
    },
    {
      label: "GAP",
      value: `MDM: ${gapMdm} (${t.pctInstalled(gapMdm)}), SAP: ${gapSap} (${t.pctInstalled(gapSap)})`,
      badge: "GAP"
    },
    {
      label: "Communicating",
      value: `${t.comm} (${t.pctTotal(t.comm)})`,
      badge: "Comm"
    },
    {
      label: "Non-Communicating",
      value: `${t.non} (${t.pctTotal(t.non)})`,
      badge: "Non-Comm"
    },
    {
      label: "NeverComm",
      value: `${t.never} (${t.pctTotal(t.never)})`,
      badge: "NeverComm"
    },
    {
      label: "Unmapped",
      value: `${t.unmapped} (${t.pctTotal(t.unmapped)})`,
      badge: "Unmapped",
      subtitle: `Feeder: ${t.unmappedBreak.Feeder}, DT: ${t.unmappedBreak.DT}, WC: ${t.unmappedBreak.WC}`
    },
    {
      label: "Aging (Above 1 Month)",
      value: `Non-Comm: ${t.agingNonComm} (${t.pctTotal(t.agingNonComm)}) • Unmapped: ${t.agingUnmapped} (${t.pctTotal(t.agingUnmapped)})`,
      badge: "Aging"
    }
  ];

  const html = tiles.map(card => `
    <div class="kpi ${card.wide ? "kpi--wide" : ""} kpi--${card.badge.toLowerCase().replace(/\s+/g,'-')}">
      <div class="label">${card.label}</div>
      <div class="value">${card.value}</div>
      ${card.subtitle ? `<div class="subtitle">${card.subtitle}</div>` : ""}
      <div class="badge">${card.badge}</div>
    </div>
  `).join("");

  const el = document.getElementById("kpiGrid");
  if (el) el.innerHTML = html;
}

// --- Load Excel and init ---
// async function loadKpiExcel() {
//   try {
//     const resp = await fetch(EXCEL_PATH);
//     if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
//     const buf = await resp.arrayBuffer();
//     const wb = XLSX.read(buf, { type: "array" });
//     const sheet = wb.Sheets[wb.SheetNames[0]];
//     kpiData = XLSX.utils.sheet_to_json(sheet, { defval: null });
//     renderKpis();
//   } catch (err) {
//     console.error("Failed to load KPI Excel:", err);
//     // leave kpiGrid empty if error
//   }
// }

document.addEventListener("DOMContentLoaded", loadExcel);

document.addEventListener("DOMContentLoaded", () => {
  const toggle = document.getElementById("themeToggle");
  toggle.addEventListener("click", () => {
    const root = document.documentElement;
    const current = root.getAttribute("data-theme");
    root.setAttribute("data-theme", current === "light" ? "dark" : "light");
    localStorage.setItem("theme", root.getAttribute("data-theme"));
  });

  // Persist theme
  const saved = localStorage.getItem("theme") || "dark";
  document.documentElement.setAttribute("data-theme", saved);
});

