// kpi.js
// Place this file in ./assets/kpi.js (or change EXCEL_PATH) and include after XLSX script.
// Requires XLSX global (xlsx.full.min.js) loaded before this script.

const EXCEL_PATH = "meter_sample_filled.xlsx"; // adjust if needed
let kpiData = [];

// --- Helpers ---
const norm = v => (v ?? "").toString().trim();
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
      value: `Installed–MDM: ${gapMdm} (${t.pctInstalled(gapMdm)}), Installed–SAP: ${gapSap} (${t.pctInstalled(gapSap)})`,
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
      label: "Aging (>30d)",
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
async function loadKpiExcel() {
  try {
    const resp = await fetch(EXCEL_PATH);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    kpiData = XLSX.utils.sheet_to_json(sheet, { defval: null });
    renderKpis();
  } catch (err) {
    console.error("Failed to load KPI Excel:", err);
    // leave kpiGrid empty if error
  }
}

document.addEventListener("DOMContentLoaded", loadKpiExcel);
