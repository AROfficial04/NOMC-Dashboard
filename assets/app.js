/*
  Framework-free dashboard logic
  - Parses Excel via SheetJS (XLSX) from CDN
  - Holds data entirely in-memory
  - Renders KPIs, region table, explorer tree, and detail views
*/

// Threshold helpers
const LOSS_THRESHOLDS = { greenMax: 2, amberMax: 5 }; // percent
// percent

function classifyLoss(pct) {
  if (pct == null || isNaN(pct)) return { cls: "", label: "—" };
  if (pct < LOSS_THRESHOLDS.greenMax) return { cls: "badge-green", label: "Green" };
  if (pct <= LOSS_THRESHOLDS.amberMax) return { cls: "badge-amber", label: "Amber" };
  return { cls: "badge-red", label: "Red" };
}


// State
const state = {
  raw: null,
  regions: [], // [{ name, feeders:[{...}], metrics }]
  lookups: {
    regionByName: new Map(),
    feederById: new Map(),
    dtById: new Map(),
    meterById: new Map(),
  },
  currentSelection: { level: "ALL", id: null },
  sort: { key: "region", dir: "asc" },
  resultsRows: [],
  slaGlobal: { totalRows: 0, dailyYes: 0, loadYes: 0 },
};

// Sample schema note:
// Expect an Excel workbook with sheets providing Region, Feeder, DT, Meter relationships and readings.
// For demo, if no file uploaded, generate mock data to showcase UI.

document.addEventListener("DOMContentLoaded", () => {
  const globalSearch = document.getElementById("globalSearch");
  const regionFilter = document.getElementById("regionFilter");
  
  const accordionRegionFilter = document.getElementById("accordionRegionFilter");

  globalSearch.addEventListener("keydown", (e) => {
    if (e.key === "Enter") onGlobalSearch(e.target.value.trim());
  });
  regionFilter.addEventListener("change", renderAll);
  
  accordionRegionFilter.addEventListener("change", renderAccordionTable);

  // Auto-load workbook from hardcoded path, fallback to mock
  autoLoadWorkbook();
});

async function autoLoadWorkbook() {
  const path = "sample_3sheets_clean.xlsx";
  try {
    const resp = await fetch(path);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buf = await resp.arrayBuffer();
    const data = new Uint8Array(buf);
    const workbook = XLSX.read(data, { type: "array" });
    document.getElementById("fileName").textContent = path;
    document.getElementById("processedTime").textContent = new Date().toLocaleString();
    parseWorkbook(workbook);
  } catch (err) {
    console.warn("Failed to load Excel from hardcoded path, using mock data.", err);
    bootstrapWithMockData();
  }
}

function parseWorkbook(workbook) {
  const feeders = XLSX.utils.sheet_to_json(workbook.Sheets["Feeders"], { defval: null });
  const dts     = XLSX.utils.sheet_to_json(workbook.Sheets["DTs"], { defval: null });
  const meters  = XLSX.utils.sheet_to_json(workbook.Sheets["Meters"], { defval: null });

  buildModelFromThreeSheets(feeders, dts, meters);
  renderAll();
}


function buildModelFromThreeSheets(feedersSheet, dtsSheet, metersSheet) {
  // Reset state
  state.regions = [];
  state.lookups.regionByName.clear();
  state.lookups.feederById.clear();
  state.lookups.dtById.clear();
  state.lookups.meterById.clear();
  state.slaGlobal = { totalRows: 0, commYes: 0, nonCommYes: 0, unmappedYes: 0, neverCommYes: 0 };

  // --- Feeders ---
  for (const f of feedersSheet) {
    const regionName = f["Region Name"];
    const feederId   = f["Feeder Code"];
    const feederName = f["Feeder Name"];
    const fDay1      = Number(f["Feeder Day1 reading"]);
    const fDay2      = Number(f["Feeder Day2 reading"]);
    const mf         = Number(f["MF Feeder"]) || 1;
    const comm       = parseYesNo(f["Comm"]);
    const nonComm    = parseYesNo(f["NonComm"]);

    // track SLA global for feeders
    state.slaGlobal.totalRows += 1;
    if (comm) state.slaGlobal.commYes += 1;
    if (nonComm) state.slaGlobal.nonCommYes += 1;

    // region setup
    let region = state.lookups.regionByName.get(regionName);
    if (!region) {
      region = { name: regionName, feeders: [], metrics: {} };
      state.lookups.regionByName.set(regionName, region);
      state.regions.push(region);
    }

    const feederEnergy = diffEnergy(fDay1, fDay2, mf);
    const feederObj = {
      id: feederId, name: feederName,
      dts: [], meters: [],
      metrics: { feederEnergy, comm, nonComm },
      region: regionName
    };

    state.lookups.feederById.set(feederId, feederObj);
    region.feeders.push(feederObj);
  }

  // --- DTs ---
  for (const d of dtsSheet) {
    const regionName = d["Region Name"];
    const feederId   = d["Feeder Code"];
    const dtId       = d["DT Code"];
    const dtName     = d["DT Name"];
    const dtDay1     = Number(d["DT Day1 Reading"]);
    const dtDay2     = Number(d["DT Day2 Reading"]);
    const mf         = Number(d["MF DT"]) || 1;
    const comm       = parseYesNo(d["Comm"]);
    const nonComm    = parseYesNo(d["NonComm"]);
    const unmapped   = parseYesNo(d["Unmapped"]);
    const neverComm  = parseYesNo(d["NeverComm"]);

    // track SLA global for DTs
    state.slaGlobal.totalRows += 1;
    if (comm) state.slaGlobal.commYes += 1;
    if (nonComm) state.slaGlobal.nonCommYes += 1;
    if (unmapped) state.slaGlobal.unmappedYes += 1;
    if (neverComm) state.slaGlobal.neverCommYes += 1;

    const dtEnergy = diffEnergy(dtDay1, dtDay2, mf);
    const dtObj = {
      id: dtId, name: dtName, meters: [],
      metrics: { dtEnergy, comm, nonComm, unmapped, neverComm },
      feederId
    };

    state.lookups.dtById.set(dtId, dtObj);
    const feederObj = state.lookups.feederById.get(feederId);
    if (feederObj) feederObj.dts.push(dtObj);
  }

  // --- Meters ---
  for (const m of metersSheet) {
    const meterId  = m["Meter No."];
    const feederId = m["Feeder Code"];
    const dtId     = m["DT Code"];
    const mDay1    = Number(m["Meter Day1 Reading"]);
    const mDay2    = Number(m["Meter Day2 Reading"]);
    const comm     = parseYesNo(m["Comm"]);
    const nonComm  = parseYesNo(m["NonComm"]);
    const unmapped = parseYesNo(m["Unmapped"]);
    const neverComm= parseYesNo(m["NeverComm"]);

    const energy   = diffEnergy(mDay1, mDay2, 1);

    // track SLA global
    state.slaGlobal.totalRows += 1;
    if (comm) state.slaGlobal.commYes += 1;
    if (nonComm) state.slaGlobal.nonCommYes += 1;
    if (unmapped) state.slaGlobal.unmappedYes += 1;
    if (neverComm) state.slaGlobal.neverCommYes += 1;

    const meterObj = { id: meterId, readings: { day1: mDay1, day2: mDay2 }, energy, comm, nonComm, unmapped, neverComm, dtId, feederId };
    state.lookups.meterById.set(meterId, meterObj);

    // If not unmapped, link into hierarchy
    if (!unmapped) {
      const dtObj = state.lookups.dtById.get(dtId);
      if (dtObj) dtObj.meters.push(meterObj);
      const feederObj = state.lookups.feederById.get(feederId);
      if (feederObj) feederObj.meters.push(meterObj);
    }
  }

  // Aggregate metrics per region
  // Aggregate metrics per region
  for (const region of state.regions) {
    region.metrics.feeders = region.feeders.length;
    region.metrics.dts = region.feeders.reduce((a,f) => a+f.dts.length, 0);
    region.metrics.meters = region.feeders.reduce((a,f) => a+f.meters.length, 0);

    // Calculate communication metrics from all asset types (feeders, DTs, meters)
    region.metrics.communicating = region.feeders
      .reduce((acc, f) => acc + (f.metrics.comm ? 1 : 0) + f.dts.reduce((dtAcc, dt) => dtAcc + (dt.metrics.comm ? 1 : 0), 0) + f.meters.filter(m => m.comm).length, 0);
    region.metrics.nonCommunicating = region.feeders
      .reduce((acc, f) => acc + (f.metrics.nonComm ? 1 : 0) + f.dts.reduce((dtAcc, dt) => dtAcc + (dt.metrics.nonComm ? 1 : 0), 0) + f.meters.filter(m => m.nonComm).length, 0);
    region.metrics.unmapped = region.feeders
      .reduce((acc, f) => acc + f.dts.reduce((dtAcc, dt) => dtAcc + (dt.metrics.unmapped ? 1 : 0), 0) + f.meters.filter(m => m.unmapped).length, 0);
    region.metrics.neverComm = region.feeders
      .reduce((acc, f) => acc + f.dts.reduce((dtAcc, dt) => dtAcc + (dt.metrics.neverComm ? 1 : 0), 0) + f.meters.filter(m => m.neverComm).length, 0);
  }

}


function renderAll() {
  // renderKpis();
  renderRegionFilter();
  // renderRegionTable();
  renderTree();
  renderDetail();
  renderAccordionFilters();
  renderAccordionTable();
}

// function renderKpis() {
//   const kpiGrid = document.getElementById("kpiGrid");
//   const totals = computeGlobalTotals();
//   const tiles = [
//     { label: "Total Feeders", value: totals.feeders, hint: "Total feeders", badge: "Static" },
//     { label: "Total L1 Approved", value: totals.l1Approved, hint: "Total L1 Approved", badge: "Static" },
//     { label: "Total L2 Approved", value: totals.l2Approved, hint: "Total L2 Approved", badge: "Static" },
//     { label: "Total MDM", value: totals.mdm, hint: "Total MDM", badge: "Static" },
//     { label: "Total SAP", value: totals.sap, hint: "Total SAP", badge: "Static" },
//     { label: "Total SAT", value: totals.sat, hint: "Total SAT", badge: "Static" },
//     { label: "Total Daily Energy", value: totals.dailyEnergy, hint: "Total Daily Energy", badge: "Static" },
//     { label: "Total DTs", value: totals.dts, hint: "Total DTs", badge: "Static" },
//     { label: "Total Consumer Meters", value: totals.meters, hint: "Total consumer meters", badge: "Static" },
//     { label: "Total Meters", value: totals.feeders + totals.dts + totals.meters, hint: "Total meters (Feeders + DTs + Consumer Meters)", badge: "Static" },
//     { label: "Communicating", value: `${fmtPct(totals.communicatingPct)} (${totals.communicating})`, hint: "Communicating meters percentage and count", badge: "Comm" },
//     { label: "Non-Communicating", value: `${fmtPct(totals.nonCommunicatingPct)} (${totals.nonCommunicating})`, hint: "Non-Communicating meters percentage and count", badge: "Non-Comm" },
//     { label: "Unmapped", value: `${fmtPct(totals.unmappedPct)} (${totals.unmapped})`, hint: "Unmapped meters percentage and count", badge: "Unmapped" },
//     { label: "NeverComm", value: `${fmtPct(totals.neverCommPct)} (${totals.neverComm})`, hint: "Never Communicating meters percentage and count", badge: "NeverComm" },
//   ];

//   kpiGrid.innerHTML = tiles.map((t) => `
//     <div class="kpi" title="${escapeHtml(t.hint)}">
//       <div class="label">${t.label}</div>
//       <div class="value ${t.className ?? ''}">${t.value}</div>
//       <div class="badge">${t.badge}</div>
//       <div class="hint">${t.hint}</div>
//     </div>
//   `).join("");
// }

function renderPerAssetLossList() {
  const container = document.getElementById("perAssetLossList");
  const feeders = Array.from(state.lookups.feederById.values());
  const html = feeders.map((f) => {
    const fE = f.metrics.feederEnergy ?? 0;
    const sumDt = f.dts.reduce((acc, d) => acc + (d.metrics.dtEnergy ?? 0), 0);
    const fLoss = fE ? Number((((fE - sumDt) / fE) * 100).toFixed(2)) : null;
    const fHeader = `<div><strong>${escapeHtml(f.id)}</strong> – Feeder→DT Loss % = ${fmtPctOrDash(fLoss)}</div>`;
    const dtLines = f.dts.map((d) => {
      const dE = d.metrics.dtEnergy ?? 0;
      const sumCons = d.meters.reduce((acc, m) => acc + (m.energy ?? 0), 0);
      const dLoss = dE ? Number((((dE - sumCons) / dE) * 100).toFixed(2)) : null;
      return `<div style="margin-left:16px;">${escapeHtml(d.id)} – DT→Consumer Loss % = ${fmtPctOrDash(dLoss)}</div>`;
    }).join("");
    return fHeader + dtLines;
  }).join("");
  container.innerHTML = html || '<div class="pad">No data.</div>';
}

function computeGlobalTotals() {
  const feeders = state.lookups.feederById.size;
  const dts = state.lookups.dtById.size;
  const meters = state.lookups.meterById.size;

  const feederEnergy = sum(Array.from(state.lookups.feederById.values()).map(f => f.metrics.feederEnergy || 0));
  const dtEnergy = sum(Array.from(state.lookups.dtById.values()).map(d => d.metrics.dtEnergy || 0));
  const consumerEnergy = sum(Array.from(state.lookups.meterById.values()).map(m => m.energy || 0));

  const lossFdt = pctLoss(feederEnergy, dtEnergy);
  const lossDtc = pctLoss(dtEnergy, consumerEnergy);
  const lossFc  = pctLoss(feederEnergy, consumerEnergy);

  const totalRows = state.slaGlobal.totalRows;
  const communicatingPct = pct(state.slaGlobal.commYes, totalRows);
  const nonCommunicatingPct = pct(state.slaGlobal.nonCommYes, totalRows);
  const unmappedPct = pct(state.slaGlobal.unmappedYes, totalRows);
  const neverCommPct = pct(state.slaGlobal.neverCommYes, totalRows);

  return {
    feeders, dts, meters,
    lossFdt, lossDtc, lossFc,
    communicating: state.slaGlobal.commYes,
    communicatingPct,
    nonCommunicating: state.slaGlobal.nonCommYes,
    nonCommunicatingPct,
    unmapped: state.slaGlobal.unmappedYes,
    unmappedPct,
    neverComm: state.slaGlobal.neverCommYes,
    neverCommPct
  };
}


function renderRegionFilter() {
  const select = document.getElementById("regionFilter");
  const prev = select.value;
  const options = ["__ALL__", ...state.regions.map((r) => r.name)];
  select.innerHTML = options.map((opt) => `<option value="${escapeHtml(opt)}">${opt === "__ALL__" ? "All Regions" : escapeHtml(opt)}</option>`).join("");
  if (options.includes(prev)) select.value = prev;
}

function renderRegionTable() {
  const tbody = document.querySelector("#regionTable tbody");
  const regionFilter = document.getElementById("regionFilter").value;


  let regions = state.regions.map((r) => ({
    region: r.name,
    feeders: r.metrics.feeders,
    dts: r.metrics.dts,
    meters: r.metrics.meters,
    comm: r.metrics.communicating,
    nonComm: r.metrics.nonCommunicating,
    neverComm: r.metrics.neverComm,

  }));

  if (regionFilter !== "__ALL__") regions = regions.filter((x) => x.region === regionFilter);

  regions.sort((a, b) => {
    const { key, dir } = state.sort;
    const av = a[key];
    const bv = b[key];
    if (typeof av === "string" && typeof bv === "string") return dir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    return dir === "asc" ? (av - bv) : (bv - av);
  });

  // tbody.innerHTML = regions.map((r) => `
  //   <tr>
  //     <td><button class="link" data-action="open-region" data-region="${escapeHtml(r.region)}">${escapeHtml(r.region)}</button></td>
  //     <td>${r.feeders}</td>
  //             <td>${r.dts}</td>
  //       <td>${r.meters}</td>
  //       <td>${r.comm}</td>
  //       <td>${r.nonComm}</td>
  //       <td>${r.neverComm}</td>
  //   </tr>
  // `).join("");

  // header sort interactions
  const headers = document.querySelectorAll("#regionTable thead th");
  headers.forEach((th) => {
    th.onclick = () => {
      const key = th.dataset.key;
      if (!key) return;
      if (state.sort.key === key) state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
      else state.sort = { key, dir: "asc" };
      renderRegionTable();
    };
  });

  // open region handlers
  tbody.querySelectorAll("[data-action='open-region']").forEach((btn) => {
    btn.addEventListener("click", () => {
      const name = btn.getAttribute("data-region");
      state.currentSelection = { level: "REGION", id: name };
      renderDetail();
    });
  });
}

function renderTree() {
  const tree = document.getElementById("tree");
  const regionFilter = document.getElementById("regionFilter").value;
  const regions = state.regions.filter((r) => regionFilter === "__ALL__" || r.name === regionFilter);
  tree.innerHTML = regions.map((r) => `
    <div class="node" data-type="region" data-id="${escapeHtml(r.name)}">
      <span class="icon">▸</span>
      <span class="label">${escapeHtml(r.name)}</span>
      <span class="meta">${r.metrics.feeders}F • ${r.metrics.dts}DT • ${r.metrics.meters}M</span>
    </div>
    <div class="children">
      ${r.feeders.map((f) => `
        <div class="node" data-type="feeder" data-id="${escapeHtml(f.id)}">
          <span class="icon">▸</span>
          <span class="label">Feeder ${escapeHtml(f.name)}</span>
          <span class="meta">${f.metrics.dts ?? f.dts.length}DT • ${f.metrics.meters ?? f.meters.length}M</span>
        </div>
        <div class="children">
          ${f.dts.map((d) => `
            <div class="node" data-type="dt" data-id="${escapeHtml(d.id)}">
              <span class="icon">▸</span>
              <span class="label">DT ${escapeHtml(d.name)}</span>
              <span class="meta">${d.meters.length}M</span>
            </div>
            <div class="children">
              ${d.meters.map((m) => `
                <div class="node" data-type="meter" data-id="${escapeHtml(m.id)}">
                  <span class="icon">●</span>
                  <span class="label">Meter ${escapeHtml(m.id)}</span>
                  <span class="meta">${m.energy ?? 0} kWh</span>

                </div>
              `).join("")}
            </div>
          `).join("")}
        </div>
      `).join("")}
    </div>
  `).join("");

  // expand/collapse and select
  tree.querySelectorAll(".node").forEach((node) => {
    node.addEventListener("click", (e) => {
      const type = node.getAttribute("data-type");
      const id = node.getAttribute("data-id");
      const icon = node.querySelector(".icon");
      const next = node.nextElementSibling;
      if (next && next.classList.contains("children")) {
        node.classList.toggle("open");
        icon.textContent = node.classList.contains("open") ? "▾" : "▸";
      }
      state.currentSelection = { level: type.toUpperCase(), id };
      renderDetail();
      e.stopPropagation();
    });
  });
}

function renderDetail() {
  const title = document.getElementById("detailTitle");
  const metricsWrap = document.getElementById("detailMetrics");
  const body = document.getElementById("detailBody");
  const sel = state.currentSelection;

  // if (sel.level === "ALL") {
  //   title.textContent = "Overall";
  //   const totals = computeGlobalTotals();
  //   metricsWrap.innerHTML = renderMetricGrid([
  //     ["Feeders", totals.feeders],
  //     ["DTs", totals.dts],
  //     ["Meters", totals.meters],
  //     ["F→DT Loss%", fmtPct(totals.lossFdt)],
  //     ["DT→Cons Loss%", fmtPct(totals.lossDtc)],
  //     ["F→Cons Loss%", fmtPct(totals.lossFc)],
     
  //   ]);
  //   body.innerHTML = "<div class='pad'>Use the explorer or table to drill down.</div>";
  //   return;
  // }

  if (sel.level === "REGION") {
    const region = state.lookups.regionByName.get(sel.id);
    if (!region) return;
    title.textContent = `Region: ${region.name}`;
    const m = region.metrics;
    metricsWrap.innerHTML = renderMetricGrid([
      ["Feeders", m.feeders],
      ["DTs", m.dts],
      ["Meters", m.meters],
      ["F→DT Loss%", fmtPct(m.lossFdt)],
      ["DT→Cons Loss%", fmtPct(m.lossDtc)],
      ["F→Cons Loss%", fmtPct(m.lossFc)],
     
    ]);
    body.innerHTML = renderFeederTable(region.feeders);
    attachFeederRowHandlers();
    return;
  }

  if (sel.level === "FEEDER") {
    const feeder = state.lookups.feederById.get(sel.id);
    if (!feeder) return;
    title.textContent = `Feeder: ${feeder.name}`;
    const m = feeder.metrics;
    metricsWrap.innerHTML = renderMetricGrid([
      ["DTs", m.dts],
      ["Meters", m.meters],
      ["F→DT Loss%", fmtPct(m.lossFdt)],
      ["DT→Cons Loss%", fmtPct(m.lossDtc)],
      ["F→Cons Loss%", fmtPct(m.lossFc)],
     
    ]);
    body.innerHTML = renderDtTable(feeder.dts);
    attachDtRowHandlers();
    return;
  }

  if (sel.level === "DT") {
    const dt = state.lookups.dtById.get(sel.id);
    if (!dt) return;
    title.textContent = `DT: ${dt.name}`;
    const m = dt.metrics;
    metricsWrap.innerHTML = renderMetricGrid([
      ["Meters", m.meters],
      ["DT→Cons Loss%", fmtPct(m.lossDtc)],
      
    ]);
    body.innerHTML = renderMeterTable(dt.meters);
    return;
  }

  if (sel.level === "METER") {
    const meter = state.lookups.meterById.get(sel.id);
    if (!meter) return;
    title.textContent = `Meter: ${meter.id}`;
         metricsWrap.innerHTML = renderMetricGrid([
       ["Reading Day1", meter.readings.day1 ?? "—"],
       ["Reading Day2", meter.readings.day2 ?? "—"],
       ["Energy", meter.energy ?? 0],
      
       ["Comm", meter.comm ? "Yes" : "No"],
       ["NonComm", meter.nonComm ? "Yes" : "No"],
       ["Unmapped", meter.unmapped ? "Yes" : "No"],
       ["NeverComm", meter.neverComm ? "Yes" : "No"],
       

     ]);
    body.innerHTML = "";
    return;
  }
}

function onGlobalSearch(query) {
  if (!query) return;
  // try meter
  if (state.lookups.meterById.has(query)) {
    state.currentSelection = { level: "METER", id: query };
    renderDetail();
    return;
  }
  if (state.lookups.dtById.has(query)) {
    state.currentSelection = { level: "DT", id: query };
    renderDetail();
    return;
  }
  if (state.lookups.feederById.has(query)) {
    state.currentSelection = { level: "FEEDER", id: query };
    renderDetail();
    return;
  }
  if (state.lookups.regionByName.has(query)) {
    state.currentSelection = { level: "REGION", id: query };
    renderDetail();
    return;
  }
  alert("No match found. Use exact Region/Feeder/DT/Meter identifier.");
}

function buildResultsTable() {
  const uniqueFeeders = Array.from(state.lookups.feederById.values());
  const uniqueDts = Array.from(state.lookups.dtById.values());
  const uniqueMeters = Array.from(state.lookups.meterById.values());

  // Cons_E per meter (dedup + non-negative rule)
  const meterEnergyById = new Map();
  for (const m of uniqueMeters) {
    const energy = m.energy;
    if (energy == null) continue;
    if (energy < 0) continue; // flag for review: we exclude negatives from sums
    meterEnergyById.set(m.id, energy);
  }

  // DT_E per DT (dedup + non-negative)
  const dtEnergyById = new Map();
  for (const d of uniqueDts) {
    const e = d.metrics.dtEnergy;
    if (e == null) continue;
    if (e < 0) continue;
    dtEnergyById.set(d.id, e);
  }

  // Feeder_E per feeder (dedup + non-negative)
  const feederEnergyById = new Map();
  for (const f of uniqueFeeders) {
    const e = f.metrics.feederEnergy;
    if (e == null) continue;
    if (e < 0) continue;
    feederEnergyById.set(f.id, e);
  }

  // Aggregations
  const sumConsByDt = new Map();
  for (const m of uniqueMeters) {
    const e = meterEnergyById.get(m.id) ?? 0;
    sumConsByDt.set(m.dtId, (sumConsByDt.get(m.dtId) ?? 0) + e);
  }

  const sumDtByFeeder = new Map();
  for (const d of uniqueDts) {
    const e = dtEnergyById.get(d.id) ?? 0;
    sumDtByFeeder.set(d.feederId, (sumDtByFeeder.get(d.feederId) ?? 0) + e);
  }

  const sumConsByFeeder = new Map();
  for (const m of uniqueMeters) {
    const e = meterEnergyById.get(m.id) ?? 0;
    sumConsByFeeder.set(m.feederId, (sumConsByFeeder.get(m.feederId) ?? 0) + e);
  }

  // Build rows
  const rows = [];

  // Rows at feeder-level paired with blank DT columns
  for (const f of uniqueFeeders) {
    const Feeder_Code = f.id;
    const Feeder_E = feederEnergyById.get(Feeder_Code) ?? null;
    const Sum_DT_E = sumDtByFeeder.get(Feeder_Code) ?? 0;
    const Sum_Cons_E = sumConsByFeeder.get(Feeder_Code) ?? 0;
    const Feeder_to_DT_Loss = lossOrNull(Feeder_E, Sum_DT_E);
    const Feeder_to_Cons_Loss = lossOrNull(Feeder_E, Sum_Cons_E);

    rows.push({
      Feeder_Code,
      Feeder_E,
      Sum_DT_E,
      Feeder_to_DT_Loss,
      Sum_Cons_E,
      Feeder_to_Cons_Loss,
      DT_Code: "",
      DT_E: null,
      Sum_Cons_E_for_DT: null,
      DT_to_Cons_Loss: null,
    });
  }

  // Rows at DT-level
  for (const d of uniqueDts) {
    const DT_Code = d.id;
    const DT_E = dtEnergyById.get(DT_Code) ?? null;
    const Sum_Cons_E_for_DT = sumConsByDt.get(DT_Code) ?? 0;
    const DT_to_Cons_Loss = lossOrNull(DT_E, Sum_Cons_E_for_DT);
    rows.push({
      Feeder_Code: d.feederId,
      Feeder_E: null,
      Sum_DT_E: null,
      Feeder_to_DT_Loss: null,
      Sum_Cons_E: null,
      Feeder_to_Cons_Loss: null,
      DT_Code,
      DT_E,
      Sum_Cons_E_for_DT,
      DT_to_Cons_Loss,
    });
  }

  state.resultsRows = rows;
}

function lossOrNull(input, comparedSum) {
  if (input == null || input === 0) return null;
  return Number((((input - (comparedSum ?? 0)) / input) * 100).toFixed(2));
}

function renderResultsTable() {
  // deprecated table removed in favor of accordion
}

function renderAccordionFilters() {
  const select = document.getElementById("accordionRegionFilter");
  if (!select) return;
  const prev = select.value;
  const options = ["__ALL__", ...state.regions.map((r) => r.name)];
  select.innerHTML = options.map((opt) => `<option value="${escapeHtml(opt)}">${opt === "__ALL__" ? "All Regions" : escapeHtml(opt)}</option>`).join("");
  if (options.includes(prev)) select.value = prev;
}

function renderAccordionTable() {
  const tbody = document.querySelector('#accordionTable tbody');
  if (!tbody) return;
  const regionSel = document.getElementById('accordionRegionFilter').value;
  const regions = state.regions.filter((r) => regionSel === '__ALL__' || r.name === regionSel);

  // Build rows with data attributes for expand/collapse
  const rows = [];
  // Store rows globally for event listener access
  window.accordionRows = rows;
  for (const region of regions) {
    for (const feeder of region.feeders) {
      const feederEnergy = feeder.metrics.feederEnergy ?? 0;
      const sumDt = feeder.dts.reduce((a, d) => a + (d.metrics.dtEnergy ?? 0), 0);
      const sumConsFeeder = feeder.meters.reduce((a, m) => a + (m.energy ?? 0), 0);
      const f2dtLoss = feederEnergy ? Number((((feederEnergy - sumDt) / feederEnergy) * 100).toFixed(2)) : null;
      const f2consLoss = feederEnergy ? Number((((feederEnergy - sumConsFeeder) / feederEnergy) * 100).toFixed(2)) : null;

      rows.push({
        type: 'feeder', id: feeder.id, parentId: region.name,
        name: feeder.name,
        f2dt: f2dtLoss,
        f2cons: f2consLoss,
        dt2cons: null, // Not applicable for feeders
      });

      for (const dt of feeder.dts) {
        const dtEnergy = dt.metrics.dtEnergy ?? 0;
        const sumConsDt = dt.meters.reduce((a, m) => a + (m.energy ?? 0), 0);
        const dt2ConsLoss = dtEnergy ? Number((((dtEnergy - sumConsDt) / dtEnergy) * 100).toFixed(2)) : null;
        rows.push({
          type: 'dt', id: dt.id, parentId: feeder.id,
          name: dt.name,
          f2dt: null, // Not applicable for DTs
          f2cons: null, // Not applicable for DTs
          dt2cons: dt2ConsLoss,
        });

        for (const meter of dt.meters) {
          const consEnergy = meter.energy ?? diffEnergy(meter.readings.day1, meter.readings.day2, 1) ?? 0;
          rows.push({
            type: 'meter', id: meter.id, parentId: dt.id,
            name: meter.id,
            f2dt: null, // Not applicable for meters
            f2cons: null, // Not applicable for meters
            dt2cons: null, // Not applicable for meters
          });
        }
      }
    }
  }

  // initial render with only feeders visible
  tbody.innerHTML = rows.filter(r => r.type === 'feeder').map(r => rowHtml(r, 0, true)).join('');

  // Use event delegation to handle all row clicks
  // Remove any existing event listeners by cloning the tbody
  const newTbody = tbody.cloneNode(true);
  tbody.parentNode.replaceChild(newTbody, tbody);
  
  // Add the event listener to the new tbody
  newTbody.addEventListener('click', (e) => {
    const tr = e.target.closest('tr.row-toggle');
    if (!tr) return;
    
    const type = tr.getAttribute('data-type');
    if (type === 'feeder' || type === 'dt') {
      e.stopPropagation();
      toggleExpand(tr, window.accordionRows);
    }
  });
}

function rowHtml(r, level, collapsible) {
  const indentCls = level === 1 ? 'indent-1' : level >= 2 ? 'indent-2' : '';
  const icon = collapsible ? '<span class="icon">▸</span>' : '<span class="icon">•</span>';
  if (r.type === 'feeder') {
    return `
      <tr class="row-toggle ${indentCls}" data-type="feeder" data-id="${escapeHtml(r.id)}" data-parent="${escapeHtml(r.parentId)}">
        <td>${icon} ${escapeHtml(r.name)}</td>
        <td>${fmtPctOrDash(r.f2dt)}</td>
        <td>${fmtPctOrDash(r.f2cons)}</td>
        <td>${fmtPctOrDash(r.dt2cons)}</td>
      </tr>
    `;
  }
  if (r.type === 'dt') {
    return `
      <tr class="row-toggle ${indentCls}" data-type="dt" data-id="${escapeHtml(r.id)}" data-parent="${escapeHtml(r.parentId)}">
        <td>${icon} ${escapeHtml(r.name)}</td>
        <td>${fmtPctOrDash(r.f2dt)}</td>
        <td>${fmtPctOrDash(r.f2cons)}</td>
        <td>${fmtPctOrDash(r.dt2cons)}</td>
      </tr>
    `;
  }
  // meter
  return `
    <tr class="row-toggle ${indentCls}" data-type="meter" data-id="${escapeHtml(r.id)}" data-parent="${escapeHtml(r.parentId)}">
      <td>${icon} ${escapeHtml(r.name)}</td>
      <td>${fmtPctOrDash(r.f2dt)}</td>
      <td>${fmtPctOrDash(r.f2cons)}</td>
      <td>${fmtPctOrDash(r.dt2cons)}</td>
    </tr>
  `;
}

function toggleExpand(tr, rows) {
  const type = tr.getAttribute('data-type');
  const id = tr.getAttribute('data-id');
  const nextLevel = type === 'feeder' ? 1 : type === 'dt' ? 2 : 3;
  const tbody = tr.parentElement;
  const iconEl = tr.querySelector('.icon');
  const isOpen = tr.classList.contains('open');

  if (isOpen) {
    // collapse: remove all descendant rows
    const toRemove = [];
    let sibling = tr.nextElementSibling;
    while (sibling && sibling.getAttribute('data-parent') === id) {
      toRemove.push(sibling);
      sibling = sibling.nextElementSibling;
    }
    toRemove.forEach((el) => el.remove());
    tr.classList.remove('open');
    iconEl.textContent = '▸';
    return;
  }

  // expand children
  const children = rows.filter(r => r.parentId === id && ((type === 'feeder' && r.type === 'dt') || (type === 'dt' && r.type === 'meter')));
  const html = children.map(r => rowHtml(r, nextLevel, r.type !== 'meter')).join('');
  tr.insertAdjacentHTML('afterend', html);
  tr.classList.add('open');
  iconEl.textContent = '▾';

  // No need to attach individual handlers since we use event delegation
}

// Render helpers
function renderBadgePct(value, classifier) {
  const cls = classifier(value).cls;
  return `<span class="cell-badge ${cls}">${fmtPct(value)}</span>`;
}

function renderMetricGrid(pairs) {
  return pairs.map(([label, value]) => `
    <div class="metric">
      <div class="label">${label}</div>
      <div class="value">${value}</div>
    </div>
  `).join("");
}

function renderFeederTable(feeders) {
  return `
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr>
          <th>Feeder</th>
          <th>DTs</th>
          <th>Meters</th>
          <th>F→DT Loss%</th>
          <th>DT→Cons Loss%</th>
          <th>F→Cons Loss%</th>
          
        </tr>
      </thead>
      <tbody>
        ${feeders.map((f) => `
          <tr>
            <td><button class="link" data-action="open-feeder" data-id="${escapeHtml(f.id)}">${escapeHtml(f.name)}</button></td>
            <td>${f.metrics.dts}</td>
            <td>${f.metrics.meters}</td>
            <td>${renderBadgePct(f.metrics.lossFdt, classifyLoss)}</td>
            <td>${renderBadgePct(f.metrics.lossDtc, classifyLoss)}</td>
            <td>${renderBadgePct(f.metrics.lossFc, classifyLoss)}</td>
            
          </tr>
        `).join("")}
      </tbody>
    </table>
  </div>`;
}

function attachFeederRowHandlers() {
  document.querySelectorAll("[data-action='open-feeder']").forEach((el) => {
    el.addEventListener("click", () => {
      const id = el.getAttribute("data-id");
      state.currentSelection = { level: "FEEDER", id };
      renderDetail();
    });
  });
}

function renderDtTable(dts) {
  return `
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr>
          <th>DT</th>
          <th>Meters</th>
          <th>DT→Cons Loss%</th>
          
        </tr>
      </thead>
      <tbody>
        ${dts.map((d) => `
          <tr>
            <td><button class="link" data-action="open-dt" data-id="${escapeHtml(d.id)}">${escapeHtml(d.name)}</button></td>
            <td>${d.meters.length}</td>
            <td>${renderBadgePct(d.metrics.lossDtc, classifyLoss)}</td>
           
          </tr>
        `).join("")}
      </tbody>
    </table>
  </div>`;
}

function attachDtRowHandlers() {
  document.querySelectorAll("[data-action='open-dt']").forEach((el) => {
    el.addEventListener("click", () => {
      const id = el.getAttribute("data-id");
      state.currentSelection = { level: "DT", id };
      renderDetail();
    });
  });
}

function renderMeterTable(meters) {
  return `
  <div class="table-wrap">
    <table class="table">
      <thead>
        <tr>
          <th>Meter</th>
          <th>Day1</th>
          <th>Day2</th>
          <th>Energy</th>
         
        </tr>
      </thead>
      <tbody>
        ${meters.map((m) => `
          <tr>
            <td><button class="link" data-action="open-meter" data-id="${escapeHtml(m.id)}">${escapeHtml(m.id)}</button></td>
            <td>${m.readings.day1 ?? "—"}</td>
            <td>${m.readings.day2 ?? "—"}</td>
            <td>${m.energy ?? 0}</td>
                     </tr>
        `).join("")}
      </tbody>
    </table>
  </div>`;
}

// Utilities
function computeEnergy(day1, day2) {
  if (day1 == null || day2 == null) return null;
  const diff = Number(day2) - Number(day1);
  return isFinite(diff) ? Math.max(0, Number(diff.toFixed(2))) : null;
}
function diffEnergy(day1, day2, mf = 1) {
  const base = computeEnergy(day1, day2);
  if (base == null) return null;
  const e = base * (isFinite(mf) ? Number(mf) : 1);
  return Number(e.toFixed(2));
}
function computeSlaFlag(val) { return val != null && val !== false; }
function parseYesNo(v) {
  if (v == null) return false;
  const s = String(v).trim().toLowerCase();
  if (s === "yes" || s === "y" || s === "true" || s === "1") return true;
  if (s === "no" || s === "n" || s === "false" || s === "0") return false;
  return Boolean(v);
}
function numberOrNull(v) { const n = Number(v); return isFinite(n) ? n : null; }
function sum(arr) { return arr.reduce((a, b) => a + (Number(b) || 0), 0); }
function sumMeterEnergy(dt) { return sum(dt.meters.map((m) => m.energy ?? 0)); }
function sumDTEnergy(feeder) { return sum(feeder.dts.map((d) => d.metrics.dtEnergy ?? sumMeterEnergy(d))); }
function pct(part, whole) { if (!whole) return 0; return Number(((part / whole) * 100).toFixed(2)); }
function pctLoss(input, output) { if (!input) return 0; return Number((((input - output) / input) * 100).toFixed(2)); }
function fmtPct(v) { return (v == null || isNaN(v)) ? "—" : `${v.toFixed(2)}%`; }
function fmtPctOrDash(v) { return (v == null || isNaN(v)) ? "—" : `${Number(v).toFixed(2)}%`; }
function fmtNumOrDash(v) { return (v == null || isNaN(v)) ? "—" : Number(v).toFixed(2); }
function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// Interactions from tables to detail selection
document.addEventListener("click", (e) => {
  const t = e.target;
  if (t.matches && t.matches("button.link[data-action='open-meter']")) {
    state.currentSelection = { level: "METER", id: t.getAttribute("data-id") };
    renderDetail();
  } else if (t.matches && t.matches("button.link[data-action='open-dt']")) {
    state.currentSelection = { level: "DT", id: t.getAttribute("data-id") };
    renderDetail();
  } else if (t.matches && t.matches("button.link[data-action='open-feeder']")) {
    state.currentSelection = { level: "FEEDER", id: t.getAttribute("data-id") };
    renderDetail();
  }
});

// Mock data bootstrap
function bootstrapWithMockData() {
  const rows = [];
  const regions = ["Region 1", "Region 2", "Region 3"]; 
  let feederCounter = 0, dtCounter = 0, meterCounter = 0;
  for (const region of regions) {
    for (let f = 0; f < 5; f++) {
      const feederId = `F${++feederCounter}`;
      const feederName = feederId;
      for (let d = 0; d < 4; d++) {
        const dtId = `DT${++dtCounter}`;
        const dtName = dtId;
        for (let m = 0; m < 10; m++) {
          const meterId = `M${++meterCounter}`;
          const day1 = 100 + Math.floor(Math.random() * 900);
          const day2 = day1 + Math.floor(Math.random() * 20);
          const feederEnergy = 100 + Math.random() * 50; // synthetic
          const dtEnergy = 80 + Math.random() * 40; // synthetic
                     rows.push({ 
             Region: region, 
             FeederId: feederId, 
             FeederName: feederName, 
             DTId: dtId, 
             DTName: dtName, 
             MeterId: meterId, 
             Day1: day1, 
             Day2: day2, 
             FeederEnergy: feederEnergy, 
             DTEnergy: dtEnergy,
             "Daily energy": Math.random() > 0.3 ? "Yes" : "No",
             "Load Data": Math.random() > 0.4 ? "Yes" : "No",
             "Comm": Math.random() > 0.6 ? "Yes" : "No",
             "NonComm": Math.random() > 0.7 ? "Yes" : "No",
             "Unmapped": Math.random() > 0.9 ? "Yes" : "No",
             "NeverComm": Math.random() > 0.8 ? "Yes" : "No"
           });
        }
      }
    }
  }
  document.getElementById("fileName").textContent = "Mock Data";
  document.getElementById("processedTime").textContent = new Date().toLocaleString();
  buildModelFromThreeSheets(mockFeeders, mockDts, mockMeters);
renderAll();

  renderAll();
}


