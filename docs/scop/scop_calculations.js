/**
 * SCOP (Site Close-Out Package) Calculation Engine
 * ==================================================
 * Core business logic for the SCOP feature. Framework-agnostic — no assumptions
 * about upload handling, storage, or UI. Wire this into the same pipeline used
 * for GR/Invoicing and Decom (chunked Supabase storage, same upload trigger).
 *
 * READ THIS FIRST: SCOP_View_Build_Spec.md (companion doc) has the full narrative
 * history of every decision below — why OAD is carved out, why QuickBase has zero
 * GC-owned items, why the header row position isn't stable. This file is the
 * "what to build"; that doc is the "why", including two real bugs already found
 * and fixed (an OAD miscount and a Service-provider column misalignment) so they
 * aren't reintroduced.
 *
 * ALL settings referenced below (marked with `settings.`) should pull from
 * whatever existing settings/config table already drives GR and Decom's aging
 * thresholds, GC name aliases, etc. Do NOT create a parallel settings system —
 * extend the existing one with the new keys listed in scop_settings_schema.js.
 */

// ---------------------------------------------------------------------------
// 1. COLUMN NAME CONSTANTS
// ---------------------------------------------------------------------------
// CRITICAL: Always resolve columns by NAME, never by position/letter. The
// "Service provider" column has already been observed to shift position
// between tracker exports (was column CA, then moved to BZ) — a same-day
// re-upload during this build caused a real scope miscount (208 vs the
// correct 205) purely from position drift. Name-based lookup is what saved us.

const COLUMNS = {
  SCOPE: "Service provider",           // filter value: exact "Nokia" (case-insensitive)
  PATH_ID: "Path ID",
  HOP: "HOP",
  GENERAL_CONTRACTOR: "General Contractor",
  SITE_CM: "Site CM",
  NEAR_SITE_A: "Near Site Name A",
  FAR_SITE_B: "Far Site Name B",
  CONSTRUCTION_COMPLETE_ACTUAL: "Construction Complete Actual",  // the master gate
  FINAL_CUTOVER: "Final Cutover",       // display-only, not used for gating
  ONE_AND_DONE: "One and Done",         // authoritative Pathwave status field
  QB_DELIVERABLE_STATUS: "Nokia Quickbase Deliverable Status", // authoritative QB status field
};

// Pathwave checklist items — GC-owned only (6 of 8). Asset Form and Packing
// Slip are Nokia-owned and intentionally excluded. This mapping was originally
// discovered via a manually-added ownership tag row that does NOT exist in the
// live tracker — it must stay hard-coded here, not re-derived from the file.
const PATHWAVE_GC_ITEMS = [
  { label: "Install Photos",       colA: "Viaero Install Photos Status Site A", colB: "Viaero Install Photo Status B" },
  { label: "Decom Photos NQR",     colA: "Site A Decom Photos NQR",             colB: "Site B Decom Photos NQR" },
  { label: "Install Photos NQR",   colA: "Site A Install Photos NQR",           colB: "Site B Install Photos NQR" },
  { label: "Red-Line CD",          colA: "Site A Red-Line CD",                  colB: "Site B Red Line CDs" },
  { label: "Decom Asset Form",     colA: "Site A Decom Asset Form",             colB: "Site B Decom Asset Form" },
  { label: "POD",                  colA: "Site A POD",                          colB: "Site B POD" },
];

// QuickBase checklist items — ALL Nokia-owned, confirmed. This list is used
// ONLY for the internal engineering breakdown (never a GC-facing report).
// An earlier version of this mapping had 3 items tagged "GC Owned" — that was
// wrong and has been corrected. Do not reintroduce a GC-owned subset here.
const QUICKBASE_ALL_ITEMS = [
  { label: "B2B Test",                  colA: "QB Site A B2B Test",                colB: "Site B B2B Test" },
  { label: "RFC2455 Test",              colA: "QB Site A RFC2455 Test",            colB: "QB Site B RFC2455 Test" },
  { label: "BER Report",                colA: "QB Site A BER Report",              colB: "QB Site B BER Report" },
  { label: "Consolidated Asset Form",   colA: "QB Site A Consolidated Asset Form", colB: "QB Site B Consolidated Asset Form" },
  { label: "As-Built Final CD",         colA: "QB Site A As-Built Final CD",       colB: "QB Site B As-Built Final CD" },
  { label: "As Built WP",               colA: "QB Site A As Built WP",             colB: "QB Site B As Built WP" },
  { label: "As-Built LLD",              colA: "QB Site A As-Built LLD",            colB: "QB Site B As-Built LLD" },
  { label: "KPI Report",                colA: "QB Site A KPI Report",              colB: "QB Site A KPI Report3" }, // sic — mislabeled in source tracker, this IS Site B's field
  { label: "POD",                       colA: "Site A POD2",                       colB: "Site B POD4" },
  { label: "HW POD",                    colA: "HW POD",                            colB: "HW POD2" },
];

// ---------------------------------------------------------------------------
// 2. NORMALIZATION HELPERS
// ---------------------------------------------------------------------------

function norm(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === "" ? null : s.toLowerCase();
}

function isItemDone(v) {
  const s = norm(v);
  return s !== null && ["yes", "accepted", "complete"].includes(s);
}

function startsWithComplete(v) {
  const s = norm(v);
  return s !== null && s.startsWith("complete");
}

function containsOAD(v) {
  const s = norm(v);
  return s !== null && s.includes("oad");
}

/**
 * Parses Construction Complete Actual into a real Date or null.
 * CRITICAL BUG HISTORY: blank cells in this column have been observed coming
 * through as a literal time-only value (00:00:00 / datetime.time(0,0)) rather
 * than a true null, when read naively. Whatever date-parsing library is used
 * on the ingestion side MUST treat any value that doesn't resolve to a real
 * calendar date (including bare time-only values) as null — never as "present".
 * A naive not-null check on this column previously caused every row to read as
 * "construction complete", which is a severe, silent correctness bug.
 */
function parseConstructionCompleteDate(rawValue) {
  if (rawValue === null || rawValue === undefined) return null;
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    // Reject dates with no real calendar component (i.e., epoch-only / time-only artifacts)
    if (rawValue.getFullYear() < 1990) return null;
    return rawValue;
  }
  const parsed = new Date(rawValue);
  if (isNaN(parsed.getTime()) || parsed.getFullYear() < 1990) return null;
  return parsed;
}

function normalizeGCName(rawName, aliasMap) {
  if (!rawName || String(rawName).trim() === "") return null;
  const trimmed = String(rawName).trim();
  const lower = trimmed.toLowerCase();
  // aliasMap comes from settings (see scop_settings_schema.js) — e.g.
  // { "wavelink": "WaveLink", "viking/capital tower": "Viking" }
  return (aliasMap && aliasMap[lower]) || trimmed;
}

function daysSince(date, asOfDate) {
  if (!date) return null;
  const ms = asOfDate.getTime() - date.getTime();
  return Math.floor(ms / (1000 * 60 * 60 * 24));
}

// ---------------------------------------------------------------------------
// 3. ROW-LEVEL CLASSIFICATION
// ---------------------------------------------------------------------------

/**
 * Classifies a single tracker row into the full SCOP data model.
 * `row` is a plain object keyed by column NAME (as parsed from the sheet).
 * `settings` is the SCOP settings object (see scop_settings_schema.js).
 */
function classifyRow(row, settings, asOfDate) {
  const constructionCompleteDate = parseConstructionCompleteDate(row[COLUMNS.CONSTRUCTION_COMPLETE_ACTUAL]);
  const isConstructionComplete = constructionCompleteDate !== null;
  const agingDays = isConstructionComplete ? daysSince(constructionCompleteDate, asOfDate) : null;

  const oneAndDone = row[COLUMNS.ONE_AND_DONE];
  const isOAD = containsOAD(oneAndDone);
  const pathwaveComplete = startsWithComplete(oneAndDone);

  // 4-way Pathwave status. Order of checks matters: OAD must be checked before
  // falling through to "In Progress" — see spec §4a for the bug this prevents.
  let pathwaveStatus;
  if (!isConstructionComplete) pathwaveStatus = "PENDING_HOP_COMPLETION";
  else if (pathwaveComplete) pathwaveStatus = "COMPLETE";
  else if (isOAD) pathwaveStatus = "OAD";
  else pathwaveStatus = "IN_PROGRESS";

  const qbStatusRaw = row[COLUMNS.QB_DELIVERABLE_STATUS];
  const qbComplete = startsWithComplete(qbStatusRaw);
  const quickbaseStatus = !isConstructionComplete
    ? "PENDING_HOP_COMPLETION"
    : qbComplete ? "COMPLETE" : "PENDING";

  const gcName = normalizeGCName(row[COLUMNS.GENERAL_CONTRACTOR], settings.gcAliasMap);

  return {
    hop: row[COLUMNS.HOP],
    pathId: row[COLUMNS.PATH_ID],
    gc: gcName,
    cm: row[COLUMNS.SITE_CM],
    nearSiteA: row[COLUMNS.NEAR_SITE_A],
    farSiteB: row[COLUMNS.FAR_SITE_B],
    isConstructionComplete,
    constructionCompleteDate,
    agingDays,
    oneAndDoneRaw: oneAndDone,
    isOAD,
    pathwaveStatus,      // "COMPLETE" | "IN_PROGRESS" | "OAD" | "PENDING_HOP_COMPLETION"
    quickbaseStatus,     // "COMPLETE" | "PENDING" | "PENDING_HOP_COMPLETION"
    fullyComplete: isConstructionComplete && pathwaveComplete && qbComplete,
    _rawRow: row,        // kept for itemized missing-item lookups downstream
  };
}

/**
 * Filters raw parsed sheet rows down to Nokia's scope and classifies each one.
 * `rawRows` = array of row objects from the "HOPs" tab, keyed by column name.
 * Header row position is NOT assumed here — pass already-parsed rows in with
 * correct headers resolved upstream (see note in scop_settings_schema.js about
 * the header-row-offset hazard).
 */
function buildScopDataset(rawRows, settings, asOfDate = new Date()) {
  const scoped = rawRows.filter(row => {
    const val = norm(row[COLUMNS.SCOPE]);
    return val === "nokia";
  });

  // Exclude cancelled HOPs (CXLD anywhere in the checklist columns)
  const allChecklistCols = [
    ...PATHWAVE_GC_ITEMS.flatMap(i => [i.colA, i.colB]),
    ["Site A Asset Form", "Site B Asset Form", "Site A Packing Slip", "Site B Packing Slip"],
    ...QUICKBASE_ALL_ITEMS.flatMap(i => [i.colA, i.colB]),
  ].flat();

  const active = scoped.filter(row => {
    return !allChecklistCols.some(col => {
      const v = norm(row[col]);
      return v !== null && v.includes("cxld");
    });
  });

  return active.map(row => classifyRow(row, settings, asOfDate));
}

// ---------------------------------------------------------------------------
// 4. AGGREGATE VIEWS (for the 3 SCOP slide/page views)
// ---------------------------------------------------------------------------

function computeQuickBaseView(dataset) {
  const counts = { COMPLETE: 0, PENDING: 0, PENDING_HOP_COMPLETION: 0 };
  dataset.forEach(r => counts[r.quickbaseStatus]++);
  const total = dataset.length;
  const trackedTotal = counts.COMPLETE + counts.PENDING;
  return {
    complete: counts.COMPLETE,
    pending: counts.PENDING,
    pendingHopCompletion: counts.PENDING_HOP_COMPLETION,
    total,
    trackedTotal,
    percentComplete: trackedTotal > 0 ? Math.round((100 * counts.COMPLETE) / trackedTotal) : 0,
  };
}

function computePathwaveView(dataset) {
  const counts = { COMPLETE: 0, IN_PROGRESS: 0, OAD: 0, PENDING_HOP_COMPLETION: 0 };
  dataset.forEach(r => counts[r.pathwaveStatus]++);

  const byGC = {};
  dataset
    .filter(r => r.isConstructionComplete && r.gc)
    .forEach(r => {
      if (!byGC[r.gc]) byGC[r.gc] = { inProgress: 0, complete: 0, oad: 0 };
      if (r.pathwaveStatus === "IN_PROGRESS") byGC[r.gc].inProgress++;
      else if (r.pathwaveStatus === "COMPLETE") byGC[r.gc].complete++;
      else if (r.pathwaveStatus === "OAD") byGC[r.gc].oad++;
    });

  const oadSites = dataset
    .filter(r => r.pathwaveStatus === "OAD")
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .map(r => ({ hop: r.hop, pathId: r.pathId, gc: r.gc, cm: r.cm, note: r.oneAndDoneRaw, agingDays: r.agingDays }));

  return {
    complete: counts.COMPLETE,
    inProgress: counts.IN_PROGRESS,
    oad: counts.OAD,
    pendingHopCompletion: counts.PENDING_HOP_COMPLETION,
    total: dataset.length,
    byGC,          // { [gcName]: { inProgress, complete, oad } }
    oadSites,       // array, oldest-first
  };
}

function computeOverallView(dataset) {
  const tracked = dataset.filter(r => r.isConstructionComplete);
  const pathwaveComplete = tracked.filter(r => r.pathwaveStatus === "COMPLETE").length;
  const quickbaseComplete = tracked.filter(r => r.quickbaseStatus === "COMPLETE").length;
  const fullyComplete = tracked.filter(r => r.fullyComplete).length;
  return {
    trackedTotal: tracked.length,
    pathwaveComplete,
    pathwavePending: tracked.length - pathwaveComplete,   // NOTE: intentionally includes OAD here — View 3 does not carve OAD out, only View 2 does. See spec §4.
    quickbaseComplete,
    quickbasePending: tracked.length - quickbaseComplete,
    fullyComplete,
    notFullyComplete: tracked.length - fullyComplete,
  };
}

// ---------------------------------------------------------------------------
// 5. ITEMIZED "WHAT'S MISSING" DETAIL (for GC reports + Master Report tabs)
// ---------------------------------------------------------------------------

/**
 * Returns per-site itemized missing-item rows for a set of classified records,
 * using the given item list (PATHWAVE_GC_ITEMS or QUICKBASE_ALL_ITEMS).
 * One row per HOP+Site where at least one item in that list is not done.
 */
function buildItemizedMissingItems(classifiedRows, itemList) {
  const out = [];
  classifiedRows.forEach(r => {
    ["A", "B"].forEach(site => {
      const missing = itemList
        .filter(item => {
          const col = site === "A" ? item.colA : item.colB;
          return !isItemDone(r._rawRow[col]);
        })
        .map(item => item.label);
      if (missing.length === 0) return;
      out.push({
        hop: r.hop,
        pathId: r.pathId,
        gc: r.gc,
        cm: r.cm,
        nearSiteA: r.nearSiteA,
        farSiteB: r.farSiteB,
        site: `Site ${site}`,
        agingDays: r.agingDays,
        missingItems: missing.join(", "),
      });
    });
  });
  return out.sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0));
}

// ---------------------------------------------------------------------------
// 6. PER-GC REPORT GENERATION (populate under each GC, like other reports)
// ---------------------------------------------------------------------------

/**
 * Builds the report payload for ONE GC. This is what should populate under
 * that GC's existing reports section, same as GR/Decom do today.
 * Returns null if the GC has nothing outstanding (no report should be created).
 */
function buildGCReport(dataset, gcName, settings) {
  const gcRows = dataset.filter(r => r.gc === gcName && r.isConstructionComplete);

  const pathwayPendingRows = gcRows.filter(r => r.pathwaveStatus === "IN_PROGRESS");
  const oadRows = gcRows.filter(r => r.pathwaveStatus === "OAD");

  if (pathwayPendingRows.length === 0 && oadRows.length === 0) return null; // nothing to report

  const pathwaveMissingItems = buildItemizedMissingItems(pathwayPendingRows, PATHWAVE_GC_ITEMS);

  const oadDetail = oadRows
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .map(r => ({
      hop: r.hop, pathId: r.pathId, nearSiteA: r.nearSiteA, farSiteB: r.farSiteB,
      cm: r.cm, agingDays: r.agingDays, note: r.oneAndDoneRaw,
    }));

  const completeRows = gcRows.filter(r => r.pathwaveStatus === "COMPLETE");
  const totalOutstanding = pathwayPendingRows.length + oadRows.length;

  return {
    gc: gcName,
    generatedDate: new Date(),
    summary: {
      totalHOPs: completeRows.length + totalOutstanding,
      completed: completeRows.length,
      outstanding: totalOutstanding,       // includes OAD, for an honest total
      oldestOutstandingDays: totalOutstanding > 0
        ? Math.max(...pathwayPendingRows.concat(oadRows).map(r => r.agingDays || 0))
        : null,
    },
    sections: {
      pathwave: pathwaveMissingItems,   // excludes OAD rows — see spec §4a
      oad: oadDetail,                   // only included if non-empty
      // NOTE: deliberately no quickbase section — all QuickBase items are
      // Nokia-owned, confirmed. Never add a QuickBase section to a GC report.
    },
  };
}

/**
 * Builds GC reports for every GC with outstanding work. This is the function
 * to call on tracker upload — mirrors how GR/Decom populate per-GC on ingest.
 */
function buildAllGCReports(dataset, settings) {
  const gcNames = [...new Set(dataset.map(r => r.gc).filter(Boolean))].sort();
  return gcNames
    .map(gc => buildGCReport(dataset, gc, settings))
    .filter(report => report !== null);
}

// ---------------------------------------------------------------------------
// 7. EMAIL GENERATION (mailto-style, matching existing GC email buttons)
// ---------------------------------------------------------------------------

/**
 * Builds the subject/body for a GC's SCOP email. Pairs with the existing
 * mailto-button pattern used by other trackers — this function only produces
 * text, the actual button/contact-lookup wiring should reuse whatever
 * component already renders the Decom/GR email buttons.
 */
function buildGCEmail(gcReport, settings) {
  const { gc, summary, sections } = gcReport;
  const dateStr = gcReport.generatedDate.toLocaleDateString("en-US", { month: "2-digit", day: "2-digit", year: "numeric" });

  const threshold = settings.emailPriorityThresholdDays ?? 60;
  const cap = settings.emailPriorityCap ?? 5;

  const nonOadPending = sections.pathwave; // already excludes OAD
  const priorityItems = [...new Map(nonOadPending.map(r => [r.hop, r])).values()] // dedupe by HOP (site-level rows collapse to one HOP entry)
    .filter(r => (r.agingDays || 0) >= threshold)
    .sort((a, b) => (b.agingDays || 0) - (a.agingDays || 0))
    .slice(0, cap);

  const subject = `Weekly SCOP Action Items — ${gc} (${dateStr})`;

  let body = `Hi ${gc},\n\nSummary for ${gc}:\n`;
  body += `• Total HOPs: ${summary.totalHOPs}\n`;
  body += `• Completed: ${summary.completed}\n`;
  body += `• Outstanding: ${summary.outstanding}\n`;
  body += `• Oldest outstanding item: ${summary.oldestOutstandingDays} days since construction complete\n\n`;
  body += `Full detail is in the attached report, sorted oldest-first.\n\n`;

  if (priorityItems.length > 0) {
    body += `Top priority (aging ${threshold}+ days since construction complete):\n`;
    body += priorityItems.map(r => `• ${r.hop} — ${r.agingDays} days`).join("\n");
    body += `\n\nThese are the longest-outstanding items — let's prioritize closing these out first.\n\n`;
  } else {
    body += `Nothing is over the ${threshold}-day priority threshold yet, but please keep the attached list moving.\n\n`;
  }

  body += `Please review the attached report and let us know a target date for the outstanding items. Happy to hop on a call if anything needs clarifying.\n\nThanks,\nCJ`;

  return {
    subject,
    body,
    attachmentFilename: `SCOP_GC_Report_${gc.replace(/[/ ]/g, "_")}_${dateStr.replace(/\//g, "-")}.xlsx`,
  };
}

// ---------------------------------------------------------------------------
// EXPORTS
// ---------------------------------------------------------------------------

module.exports = {
  COLUMNS,
  PATHWAVE_GC_ITEMS,
  QUICKBASE_ALL_ITEMS,
  buildScopDataset,
  computeQuickBaseView,
  computePathwaveView,
  computeOverallView,
  buildItemizedMissingItems,
  buildGCReport,
  buildAllGCReports,
  buildGCEmail,
  // exported for testing / reuse
  parseConstructionCompleteDate,
  normalizeGCName,
  classifyRow,
};
