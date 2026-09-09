/**
 * SCOP Settings Schema — ADDITIVE ONLY
 * =====================================
 * These are new settings KEYS to add to whatever config/settings system already
 * drives GR and Decom (aging thresholds, GC name aliases, contact lists, etc.).
 *
 * DO NOT build a separate settings table/page for SCOP. If GR or Decom already
 * has a "GC Name Aliases" setting, SCOP should read/write the SAME one — a GC
 * called "WaveLink" is the same GC everywhere in Ciege, not a SCOP-specific
 * concept. Same logic applies to the GC contact list used for email To/CC.
 *
 * The settings genuinely new to SCOP (thresholds, keyword lists, the header-row
 * safety net) should still live in whichever settings UI/table already holds
 * per-report-type config for GR/Decom, just with a "scop" namespace/prefix —
 * follow whatever pattern those two already use for their own tunables.
 */

const SCOP_SETTINGS_SCHEMA = {

  // ── Reused from existing settings — do NOT duplicate ─────────────────────
  // gcAliasMap:            <- same setting GR/Decom already use for name merging
  // gcContactList:         <- same setting GR/Decom already use for email To/CC
  // reportGenerationCadence: <- if GR/Decom already have a schedule setting, reuse it

  // ── New settings specific to SCOP ───────────────────────────────────────

  agingColorThresholds: {
    label: "Aging Color Thresholds (days)",
    type: "object",
    default: { green: 30, amber: 60 },  // < green = green, green–amber = amber, >= amber = red
    description:
      "Days since Construction Complete before an item's aging number turns amber, " +
      "then red. Picked from the current backlog's actual spread (median ~93 days at " +
      "launch) — expected to need tightening (e.g. 15/30) as the backlog clears and " +
      "the current 30/60 split stops being meaningfully selective.",
    uiHint: "Two number inputs: 'Amber after ___ days', 'Red after ___ days'",
  },

  emailPriorityThresholdDays: {
    label: "Email Priority Callout Threshold (days)",
    type: "number",
    default: 60,
    description:
      "Minimum aging (days since construction complete) for a HOP to appear in a " +
      "GC email's 'Top Priority' callout list. Independent from agingColorThresholds " +
      "— this gates what shows in the email, the other gates display color.",
  },

  emailPriorityCap: {
    label: "Max Priority Items Per Email",
    type: "number",
    default: 5,
    description:
      "Maximum number of HOPs shown in a GC email's priority callout list, even if " +
      "more qualify under the threshold above. Shows the N oldest.",
  },

  oadKeywords: {
    label: "OAD Detection Keywords",
    type: "string[]",
    default: ["oad"],
    description:
      "Case-insensitive substrings checked against the 'One and Done' column's free-" +
      "text value. Any match reclassifies that HOP from 'In Progress' into a separate " +
      "'OAD' bucket, excluded from GC action-item counts and reports. Currently a " +
      "single keyword; exposed as a list in case a second free-text carve-out category " +
      "is ever needed the way OAD was.",
    uiHint: "Tag input, comma-separated",
  },

  headerRowOverride: {
    label: "Tracker Header Row (safety override)",
    type: "number | 'auto'",
    default: "auto",
    description:
      "Row number where column headers live in the uploaded tracker's 'HOPs' tab. " +
      "Auto-detect by scanning for a row containing 'HOP' and 'General Contractor' as " +
      "cell values. This setting exists because the header row has already moved once " +
      "(row 4 vs row 5) between a working copy and the live tracker in the same week — " +
      "if auto-detection ever fails silently, this override lets CJ correct it without " +
      "a code change. Show the auto-detected value in the UI even when set to 'auto', " +
      "so a wrong detection is visible before it produces bad numbers.",
    uiHint: "Number input with an 'Auto-detect' toggle; always display the resolved value",
  },

  scopReportAsOfDate: {
    label: "Report As-Of Date Override",
    type: "date | null",
    default: null,
    description:
      "For regenerating a past week's numbers or testing — overrides 'today' for all " +
      "aging calculations. Null means use the actual current date.",
    uiHint: "Optional date picker, defaults empty/unused",
  },

  // ── Ownership mapping — exposed as read/adjustable, not hidden in code ────
  pathwaveGcOwnedItems: {
    label: "Pathwave Items Owned by GC",
    type: "string[]",
    default: [
      "Install Photos", "Decom Photos NQR", "Install Photos NQR",
      "Red-Line CD", "Decom Asset Form", "POD",
    ],
    description:
      "Which of the 8 Pathwave checklist items count as GC action items on reports " +
      "and emails. Currently 6 of 8 — Asset Form and Packing Slip are Nokia-owned and " +
      "excluded. If this ownership split is ever renegotiated, change it here rather " +
      "than in code. NOTE: the live tracker carries no ownership metadata of its own " +
      "(a manually-added tag row existed only in a working copy, never the production " +
      "file) — this setting is the sole source of truth.",
    uiHint: "Checklist of all 8 Pathwave item names, checked = GC-owned",
  },

  quickbaseHasGcOwnedItems: {
    label: "QuickBase Has Any GC-Owned Items",
    type: "boolean",
    default: false,
    description:
      "Confirmed false — all QuickBase items are Nokia-owned. Kept as an explicit " +
      "setting (rather than just omitting QuickBase from GC reports in code) so that " +
      "if this business rule ever changes, it's a settings flip, not a redeploy. If " +
      "flipped true, a UI for selecting which QuickBase items are GC-owned would need " +
      "to be added alongside — mirror the pathwaveGcOwnedItems pattern above.",
  },
};

module.exports = { SCOP_SETTINGS_SCHEMA };
