# SCOP — Claude Code Implementation Brief
*Read alongside `SCOP_View_Build_Spec.md` (full business-logic history and every
decision's rationale), `scop_calculations.js` (the actual logic to wire in), and
`scop_settings_schema.js` (new settings to add). This doc is the "where it plugs in."*

**Status: approved to build. CJ has reviewed and signed off on the deck, Master
Report, GC reports, and email pattern end-to-end.**

---

## 1. Where This Lives in Ciege

Add a new tab under the **Reports** section, alongside the existing GR/Invoicing
and Decom tabs. Call it **SCOP**. Same navigation pattern, same visual shell —
this should feel like a sibling feature, not a bolted-on one-off.

## 2. Upload → Auto-Populate (matches GR/Decom exactly)

When a SCOP tracker file is uploaded:

1. Parse the `HOPs` tab. **Auto-detect the header row** (scan for the row
   containing both `"HOP"` and `"General Contractor"` as cell values — do not
   hardcode row 4 or row 5; see `headerRowOverride` in the settings schema for
   why this specific column has already moved once).
2. Run `buildScopDataset(rawRows, settings)` from `scop_calculations.js`.
3. Run `buildAllGCReports(dataset, settings)` — this returns one report object
   per GC that has outstanding work.
4. **Populate each GC's existing reports area with its SCOP report**, the same
   way a Decom or GR report already shows up under that GC today. If GR/Decom
   reports live at a path like `/gc/{gcId}/reports`, SCOP's report should
   appear there as another entry, not a separate page.
5. Store the full dataset (all classified rows, not just the GC-facing subset)
   so the Master Report views and the 3 SCOP dashboard views can be rendered
   from it without re-parsing the upload.

## 3. The Three Dashboard/Slide Views

Build these as live Ciege views (not just exportable slides) using
`computeQuickBaseView`, `computePathwaveView`, and `computeOverallView`:

- **View 1 — QuickBase Status**: `computeQuickBaseView(dataset)`. 3 cards
  (Complete → In Progress → Pending HOP Completion, in that order), a
  full-scope proportional bar, headline % using `trackedTotal` as denominator.
- **View 2 — Pathwave SCOP — Contractor Action Items**: `computePathwaveView(dataset)`.
  4 cards (Complete & Approved → In Progress — GC Action → OAD — Tracked
  Separately → Pending HOP Completion). Two side-by-side tables below: by-GC
  breakdown (`byGC`) and the OAD sites list (`oadSites`).
- **View 3 — Overall SCOP Status**: `computeOverallView(dataset)`. Two side-by-
  side card pairs (Pathwave Complete/Pending, QuickBase Complete/Pending), plus
  a "both complete" pair. Scoped only to construction-complete HOPs — the
  Pending-HOP-Completion bucket does not appear on this view at all.

Exact colors, ordering, and card labels are specified in `SCOP_View_Build_Spec.md`
§4 — follow that literally, several of these (wording, ordering, which bucket is
red vs amber) were iterated on with CJ and are not arbitrary.

## 4. Master Report — Same Data, Tabular/Filterable View

The 9-tab Excel Master Report CJ has been using (Summary, Outstanding List,
Pathwave Pending, Pathwave Complete, Pathwave OAD Sites, QuickBase Pending,
QuickBase Complete, Fully Complete, Filter Cheat Sheet) should become a
**filterable table view inside the SCOP tab**, not just a downloadable file —
though keeping an "export to Excel" action for it is fine and expected (matches
how GR/Decom likely already support export). Each "tab" from the spec maps to a
filter/segment of the same underlying dataset:

- Outstanding List → `dataset.filter(r => !r.isConstructionComplete)`
- Pathwave Pending → `buildItemizedMissingItems(dataset.filter(r => r.pathwaveStatus === "IN_PROGRESS"), PATHWAVE_GC_ITEMS)`
- Pathwave OAD Sites → `computePathwaveView(dataset).oadSites`
- QuickBase Pending → `buildItemizedMissingItems(dataset.filter(r => r.quickbaseStatus === "PENDING"), QUICKBASE_ALL_ITEMS)`
- Fully Complete → `dataset.filter(r => r.fullyComplete)`

The Filter Cheat Sheet tab's content should become inline help text/tooltips in
the UI rather than a literal spreadsheet tab, if that fits Ciege's existing
patterns better — its job (letting CJ understand *why* a HOP landed where it
did) is more naturally a UI affordance than a static reference sheet once this
is live.

## 5. GC Emails

Use `buildGCEmail(gcReport, settings)` to produce subject/body/attachment
filename. Wire the send button to the **same mailto-link mechanism already used
for Decom/GR GC emails** — auto-populate To/CC from the existing GC contact
list setting, pre-fill subject and body, and require manual attachment of the
downloaded report (mailto can't carry attachments — this is a hard platform
constraint, not a gap to build around).

Only render an email button for a GC if `buildGCReport()` returned non-null for
them (i.e., they have outstanding work) — don't show a send option for a
fully-clean GC.

## 6. "Download All Reports"

Add each generated GC SCOP report (as an `.xlsx`, matching the existing weekly
report format) into whatever bundling mechanism already produces "Download All
Reports" for GR/Decom. This should be a small addition to that existing
bundler, not a new download pathway — SCOP's per-GC files should just become
additional entries alongside the GR/Decom files already in that zip/bundle.

## 7. Settings

Add the new keys from `scop_settings_schema.js` into wherever GR/Decom's
tunables already live in the settings UI — likely a "SCOP" section/card on the
same settings page, matching how GR and Decom presumably each have their own
section there already. Reuse (don't duplicate) the GC alias map and GC contact
list settings — those are program-wide concepts, not SCOP-specific.

## 8. Known Data Hazards (don't relearn these the hard way)

Both of these already caused real, silent bugs during development — see
`SCOP_View_Build_Spec.md` for the full incident write-ups:

1. **Header row is not stable between tracker versions.** Auto-detect it (§2
   above); never hardcode a row number.
2. **The `Service provider` column's position is not stable between tracker
   exports.** Always resolve columns by name (already how `scop_calculations.js`
   is written) — never by letter/index. If the Nokia-scope count ever comes
   back meaningfully different from the last known-good run (~205 as of this
   writing) without an obvious reason, check column alignment before trusting
   the new number.
3. **`Construction Complete Actual` blank cells can parse as a time-only
   artifact instead of null**, depending on the parsing library — `parseConstructionCompleteDate()`
   in `scop_calculations.js` already guards against this (rejects any parsed
   date with a pre-1990 year), but if this column is ever read anywhere else
   in the codebase, apply the same guard.

## 9. Validation Checklist Before Calling This Done

Reload the current live tracker and confirm these cross-foot (values as of the
09/08/2026 corrected run — use as a regression baseline, not a hardcoded
expectation):

- Total Nokia HOPs ≈ 205, and Pathwave 4-way split sums to that total exactly
- QuickBase Complete + Pending sums to the construction-complete tracked total
- Pathwave Complete + Pending (View 3's combined bucket) also sums to tracked total
- Every GC report's `summary.totalHOPs` equals `completed + outstanding` for that GC
- No QuickBase section appears anywhere on a GC-facing report or email
