# Ciege — SCOP View Build Spec (v3)

This supersedes v1 and v2. Read the whole thing before implementing — several rules
here directly overturn earlier assumptions (see "Session Log" at the bottom for why).

---

## 0. Scope — Which HOPs Are Yours

**Simplest, current method:** the COP tracker itself now has a `Service provider`
column (Column CA). Filter it to the exact value `"Nokia"` (case-insensitive, trimmed).
As of this build that yields exactly **205 unique HOPs** — this is CJ's full program
scope (not split further by individual PM; confirmed on the call that the full 205 is
correct, not just CJ's personal subset).

**Known column quirk:** blank cells in `Service provider` render as the string
`00:00:00` in some rows (a time-format artifact, same family of bug as the date issue
below) — don't mistake that value for a real service provider; only `"Nokia"` counts.

**Fallback / cross-check method (from v2, still valid as a sanity check):** the master
tracker `Viaero_-_Construction_Schedule_Tracker_-_WIP_2_0.xlsx`, tab `HOPs`, header row
2 (`header=1`), filtering Column B (`DON 444`) for the exact value `"DON 444"` also
yields 205 HOPs, and every one of them was confirmed present in the COP tracker. If
`Service provider` ever drifts from this count, cross-check against the master tracker's
DON 444 filter the same way v2 specified.

---

## 1. Authoritative Status Fields — Use These Directly, Don't Recompute

- **Pathwave status** — Column S, `One and Done`. Free-text values include `Complete`,
  `Rejections`, `Pending...`, `Resubmitted...`. Normalize: anything starting with
  `"complete"` (case-insensitive) = **Complete**; everything else (including blank) =
  **not complete**.
- **QuickBase status** — Column BT, `Nokia Quickbase Deliverable Status`. Values include
  `Complete`, `Work in progress`, blank. Same normalization: starts with `"complete"` =
  **Complete**, else not complete.

The raw 8-item (Pathwave) / 11-item (QuickBase) checklists per site are **still needed**
for the itemized "what's actually missing" detail (weekly GC reports, Master Report
Pending tabs) — just not for the top-line status rollup anymore.

---

## 2. Item-Level Ownership — Row 2 Tags Are Authoritative, Ignore Named Individuals

### Pathwave — GC-Owned (6 of 8 items per site)
Install Photos, Decom Photos NQR, Install Photos NQR, Red-Line CD, Decom Asset Form, POD

### Pathwave — Nokia-Owned (2 of 8 items per site)
Asset Form, Packing Slip

### QuickBase — (SUPERSEDED by §9a) all Nokia-owned

**Weekly GC reports and the Master Report's Pending tabs must only list GC-owned items**
as that GC's action items.

---

## 3. The Master Gate — Construction Complete Actual, With Exact Terminology

**Column P, `Construction Complete Actual`, is the single master gate for all headline
numbers.** Every KPI card, chart, and percentage across all 3 views must be built off
this same gate.

### Required terminology (exact wording)
- **"Pending HOP Completion"** = HOPs where `Construction Complete Actual` has no valid
  date. Never called "Outstanding" or "No Action Yet" in the final UI.
- **"In Progress"** = construction IS complete, but the relevant status field is not yet `Complete`.
- **"Complete"** = construction complete AND the relevant status field is `Complete`.

### Critical parsing bug (do not reintroduce)
`Construction Complete Actual` blank cells come through as `datetime.time(0, 0)`, not a
true null. Always parse and check the *parsed* result; reject any parsed date with a
pre-1990 year. A raw not-null check reads every row as "construction complete."

### Validated numbers — 09/04 CJ working copy (older baseline)
Total 205, Construction Complete 153, Pending HOP Completion 52, Pathwave Complete 96 /
In Progress 39 / OAD 18, QuickBase Complete 111 / In Progress 42, Fully Complete 75, Not
Yet Fully Complete 78.  (96+39+18+52=205; 111+42=153; 75+78=153)

---

## 4. The Three Views

### View 1 — "QuickBase Status" (internal, no GC anywhere)
- Header is just "QuickBase Status".
- 3 KPI cards, left to right: **Complete → In Progress → Pending HOP Completion.**
- Colors: Complete = green, In Progress = **amber** (not red), Pending HOP Completion = grey.
- Headline stat: `% = QuickBase Complete / trackedTotal` (construction-complete denominator, NOT 205).
- Horizontal proportional bar showing **all 205 HOPs**, segments green / amber / grey.
- Caption: "Full program view — all 205 HOPs."

### View 2 — "Pathwave SCOP — Contractor Action Items"
- 4 KPI cards, left to right: **Complete & Approved → In Progress — GC Action → OAD —
  Tracked Separately → Pending HOP Completion.**
- Colors: Complete = green, In Progress = amber, OAD = navy (informational), Pending HOP Completion = grey.
- Full-205 proportional bar, same 4 segments/colors/order.
- Two tables side by side:
  - **Left — "Pathwave Items by GC"**: GC, In Progress count, Complete count. Excludes
    OAD from In Progress. Sorted by In Progress desc.
  - **Right — "OAD Sites (Awaiting OAD, Not a GC Item)"**: HOP, GC, raw `One and Done` text verbatim.
- Never use red in card/section colors. In Progress = amber everywhere.
- Complete card label is **"Complete & Approved"**.

### View 3 — "Overall SCOP Status — Program Summary" (customer-shareable)
- No GC names.
- Two side-by-side pairs: PATHWAVE (Complete / Pending) and QUICKBASE (Complete / Pending).
  Uses the plain word "Pending".
- "Both Complete — Full SCOP Close-Out": 2 cards, "100% Complete" and "Not Yet Fully Complete."
- **Does NOT show a Pending-HOP-Completion bucket at all** — scoped only to
  construction-complete HOPs.
- **Deliberate exception:** View 3's Pathwave "Pending" still includes the OAD HOPs
  uncarved. Do not "fix" this to match View 2.

### Design tokens
Navy `#124191`, Teal `#00A0B0`, Green `#1E8449`, Amber `#B7791F`, Grey `#6B7280` /
light-grey bar `#D0D5DD`, light bg `#F7F9FB`, border `#E2E8F0`. No red in section/card
colors. Red is used ONLY for the aging/priority number in itemized tables (§6).

---

## 4a. The OAD Carve-Out — Root Cause of a Real Data-Alignment Bug

18 HOPs (of the 153 construction-complete ones, 09/04 baseline) have the substring
`"OAD"` somewhere in their `One and Done` value (free text). They're waiting on an OAD
milestone, not a GC checklist item. Old logic lumped them into "In Progress", inflating
it by 18.

```
is_oad = "oad" in str(row["One and Done"]).lower()   # case-insensitive substring, not exact
is_complete = normalize(row["One and Done"]).startswith("complete")

if not constr_complete:      status = "Pending HOP Completion"
elif is_complete:            status = "Complete"
elif is_oad:                 status = "OAD"
else:                        status = "In Progress"
```

Applies to View 2, every weekly GC report, and the Master Report's Pathwave Pending /
OAD Sites tabs. Does NOT apply to View 3 or the QuickBase side.

---

## 5. Master SCOP Report (Excel, 9 tabs)

1. **Summary** — KPI table (4-way Pathwave split + Outstanding + QuickBase + Fully Complete).
2. **Outstanding List** — HOPs with Pending HOP Completion. Cols: HOP, Path ID, GC, CM, Near Site A, Far Site B.
3. **Pathwave Pending** — itemized, one row per HOP+Site with a GC-owned Pathwave item
   missing. **Excludes OAD.** Cols: HOP, Path ID, GC, CM, Near Site A, Far Site B, Site,
   Days Since Complete, Missing Items. Sorted oldest-first.
4. **Pathwave Complete** — simple list (HOP, Path ID, GC, CM, Near/Far Site).
4a. **Pathwave OAD Sites** — the OAD HOPs, kept separate. Cols: HOP, Path ID, GC, CM,
   Near Site A, Far Site B, Days Since Complete, and the raw `One and Done` verbatim ("Note").
5. **QuickBase Pending** — itemized like tab 3, against ALL 11 QuickBase items.
   Reframed (§9a) as "QuickBase — Pending Items (Nokia-Internal, NOT a GC Action Item)".
   OAD carve-out does NOT apply here.
6. **QuickBase Complete** — simple list.
7. **Fully Complete** — HOPs 100% done on both Pathwave and QuickBase.
8. **Filter Cheat Sheet** — plain-language column + condition table for every list.
   Must document the OAD substring rule.

**Formula-injection hazard:** never write a cell value starting with `=`. Prefix
descriptive text (e.g. `Exact value "Nokia"...`).

**Print layout:** every tab needs landscape, fitToWidth=1, fitToHeight=0,
fitToPage=True, explicit print_area. Re-set page setup after any column insert.

---

## 6. Aging / Priority Column

Added to: Master Report's Pathwave Pending and QuickBase Pending tabs, and both sections
of every individual GC weekly report.

```
aging_days = (today - Construction Complete Actual).days   # only for rows with a valid date
```

- Column label: **"Days Since Complete"**.
- Color-code: **green < 30, amber 30–59, red 60+.** (Only place red is used.)
- **Sort every itemized list oldest-first** (highest aging at top).
- Real distribution: median ~93 days, some 200+. The 30/60 thresholds may need
  tightening as the backlog clears.

---

## 7. Weekly GC Reports (Excel, one per GC) — QuickBase removed (§9a)

Structure per file:
- Title + generated date + one-line instruction.
- **Two sections**:
  1. "PATHWAVE — GC-Owned Items" (navy header) — **excludes OAD**. One row per HOP+Site
     with an outstanding GC-owned Pathwave item. Cols: HOP, Path ID, Near Site (A), Far
     Site (B), Site CM, Site, Days Since Complete, Missing Items. Sorted oldest-first.
  2. "OAD — Awaiting OAD (Not a GC Checklist Item)" *(only for GCs with ≥1 OAD HOP)* —
     Cols: HOP, Path ID, Near Site (A), Far Site (B), Site CM, Days Since Complete, raw
     `One and Done` comment. No Site A/B split or Missing Items column.
- Gate: only construction-complete HOPs appear.
- **A GC only gets a file if they have ≥1 item in Pathwave or OAD.** Currently 9 of 14
  active GCs qualify (InSite, Tech CX, Vikor dropped — QuickBase-only).

---

## 8. Weekly GC Emails — Tied 1:1 to the Reports

- **One combined email per GC** (Pathwave + OAD together). No QuickBase.
- Subject: `Weekly SCOP Action Items — {GC} (MM/DD/YYYY)`
- Body:
  1. Greeting addresses the **GC by name** ("Hi MZI,").
  2. **Summary block**: Total HOPs (complete + outstanding), Completed, Outstanding
     (INCLUDES OAD-blocked HOPs — don't drop OAD from this count), Oldest outstanding
     item's age.
  3. **Priority callout list**: pulled from **non-OAD outstanding items only**. Any
     qualifying HOP aging **60+ days** included, capped at 5 (oldest). If zero qualify,
     say so plainly rather than an empty list.
  4. Closing ask for a target completion date, offer a call.
- **Attach the GC's Excel report file** (Pathwave + OAD only).
- **Only send to GCs with ≥1 outstanding item.**
- Send mechanism — CONFIRMED: same `mailto:` pattern as Decom/GR. To/CC auto-populated
  from GC contact info. Report file downloaded separately and manually attached.

---

## 9. Implementation Notes
- Chunked Supabase storage required for large payloads (established pattern from Decom/CR).
- Do not treat legacy `Buckets`, `COP Deliverables`, `Asset Form Tracker` tabs, or old
  `SCOP Owner`/`SCOP Status`/`SCOP Notes` columns (BV–BX) as live data. `One and Done`
  and `Nokia Quickbase Deliverable Status` are the correct live fields.
- GC name normalization: trim + case-insensitive. Known collisions: `WaveLink` /
  `Wavelink` → `WaveLink`; `Viking` / `Viking/Capital Tower` → `Viking` (assumed, not
  explicitly confirmed).

---

## 9a. QuickBase Is NOT a GC Action Item — Corrected

**Supersedes §2's QuickBase GC-owned claim.** All QuickBase items are Nokia-owned.
- Weekly GC reports no longer have a QuickBase section — only Pathwave and OAD.
- A GC with only QuickBase items pending, clean Pathwave, no OAD → no report file.
  Dropped InSite, Tech CX, Vikor (12 → 9 report-receiving GCs).
- Master Report's `QuickBase Pending` tab reframed as Nokia-internal, itemizes against
  all 11 QuickBase items, subtitle says never send to a GC.
- GC emails no longer mention QuickBase.

---

## 10. Validated Against the Real Live Tracker

### Critical structural difference: header row offset
CJ's working copy has headers on **row 5** (`header=4`) because of an inserted
ownership-tag row. **The real live tracker has headers on row 4** (`header=3`). Any
implementation must detect or configure this per-file. Ciege should assume row 4 going
forward.

### Row 2 ownership tags are not in the live file — and that's fine
The ownership mapping is hard-coded in the build logic, never re-read from the file at
runtime. If ownership ever changes it must be communicated directly.

### Everything else checked clean
Column names, sheet names, value conventions (Yes/No/Accepted, OAD free-text, WaveLink/
Wavelink casing) all match. No CXLD HOPs currently in Nokia scope. Path ID fully
populated. Scope grew 205 → 208 → (corrected) 205 between snapshots.

### Numbers from the actual production run (09/08/2026, live file, CORRECTED)
| Metric | 09/04 (CJ copy) | 09/08 (live, corrected) |
|---|---|---|
| Total Nokia HOPs | 205 | **205** |
| Construction Complete | 153 | **155** |
| Pending HOP Completion | 52 | **50** |
| Pathwave Complete | 96 | **94** |
| Pathwave In Progress | 39 | **42** |
| Pathwave OAD | 18 | **19** |
| QuickBase Complete | 111 | **114** |
| QuickBase Pending | 42 | **41** |
| Fully Complete | 75 | **77** |

First live upload gave 208 (wrong) — the `Service provider` column had drifted position.
Read by NAME, always. Any Nokia-scope count not near 205 = check column alignment first.
This corrected 09/08 run is the regression baseline.

---

## Session Log (abridged)
- v1→v2: COP tracker has no DON 444 column; added master-tracker cross-check; fixed the
  Construction Complete blank-cell parse bug the first time.
- v2→v3: real maintained status fields (`One and Done`, `Nokia Quickbase Deliverable
  Status`) replace from-scratch checklist derivation; ownership tagged per-column;
  `Service provider` column added; `Construction Complete Actual` is the one true gate;
  deck restructured to QuickBase/Pathwave/Overall; aging + GC emails added.
- v3 mid-review: OAD carve-out bug (§4a) — 18 OAD HOPs inflating In Progress. Fixed
  across View 2, GC reports, Master Report. Renamed View 2 Complete card to "Complete &
  Approved". View 1 / View 3 NOT given the carve-out.
- v3 continued: QuickBase ownership fully corrected (§9a) — all Nokia-owned. QuickBase
  removed from GC reports and emails. GC email summary block added (Outstanding includes
  OAD for an honest total; priority callout excludes OAD).
- Final review pass: Master Report `QuickBase Pending` tab + Filter Cheat Sheet updated
  to the §9a model. Every deliverable now internally consistent.
