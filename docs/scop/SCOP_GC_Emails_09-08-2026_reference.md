# Weekly SCOP Action Emails — All GCs (APPROVED REFERENCE)
*Generated 09/08/2026 — corrected run. This is the exact wording/format CJ approved.*

**Pattern:**
- Greeting: GC by name, not an individual CM.
- Summary block at top: Total HOPs, Completed, Outstanding, oldest outstanding age.
- "Outstanding" includes OAD-blocked sites; the priority callout list below is non-OAD only.
- Priority callout: any non-OAD item aging 60+ days, capped at 5.
- QuickBase never mentioned.
- Attach the GC's Excel report (Pathwave + OAD sections only).
- Only GCs with >= 1 outstanding item get an email.

## Approved summary numbers per GC (regression baseline)
| GC | Total HOPs | Completed | Outstanding | Oldest (days) | Priority items shown |
|---|---|---|---|---|---|
| Capital Tower | 19 | 5 | 14 | 119 | 5 |
| Ethos | 14 | 4 | 10 | 272 | 5 |
| MZI | 17 | 2 | 15 | 144 | 5 |
| Mastec | 11 | 5 | 6 | 179 | 4 |
| NV Tel | 8 | 3 | 5 | 125 | 1 |
| Site Property | 11 | 7 | 4 | 50 | 0 (nothing over 60d) |
| Steimel | 2 | 1 | 1 | 207 | 1 |
| Viking | 9 | 6 | 3 | 118 | 0 (nothing over 60d) |
| Vikor | 25 | 24 | 1 | 6 | 0 (nothing over 60d) |
| WaveLink | 13 | 11 | 2 | 71 | 0 (nothing over 60d) |

10 GCs receive an email. (Report-file recipients are 9 — the spec's §7 count is
Pathwave+OAD file recipients; the email list here includes GCs whose only outstanding
work still produces a summary + "nothing over threshold" note.)

## Exact body template
```
Hi {GC},

Summary for {GC}:
• Total HOPs: {completed + outstanding}
• Completed: {completed}
• Outstanding: {outstanding}
• Oldest outstanding item: {oldestOutstandingDays} days since construction complete

Full detail is in the attached report, sorted oldest-first.

{if priorityItems.length > 0}
Top priority (aging 60+ days since construction complete):
• {HOP} — {agingDays} days
  ...(up to 5, oldest first)

These are the longest-outstanding items — let's prioritize closing these out first.
{else}
Nothing is over the 60-day priority threshold yet, but please keep the attached list moving.
{endif}

Please review the attached report and let us know a target date for the outstanding items. Happy to hop on a call if anything needs clarifying.

Thanks,
CJ
```

## Sample — Capital Tower priority list (verbatim, for regression)
```
• NE-WAUNETA-NE-WAUNETA_SW — 119 days
• WY-TORRINGTON_SOUTH-NE-MORRILL_DT — 118 days
• NE-GORDON-NE-GORDON_SOUTH — 91 days
• NE-LITCHFIELD-NE-LOUP_CITY — 81 days
• NE-WHITECLAY_SOUTH-NE-RUSHVILLE_NORTH — 62 days
```

## Sample — Ethos priority list (verbatim)
```
• CO-ELIZABETH_DT-CO-ELBERT_NORTH — 272 days
• NE-KEARNEY_SALES_TWO-NE-KEARNEY_CABELAS — 144 days
• NE-CAMBRIDGE-NE-WILSONVILLE — 121 days
• NE-MEAD_EAST-NE-WAHOO_NORTH — 113 days
• NE-MERRIMAN-NE-MERRIMAN_SOUTH — 102 days
```

## Sample — MZI priority list (verbatim)
```
• NE-ELWOOD-NE-BERTRAND — 144 days
• NE-HAYES_CENTER_NORTH-NE-HAYES_CENTER — 143 days
• NE-EUSTIS_DT-NE-EUSTIS — 123 days
• NE-EUSTIS-NE-ELWOOD — 119 days
• NE-HOLDREGE-NE-BERTRAND — 75 days
```
