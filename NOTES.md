# CAB Digital Report — what the dashboard found in the workbook

Notes for whoever maintains `EXTERNAL_ CAB Digital Report.xlsx`. These are things
the dashboard has to work around. Fixing them in the workbook would make several
panels better; none of them stop the dashboard from running.

Verified against the file as of 24 Aug 2026.

---

## 1. `Digital Projections` is three tables stacked in one tab

The tab is not one projection. It holds three scenario blocks, each with its own
header row and its own `Totals` row:

| Block | Title in column A | Header row | Rows | State |
|---|---|---|---|---|
| 1 | Digital Projections (Low Investment) | 2 | Jan–Dec 2026 | **Unusable** — every value is `$0.00` or `#N/A`, and the Totals row at 16 is `#REF!` |
| 2 | Digital Projections (Medium Investment) | 20 | Jan 2023–Dec 2026 | Usable — real figures for 2023–2024, then `$0.00` from Jan 2025 onward |
| 3 | Digital Projections (High Investment) | 74 | 48 rows | **Unusable** — see below |

The High Investment block is corrupted. Its `Month` column runs correctly to row 97
(Jul 2024), then continues as *consecutive daily* serials — 45598, 45599, 45600 —
so rows 98 onward are days, not months. This looks like a fill-handle drag that
picked up the wrong step.

The dashboard detects the three blocks separately, excludes each Totals row, and
offers only the usable scenario in the projection selector. The other two are listed
with the reason they cannot be plotted.

**To fix:** rebuild the Low and High blocks, or delete them if they are dead.

## 2. `High-Dollar Donations` holds 22 rows, not ~500

Columns A–H carry 22 real donations (3 Aug 2026 – 24 Aug 2026, $29,500 total).
Columns I and J (`Finance Claim`, `Authentic Added`) are filled down to row 500,
and the sheet is pre-formatted to row 500, which makes the tab look roughly
twenty times larger than it is.

Any row count that anchors on column I or J will be wrong by a factor of ~22. The
dashboard anchors on `date` and `amount` instead.

## 3. Neither monthly tab contains actuals

`April 2026 Updated Goals` is titled **"Digital Goals"** in cell A3 and holds
Jan–Dec 2026 **targets**. `Digital Projections` holds scenarios. Nothing in either
tab is a record of what actually happened.

The dashboard therefore *computes* actuals from send-level rows in `Email Statistics`
and `P2P Statistics`. The consequence, stated on the panel itself:

- Actuals exist for **Email** and **SMS P2P** only.
- `Recurring`, `Ads`, `Website`, `Tandem`, `SMS Broadcast`, `Social` and `All Other`
  have **no actuals source anywhere in this workbook**.

So the "% to goal" figure compares two-channel actuals against an all-channel goal.
It is a floor, not a verdict. For example, August 2026 email raised $42,309 against
a $299,869 email goal — but a like-for-like read of total performance is not possible
from this file.

**To fix:** add a monthly actuals tab, or a channel column to the send-level tabs.

## 4. Totals rows sit one blank row below the data

`April 2026 Updated Goals` row 18, and `Digital Projections` rows 16, 70 and 124.
A single blank row separates each from its data, which is not enough to look like
the end of a table. Anything that reads "until the rows run out" will pull the
Totals row in as data and roughly double every monthly figure.

The dashboard stops at any row whose first cell reads `Total`, `Totals`, `Subtotal`
or `Sum`, and reports each one as an excluded row.

## 5. Excel errors are stored as text, not as error cells

This workbook was exported from Google Sheets, so `#VALUE!`, `#REF!` and `#N/A` are
saved as ordinary **strings** rather than as Excel error cells. Tools that check the
cell type will not see them as errors.

Live instances: `FB Audience Report` C4 (`#VALUE!`, a broken `QUERY` formula) and B14
(`#REF!`); `Digital Projections` D/H/L/M/N columns throughout (`#N/A`).

`FB Audience Report` has no header row at all and is skipped with a warning.

## 6. `Sender` has 17 distinct values

More than the 12 previously noted, and they span multiple principals:

> Cory Booker (147), Team Booker (61), Booker HQ (5), Cory Booker HQ (4),
> CoryBooker.com (3), Mark Kelly (2), CBCPAC (2), Booker for Senate (2),
> Gabby Giffords, Poll Alert, Mary Peltola, Sherrod Brown, Jon Ossoff,
> Carolyn Booker, Booker Finance Team, Finance, Team Booker, Booker HQ Finance Team

The near-duplicates — "Booker HQ" / "Cory Booker HQ" / "Booker HQ Finance Team",
and "Booker Finance Team" / "Finance, Team Booker" — are **not merged automatically**,
because they may be genuinely different senders. They are shown separately in the
Sender breakdown so the decision stays with staff.

## 7. What is clean

`Email Statistics` is in good shape and is the backbone of the dashboard: 235 rows,
7 Jan 2026 – 24 Aug 2026, every row with a valid date, and no blank or non-numeric
values in `Recipients`, `Raised`, `Donors` or `Unsubs`. All rate columns are proper
decimal fractions (highest open rate 45.6%), so nothing is double-converted.

Other row counts, for reference: `P2P Statistics` 11 rows (23 Jul – 24 Aug 2026),
`Ads Report` 3 rows (all Sep 2021, all four revenue columns blank),
`Email Sending CalendarTracker` 96, `Partner Toolkits` 24, `P2P Calendar` 10.

---

## How the dashboard is laid out

One tab per sheet in the workbook, plus two others:

- **Overview** — the figures that combine several sheets: the KPI row, raised vs goal,
  and a directory of every sheet.
- **One tab per sheet** — its charts, then a **Sheet data** table at the bottom showing
  the rows exactly as the workbook formats them, with real Excel row numbers and column
  letters, and a filter box. Sheets the parser cannot interpret (`FB Audience Report`,
  `Digital Report`, `Paid Media Report`) still get their raw cells shown, so nothing in
  the workbook is unreachable.
- **Data sources** — status, header row, row counts and warnings for all 12 sheets.

**Column coverage:** the workbook names 108 columns across its sheets. 78 appear in a
chart or a summary table; all 108 appear in a sheet-data table. The 30 that are listed
rather than charted are the ones that do not summarise usefully — `Subject Line`,
`Mailing Topic`, `Send Time (ET)`, `Label`, `Notes`, `Link to draft`, the redundant
`Year` / `Quarter` / `Month` label columns, and the projection columns that are `#N/A`
at source.

The five identity columns on `High-Dollar Donations` are the only ones absent
everywhere, by design.

Two columns are deliberately **recomputed rather than read**: donate rate comes from
`Donors ÷ Recipients` and unsub rate from `Unsubs ÷ Recipients`, so they stay weighted
by send size. The sheet's own `Donate Rate` and `Unsub Rate` columns are shown unaltered
in the sheet-data table.

Every ⓘ opens on hover; click one to keep it open while you read or scroll it.

---

## Safe to change without touching the code

The parser is built to absorb ordinary editing:

- **Append rows** to any tab — picked up automatically, no size limit.
- **Small header typos** — `Recipiants` still resolves to `Recipients`, with a
  visible warning in the Data Sources panel and in every affected ⓘ bubble.
- **Reorder columns** — matching is by name, not position (the four Ads revenue
  columns are the deliberate exception, since their names are duplicated).
- **Add columns** — listed by name in the Data Sources panel so you can see they
  are not being used.
- **Delete an optional column** — only the panels that need it go blank.
- **Rename a tab** — reported by name rather than silently ignored.

What will break a tab, by design rather than by accident: deleting the `Date`
column from `Email Statistics` or `P2P Statistics`, or the `Month` column from a
monthly tab. Those tabs fail loudly and name the missing column; every other tab
keeps working.
