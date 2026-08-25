# CAB Digital Report — Dashboard Build Plan

## Objective

Build a **single self-contained HTML file** that a campaign staffer opens in any browser, drags an Excel workbook onto, and gets a full analytics dashboard from. No backend, no accounts, no deploy step, no install.

Constraints that drive every decision below:

- **Staff have no Claude access.** The deliverable must be a plain file that works standing alone.
- **The workbook contains donor PII.** Parsing must happen entirely client-side. The file must never be uploaded anywhere.
- **This is a stepping stone to a React app.** Keep the parse/normalize layer as a clean standalone module so it ports over without a rewrite.
- **Two non-negotiable features:** full source transparency on every rendered element, and a parser that survives staff appending rows and making small header edits/typos.

---

## Tech constraints

- One output file: `dashboard.html`. HTML + CSS + JS inline. No build step, no bundler.
- **SheetJS (xlsx)** for parsing, vendored inline (not from CDN — this must work offline and on locked-down machines).
- Charting: any dependency-light library vendored inline, or hand-rolled SVG. Do **not** pull from a CDN at runtime.
- No `localStorage` / `sessionStorage` / `IndexedDB`. All state in memory for the session.
- Target: current Chrome/Edge/Firefox/Safari. Must work from `file://`.

---

## Source workbook: verified structure

Reference file: `EXTERNAL__CAB_Digital_Report.xlsx`, 12 tabs. **Header rows differ per tab** and are 1-indexed as they appear in Excel.

| Tab | Header row | Data starts | Verified rows | Role |
|---|---|---|---|---|
| `Email Statistics` | 5 | 6 | 235 (ends row 240) | **Primary** — send-level email |
| `P2P Statistics` | 5 | 6 | 11 | **Primary** — send-level P2P/SMS |
| `Ads Report - Finance-Adjusted` | 3 | 4 | 3 | Monthly ads by placement |
| `Digital Projections` | 2 | 3 | ~122 | Monthly projections |
| `April 2026 Updated Goals` | 4 | 5 | 12 | Monthly actuals + goals |
| `High-Dollar Donations` | 4 | 5 | ~495 | **PII** — donation-level |
| `Email Sending CalendarTracker` | 3 | 4 | ~95 | Ops calendar |
| `P2P Calendar` | 1 | 2 | ~22 | Ops calendar |
| `Partner Toolkits` | 1 | 2 | ~24 | Reference list |
| `Digital Report` | n/a | n/a | n/a | Formatted cover page — **do not parse** |
| `Paid Media Report` | n/a | n/a | n/a | Metric definitions — **do not parse** |
| `FB Audience Report` | none found | n/a | n/a | Currently broken (`#VALUE!` in C4) — **skip, surface as warning** |

### Column names, exactly as they appear

**`Email Statistics`** (A5:X5) — `Year`, `Quarter`, `Month`, `Day of Week`, `Date`, `Send Time (ET)`, `Ask`, `Subtype`, `Mailing Topic`, `Sender`, `Subject Line`, `Audience`, `Recipients`, `Open Rate`, `Click Rate`, `Raised`, `Donors`, `Average`, `Donate Rate`, `Actions`, `Action Rate`, `Unsubs`, `Unsub Rate`, `Label`

**`P2P Statistics`** (A5:Y5) — `Client`, `Year`, `Quarter`, `Month`, `Day of Week`, `Date`, `Goal`, `Topic`, `Audience`, `Recipients`, `Gross Spend`, `Immediate Raised`, `Immediate ROAS`, `Click Rate`, `Donors`, `Average Gift`, `Donate Rate`, `NTL`, `Gross CPA`, `LTV from NTL`, `LTV ROAS from NTL`, `Immediate Raised from NTL`, `Unsubs`, `Unsub Rate`, `Source Code`

**`Ads Report - Finance-Adjusted`** (B3:K3) — `Month`, `Placement`, `Goal`, `Gross Spend`, `NTL`, `Gross CPA`, `Lifetime Raised`, `Lifetime ROAS`, `Lifetime Raised`, `Lifetime ROAS`

> ⚠️ **Duplicate headers.** Row 1 carries merged group labels: **"Standard Toplines"** spans H:I, **"Finance-Adjusted"** spans J:K. Row 3 repeats `Lifetime Raised` / `Lifetime ROAS` under each. **Disambiguate by column position, not by name.** Map H→`lifetimeRaisedStandard`, I→`lifetimeRoasStandard`, J→`lifetimeRaisedFinanceAdj`, K→`lifetimeRoasFinanceAdj`. Note column A is blank — the table starts at B.

**`Digital Projections`** (row 2) and **`April 2026 Updated Goals`** (row 4) share an identical shape — `Month`, `Email`, `Recurring`, `Ads`, `Website`, `Tandem`, `SMS Broadcast`, `SMS P2P`, `Social`, `All Other`, **(column K blank)**, `Gross Raised`, `Gross Spend`, `Net Raised`. The Goals tab adds `Year`, `Month` in O:P (both empty — ignore). The blank column K must not shift the mapping of L/M/N.

**`High-Dollar Donations`** (A4:J4) — `Category`, `receipt_id`, `date`, `amount`, `fundraising_page`, `first`, `last`, `email`, `Finance Claim`, `Authentic Added`

> 🔒 **PII.** Columns `first`, `last`, `email`, `receipt_id`, `fundraising_page` must be **dropped during parsing** and never held in dashboard state. See PII Handling below.

### Known categorical values (for canonicalization, not validation)

- `Ask`: Cultivation, Engagement, Fundraising, GOTV, Other
- `Subtype`: For Others, General, Reactivation, Tandem, Welcome
- `Audience` (email): Actives, Actives (90-Day), Actives (120-Day), Donors, Inactives, New-to-list, Nonopeners
- `Audience` (P2P): CTD Donors, Core(y), Form Signers, GA Text, YTD Donors — **all values carry a leading space in the source; trim required**
- `Sender`: 12 variants including near-duplicates — "Booker HQ", "Cory Booker HQ", "Booker HQ Finance Team", "Finance, Team Booker", "Booker Finance Team". Do **not** auto-merge these; see Canonicalization.
- `Placement` (ads): PACtion, Protect the Vote, Really American

### Data-quality issues already confirmed

1. `Email Statistics` has `max_row` 1181 but only 235 real rows — trailing pre-formatted blank rows. **Never trust `max_row`/sheet dimensions to find the end of data.**
2. Same pattern in `P2P Statistics` (1000), `Ads Report` (1001), `Partner Toolkits` (985), `High-Dollar Donations` (500).
3. `FB Audience Report` C4 contains `#VALUE!`.
4. `Ads Report` has 3 rows, all Sep 2021, all `Sign Ups`, and **all four revenue columns are empty**. There is currently no ads revenue data.
5. `P2P Calendar` header C1 contains an embedded newline: `"Topic \n(Link copy here)"`.
6. Rate columns (`Open Rate`, `Click Rate`, etc.) are stored as **decimal fractions** (0.135 = 13.5%). Do not double-convert.

---

## Requirement 1: Provenance / transparency

Every KPI, chart, and table renders an **ⓘ info bubble**. This is a core feature, not a nice-to-have. Nothing renders without it.

### Provenance contract

Every derived value carries a provenance object alongside it. Metrics and their provenance must be produced by the *same* function — never assembled separately, or they drift.

```js
{
  tab: "Email Statistics",
  headerRow: 5,
  dataRange: "A6:X240",          // actual resolved range, not the sheet's nominal dimensions
  columnsUsed: [
    { name: "Raised",     letter: "P", role: "numerator" },
    { name: "Recipients", letter: "M", role: "denominator" }
  ],
  rowsAvailable: 235,
  rowsUsed: 47,
  rowsExcluded: [
    { reason: "outside selected date range", count: 186 },
    { reason: "Recipients blank or zero",     count: 2 }
  ],
  transform: "SUM(Raised) / SUM(Recipients) × 1000",
  notes: []                       // caveats, fuzzy-match warnings, fallback notices
}
```

### What the bubble displays

- Source tab name and header row
- Resolved cell range the figure was computed from
- Every column used, **by name and Excel letter** (`Raised` (P))
- Rows used out of rows available
- Rows excluded, grouped by reason
- The transform in plain arithmetic
- Any warnings — fuzzy-matched column names, coerced values, fallbacks

### Rules

- **Excel letters, always.** A staffer must be able to open the workbook and land on the exact column. Internal camelCase keys never appear in the UI.
- **Resolved ranges, never nominal.** `A6:X240`, not `A6:X1181`.
- **Exclusions are itemized.** "235 rows available, 47 used" without the 188-row explanation is worse than no number at all.
- **Warnings propagate upward.** If a fuzzy match happened during parsing, every chart downstream of that column shows the warning in its bubble.
- **Every exclusion needs a stated reason.** If a row is dropped and no reason string exists, that's a bug.

Add a **"Data sources" panel** listing every tab: parsed / skipped / failed, header row found, rows loaded, date range covered, and unmatched columns. This is the first thing a staffer checks when a number looks wrong.

---

## Requirement 2: Robust parsing

Assume staff will append rows, rename headers slightly, introduce typos, reorder columns, add columns, and paste text into numeric cells. The parser degrades gracefully and reports loudly; it never silently guesses.

### Finding the header row

Do not hardcode row numbers — use the table above as a **hint**, then verify. Scan the first ~15 rows and score each on: count of non-empty string cells, how many match expected names for that tab (after normalization), and whether the row below looks like data. Take the best score above a threshold.

- Header found at the expected row → proceed.
- Header found elsewhere → proceed, note the shift in provenance (`notes`), surface in Data Sources panel.
- No header found → mark tab failed, name it in the UI, continue with other tabs. **One bad tab never blocks the dashboard.**

### Finding the end of data

Walk from the first data row until **10 consecutive fully-blank rows**, then stop. Anchor the check on the tab's key column (`Date` for Email/P2P, `Month` for monthly tabs). Rows blank in the key column but populated elsewhere get skipped and counted with a reason. Never use sheet dimensions.

### Matching columns (the typo-tolerance layer)

Three-stage cascade per expected column:

1. **Exact match** after normalization: lowercase, trim, collapse internal whitespace, strip newlines/punctuation. (`"Topic \n(Link copy here)"` → `topic link copy here`.)
2. **Alias match** against a hand-maintained alias table:
   ```
   raised       ← raised, gross raised, revenue, total raised, amount raised
   recipients   ← recipients, sent, delivered, recips, quantity
   openRate     ← open rate, opens, open %, openrate, unique open rate
   clickRate    ← click rate, clicks, ctr, click %, click through rate
   unsubs       ← unsubs, unsubscribes, unsub, opt outs, optouts
   grossSpend   ← gross spend, spend, cost, gross cost
   ntl          ← ntl, new to list, new-to-list, newtolist
   ...
   ```
3. **Fuzzy match** — Levenshtein distance ≤ 2 on normalized strings, or ≥ 0.85 token-set similarity. **Any fuzzy match raises a warning** naming both the found and expected header, propagated into every dependent chart's bubble.

Then:
- **Required column missing** (e.g. `Date` in Email Statistics) → that tab fails; panels depending on it show a named empty state explaining which column is missing.
- **Optional column missing** → dependent panels only show the empty state. Everything else renders.
- **Unrecognized extra columns** → ignored for computation, but **listed by name in the Data Sources panel** so staff know a new column isn't being picked up.
- **Ads tab positional columns** (H/I/J/K) bypass name matching entirely.

### Type coercion

Every cell goes through a typed coercion that returns a value **or** a rejection reason (never a silent `NaN`):

- **Dates** — Excel serials, real datetimes, and common strings. Reject anything unparseable with a reason.
- **Numbers** — strip `$`, `,`, `%`, whitespace, parens-as-negative. **If a `%` was stripped, divide by 100** — a staffer typing `13.5%` into a rate column must not become 1350%.
- **Rates** — already fractions. Flag any value > 1 in a rate column as suspect rather than silently charting it.
- **Text** — trim (handles the P2P leading-space issue).

Every rejection is counted with its reason and shown in the relevant bubble.

### Canonicalization

- Trim and collapse whitespace on all categorical values.
- Case-insensitive grouping for near-identical labels (`actives` = `Actives`).
- **Do not auto-merge the `Sender` variants.** "Booker HQ" and "Booker HQ Finance Team" may be legitimately distinct. Show them separately; optionally offer a manual grouping toggle in the UI.

---

## Requirement 3: PII handling

`High-Dollar Donations` contains donor names, email addresses, and receipt IDs.

- **Drop `first`, `last`, `email`, `receipt_id`, `fundraising_page` at the parser boundary.** They must not enter dashboard state — not in a filtered-out field, not in a tooltip, not in a debug object.
- Keep only `Category`, `date`, `amount`, `Finance Claim`, `Authentic Added`.
- Render **aggregates only**: count and sum by category, by month, by amount band. No drill-through to individual donations.
- Amount bands: `$500–999`, `$1,000–2,499`, `$2,500–4,999`, `$5,000+`.
- Never `console.log` raw rows from this tab.
- The tab's ⓘ bubble states explicitly that identity columns were dropped at parse time and why.
- Add a visible one-line note near the panel: aggregates only, no individual donor data.

---

## Dashboard panels

Global controls: **date-range filter** (default: trailing 6 months) and channel/audience filters. Every panel respects them, and every bubble reflects the filtered `rowsUsed`.

**KPI row** — Gross Raised, Net Raised, Gross Spend, blended ROAS, % to goal. Each with its own bubble.

1. **Raised vs. goal by month** — `April 2026 Updated Goals` (actuals) vs `Digital Projections` (plan), with variance. Bubble must make clear which tab supplied which series; this is the panel most likely to be misread.
2. **Channel mix over time** — stacked area across `Email`, `Recurring`, `Ads`, `Website`, `Tandem`, `SMS Broadcast`, `SMS P2P`, `Social`, `All Other`.
3. **Email performance trend** — revenue per 1,000 recipients (primary), plus open / click / donate rates. From `Email Statistics`.
4. **List fatigue** — send cadence (sends per week) against unsub rate, same axis.
5. **Email breakdown** — performance by `Audience`, `Ask`, `Subtype`. Table with sortable columns.
6. **Ads CPA & ROAS by placement** — standard vs finance-adjusted side by side. **Ships as an empty state**: "No ads revenue data in source (3 rows found, all Sep 2021, revenue columns blank)." Lights up automatically when data appears.
7. **P2P performance** — immediate vs LTV ROAS, CPA, NTL by send. 11 rows only; label the small sample in the panel itself, not just the bubble.
8. **High-dollar aggregates** — count and volume by category and amount band. Aggregates only.

Panels 6 and 7 prove the empty-state and small-sample handling. Get those right rather than hiding thin data.

---

## Build order

**Phase 1 — Parser module.** Self-contained, no DOM dependencies (this is what ports to React). Header detection, column matching cascade, coercion, canonicalization, provenance emission. Outputs: `emailSends`, `p2pSends`, `adsMonthly`, `monthlyPlan`, `monthlyActuals`, `highDollarAgg`, plus a `parseReport`.

**Phase 2 — Provenance plumbing + Data Sources panel.** Before any chart. Wire the contract first so no panel can be built without it.

**Phase 3 — KPIs and panels 1–3.** Highest-value, densest data.

**Phase 4 — Panels 4–8**, including empty and small-sample states.

**Phase 5 — Hardening.** Wrong file dropped in, workbook with renamed tabs, workbook with zero matching tabs, tab present but empty, `#VALUE!` cells, appended rows, 3× data volume for performance.

---

## Acceptance criteria

- [ ] Single HTML file, opens from `file://`, works fully offline.
- [ ] Drag-and-drop plus a file-picker fallback.
- [ ] Workbook never leaves the browser; no network calls after load.
- [ ] Every KPI, chart, and table has a working ⓘ bubble with tab, header row, resolved range, columns by Excel letter, rows used/available, itemized exclusions, and transform.
- [ ] Data Sources panel lists all 12 tabs with status and unmatched columns.
- [ ] Appending 50 rows to `Email Statistics` is picked up with no code change.
- [ ] Renaming `Recipients` to `Recipiants` still parses, with a visible warning in the panel and in every dependent bubble.
- [ ] Deleting the `Open Rate` column degrades only its panels; everything else renders.
- [ ] Renaming a tab is reported clearly, not silently ignored.
- [ ] Trailing blank rows never inflate counts.
- [ ] No PII column value appears anywhere in state, DOM, or console.
- [ ] `#VALUE!` in FB Audience Report produces a warning, not a crash.
- [ ] Ads panel shows its explanatory empty state.
- [ ] Parser module has zero DOM/global dependencies.

---

## Open questions

1. **P2P history** — only 11 rows from Jul 23 2026 onward. Is that all the data, or does history live elsewhere?
2. **Ads data** — 3 rows from Sep 2021 with no revenue. Is a fuller export available, and should the default view stay inside 2026?
3. **`Digital Projections` vs `April 2026 Updated Goals`** — confirm the intended reading is plan vs actuals. The panel-1 labeling depends on it.
4. **Multi-client** — `P2P Statistics` has a `Client` column (currently all `CAB`) and `Sender` spans multiple principals (Mark Kelly, Jon Ossoff, Gabby Giffords). Should the dashboard filter by client, or assume single-client?