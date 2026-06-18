# SeatGeek Report

Web app: upload this month's **Invoice Details** export(s), **Purchase Details**
export(s), and the prior month's **Category List** (`.csv`, `.xlsx`, or `.xlsm`, one or
many of each). Download the report workbook with four tabs — **Summary**, **Category**,
**Invoice Details**, **Purchase Details** — plus a standalone updated **Category List**.

The only thing the app decides is the **Category** (Inventory Type) for each invoice
row. Everything on the Summary is a live Excel formula that references the Invoice
Details tab, so the workbook stays auditable and recalculates if you edit a category by
hand.

## Output workbook

- **Summary** — YTD P&L by Inventory Type (NFL, MLB, NBA, NHL, Concerts, Other, Totals).
  Built from live formulas:
  - `Sales within deal` = `SUMIF` of Total Price by Category
  - `% Sales on SG / VS / SH` = `SUMIFS` of Total Price by Client (`SeatGeek`,
    `Vivid Seats`, `StubHub`); `% Sales Other` = the remainder
  - `COGS` = `SUMIF` of Total Cost by Category; `Profit` = Sales − COGS
  - `Profit Share to SG` = `Profit × 30%`
  - `Profit Share to SG Less fees` = `(Profit − 7% × Revenue) × 30%`
  - `Capital invested by SG to date` = fixed constant
  - `Size of inventory Fund` = `SUM` of the Purchase Details Total Cost column
  Column letters in the formulas are detected from the actual headers, so the report
  still works if the export's column order changes.
- **Category** — the running master list (`Performer/Team`, `League`), sorted by category
  then performer. This is the prior list plus any performers new this month, so the tab
  (or the standalone Category List file) is what you upload next month.
- **Invoice Details** — every uploaded invoice row, cancelled rows removed, with a
  `Category` column added as column A.
- **Purchase Details** — every uploaded purchase row, passed through as-is.

A standalone **Category List** file (just the running list) is also produced for an easy,
lightweight re-upload next month.

## How a row gets its Category

Rules run in this order (first match wins):

1. **Major-league roster** — exact team name → NFL / MLB / NBA / NHL.
2. **Known stage show** (`THEATER_TITLES`, subtitle-aware) or live-event keyword
   (`SHOW_EVENT_RE` — *on ice*, *Cirque*, *Monster Jam*, *rodeo*, *NASCAR*, and
   wrestling/MMA: *WWE*, *UFC*, *AEW*, *Bellator*, *PFL*) → Other.
3. **Has an opponent** (the Performer/Opponent column is filled) → Other — a matchup
   that isn't a major league is treated as other sports (college, MLS, WNBA, etc.).
4. **TextTags**, if the export happens to include them (Broadway → Other, sports → Other).
5. Otherwise a lone act → **Concerts**.

Theatre goes to Other via the show-title list rather than by venue: the same theatres
(Fox, Orpheum, the Performing Arts Centers) host far more concerts than touring shows, so
a venue keyword would misfile thousands of concerts. Instead, the **most common venue is
shown next to each name in the review gate** as a clue for spotting theatrical shows the
title list doesn't yet know.

Concerts vs Other is the only genuinely subjective edge; the review gate and overrides
exist to manage it.

## The running Category List (your memory)

Each run, upload the prior month's **Category List** (or last month's report — the app
reads its `Category` tab). That list is **authoritative**: any performer on it keeps its
listed category, overriding the rules. Only performers *not* on the list fall to the
rules, and any of those the rules place in **Concerts** (a lone act that might be Other)
are held back on a short review list *before* the file is built — with the most common
venue shown as a clue. Confirm or change each, then generate. The output's `Category` tab
(and the standalone Category List file) is the prior list **plus** the new performers, so
it becomes next month's upload. The list grows; you never re-review a known name.

Because the list travels with the file, the app needs no server-side storage to remember
categories — the memory rides along in the upload.

Two small local files still exist as a fallback when no list is uploaded: `overrides.json`
(names you explicitly changed) and `seen.json` (names already processed). Both ship empty;
with the Category List workflow you can ignore them.

### Tuning

Open `app.py` and edit the **DEAL TERMS & CONFIG** block at the top: `PROFIT_SHARE`
(30%), `FEE_RATE` (7%), `CAPITAL_INVESTED`, `REPORT_NAME`, the `MARKETPLACES` map, the
league rosters, and `THEATER_TITLES` / `SHOW_EVENT_RE`.

## Persistence on Railway

Not required with the Category List workflow — the running list is uploaded each run, so
nothing needs to survive a redeploy. (If you ever prefer server-side memory instead,
attach a Railway Volume and set `OVERRIDES_PATH=/data/overrides.json` and
`SEEN_PATH=/data/seen.json`, but it's optional.)

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   `$PORT` is provided automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — file parsing, categorisation, review gate, workbook builder |
| `index.html` | Single-page upload UI (invoice / purchase / category list) with the review gate |
| `seed_lookup.py` | Optional: bulk-lock categories from a prior categorised report |
| `overrides.json` | Fallback human category locks (starts empty) |
| `seen.json` | Fallback processed-performer memory (starts empty) |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
