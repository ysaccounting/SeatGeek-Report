# SG2 Report

Web app: upload this month's **Invoice Details** export(s) and **Purchase Details**
export(s) (`.csv`, `.xlsx`, or `.xlsm`, one or many of each) and download the SG2
report workbook with three tabs — **Summary**, **Invoice Details**, **Purchase Details**.

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
- **Invoice Details** — every uploaded invoice row, cancelled rows removed, with a
  `Category` column added as column A.
- **Purchase Details** — every uploaded purchase row, passed through as-is.

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

## Review gate, overrides, and "seen" memory

- When you generate, any **new** performer the rules placed in **Concerts** (a lone act
  that could arguably be Other) is held back on a short review list *before* the final
  file is built. Confirm or change each one, then generate.
- A name you **change** is saved to `overrides.json` and that choice wins forever after.
  A name you **accept as-is** stays rule-governed (so future rule tweaks still apply to
  it) and is simply remembered so it isn't flagged again.
- `seen.json` ships pre-seeded from the April report, so the first month only reviews
  names that are new since April. League, stage-show, and opponent calls are confident
  and never block.

### Tuning

Open `app.py` and edit the **DEAL TERMS & CONFIG** block at the top: `PROFIT_SHARE`
(30%), `FEE_RATE` (7%), `CAPITAL_INVESTED`, the `MARKETPLACES` map, the league rosters,
and `THEATER_TITLES` / `SHOW_EVENT_RE`. To bulk-lock categories from a historical report
instead of reviewing them, run `python seed_lookup.py "Some_Prior_Report.xlsx"`.

## Persistence on Railway

Railway's filesystem resets on every redeploy. To keep `overrides.json` and `seen.json`
across deploys, attach a **Railway Volume** and point the app at it:

```
OVERRIDES_PATH=/data/overrides.json
SEEN_PATH=/data/seen.json
```

(Copy the bundled `overrides.json` / `seen.json` into the volume once to seed it.)

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   `$PORT` is provided automatically. Add a volume + the env vars above for persistence.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — file parsing, categorisation, review gate, workbook builder |
| `index.html` | Single-page upload UI with the review gate |
| `seed_lookup.py` | Optional: bulk-lock categories from a prior categorised report |
| `overrides.json` | Human category locks (starts empty) |
| `seen.json` | Performers already processed (pre-seeded from April) |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
