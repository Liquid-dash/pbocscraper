# PBoC statistics crawler

Crawls the People's Bank of China Statistics & Analysis Department data index
(http://www.pbc.gov.cn/diaochatongjisi/116219/116319/index.html) for **Money Supply (M0/M1/M2)**
and **Aggregate Financing to the Real Economy (AFRE, flow and stock)**, and maintains a cleaned,
long-format master spreadsheet in `data/`.

## Layout

| Path | Purpose |
|---|---|
| `scraper.py` | The whole pipeline (single file, CLI entry point) |
| `data/pboc_master_data.xlsx` | Master dataset, sheet `data`, upserted each run |
| `data/processed_urls.json` | Attachment URLs already ingested (idempotency) |
| `data/translations.json` | Fixed Chinese → English label translations |
| `data/state.json` | One-time 20-year backfill status |
| `.github/workflows/crawl-pboc.yml` | Mon + Thu 08:00 UTC and manual runs; commits `data/` to `main` |

## Site structure (as observed Nov 2026)

- Index page lists year sections (`2026年统计数据`, `2025年统计数据`, ...), each followed by
  category links (`社会融资规模`, `货币统计概览`, `金融机构信贷收支统计`, ...).
- Category pages list tables in `<tr>` rows: title (Chinese + English) plus `htm` / `xls` / `pdf`
  links pointing at `/diaochatongjisi/attachDir/YYYY/MM/<timestamp>.xls(x)`. The link text is just
  `xls`, so the **row text is the only description** of a file.
  Older files (roughly pre-2019) live under `/diaochatongjisi/fileDir/resource/cms/YYYY/MM/` instead,
  so attachments are matched by `.xls(x)` extension, not path. Before 2015 there is no xls at all,
  only an Excel-exported `.htm` page (GB2312); `collect_attachments` takes the `.htm` link only for
  rows without an xls, and `html_tables()` turns it into a read_excel-style grid (merged-cell text
  only at its top-left cell, so English label rows don't pick up copied numbers). Needs `lxml`.
- Money Supply (`货币供应量`) lives under `货币统计概览`; AFRE flow/stock under `社会融资规模`
  (AFRE pages exist from 2012).
- Sheet quirks seen in older years: numbers stored as text with trailing `\xa0` (handled by
  `to_number()`), and extra tables below the notes (a growth-rate table, or a supplementary
  "adjusted" table with months down the rows in a different unit). Only the first, main table is read.
- Each year's file is cumulative (Jan..latest month) and is re-published under a **new URL** every
  month, so a new URL means new data and overlapping periods are expected.

## Pipeline

1. **Crawl** (Playwright, headless Chromium, random 2-5 s delay between requests). Only HTML
   documents are loaded (`block_subresources` aborts scripts/CSS/images) with
   `wait_until="domcontentloaded"` and 3 attempts per page: the link tables are static HTML, but the
   site's scripts take 30+ s and made `networkidle` time out on GitHub runners. `--years N` limits to the newest N year sections (default 1).
2. **Select links** with Claude (`claude-haiku-4-5`, `client.messages.parse` + Pydantic):
   `PageSelection` for category pages, `AttachmentSelection` for spreadsheets (tags each as
   `money_supply` / `afre_flow` / `afre_stock`). Returned URLs are validated against the candidate
   set; anything not in it is dropped.
3. **Download** through the browser context (`page.request.get`) so cookies/headers match.
4. **Detect layout** with Claude (`SheetLayout`): orientation (`periods_in_columns` vs
   `periods_in_rows`), header/data/footer rows, optional `measure_row` for periods spanning several
   sub-columns (e.g. Stock / Growth Rate), item-label columns, unit note, title.
   **Cross-check:** every sheet is also parsed with `HeuristicSelector`'s layout; if the model's
   layout yields under half as many values (`CROSS_CHECK_RATIO`), the heuristic's is used. (Live
   run, Nov 2026: Claude took the merged `其中 Of which` row as the header of 2015/2018/2019 flow
   sheets and kept 1 of 9 columns.) A file that still yields nothing is a failure (not marked
   processed). Attachment picks are likewise the union of Claude's and the keyword matcher's.
5. **Reshape** deterministically (`reshape()`) into rows of
   `dataset, table_title, item, measure, period, value, unit, sheet, source_url, source_page, retrieved_at_utc`.
6. **Translate** Chinese labels in `table_title, item, measure, unit` to English (`translate()`,
   `TranslationBatch`). Translations are cached in `data/translations.json` and never re-asked, so a
   label always maps to the same English (item/measure are key columns). The whole merged master is
   translated, so older Chinese rows collapse onto their English keys. After translation every
   label goes through `normalize_label()` (NFKC, `Growth Rate(%)` -> `Growth Rate (%)`, drops
   `Of which:`) and `canonicalize_labels()` (case variants collapse to the most common spelling;
   each AFRE dataset's headline series gets one name from `AFRE_TOTAL`).
7. **Upsert** into the master on key `(dataset, item, measure, period)` (last write wins), then add
   the file URLs to `processed_urls.json`.

## One-time backfill

`data/state.json` tracks a one-time history crawl (the index only goes back to 2008, so 20 sections = all available years). Until `backfill_complete` is true, live runs crawl
`BACKFILL_YEARS` (20) year sections regardless of `--years`. It is marked complete after a run with
no failures, or after `BACKFILL_MAX_ATTEMPTS` (3) attempts; after that runs crawl `--years`
(default 1, i.e. the latest year only). Delete `data/state.json` to backfill again.

The pre-2015 `/eportal/` `.htm` files 404 almost every time from GitHub runners (over both http
and https) but load fine from a normal connection, so the Nov 2026 backfill fetched them locally
with `--no-llm` (heuristic selection/layout, no API key; writes `data/`). Labels not yet in
`translations.json` stay Chinese until the next Claude run, which re-translates the whole master.

## Rules

- **LLM decides layout, code moves data.** Claude only returns row/column indices, labels and label
  translations; numeric values are always read by pandas from the sheet. Never ask the model to transcribe numbers.
- **Idempotent runs.** A file URL is added to `processed_urls.json` only after its data is written.
  Failed files are left out so the next run retries them. Re-running with nothing new is a no-op.
- **Write atomically.** The master is written to a temp file and renamed.
- **Be polite to pbc.gov.cn.** Keep the 2-5 s randomized delay between every page load and download;
  don't parallelize requests.
- **Periods are normalized to `YYYY-MM`** (`YYYY-Qn` for quarterly, e.g. 2015 AFRE stock; `YYYY` for annual). Excel may store `2026.10` as the
  float `2026.1`; `norm_period_sequence()` resolves that from the preceding month, so always
  normalize a whole header/column at once, not cell by cell.
- **`--dry-run` must work without `ANTHROPIC_API_KEY`** and must not write to `data/`. It swaps in
  `HeuristicSelector` (keyword link matching + regex layout detection) for `ClaudeSelector`.
  Keep the two selectors' method signatures identical.
- Model is fixed to `claude-haiku-4-5` (`MODEL` constant). Structured-output schemas must stay
  compatible with the API's JSON-schema subset (no numeric min/max constraints, etc.).
- Exit code is non-zero if any file failed; the workflow still commits partial progress, then fails.

## Commands

```bash
pip install -r requirements.txt
python -m playwright install chromium
python scraper.py --dry-run            # crawl + parse, no API key, no writes
python scraper.py                      # live run (needs ANTHROPIC_API_KEY)
python scraper.py --years 5            # backfill the five most recent year sections
python scraper.py --max-files 1        # cap new files this run (testing)
python scraper.py --no-llm --years 20  # live run without Claude (local fetch of files runners can't reach)
```
