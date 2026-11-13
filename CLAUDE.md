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
| `.github/workflows/crawl-pboc.yml` | Mon + Thu 08:00 UTC and manual runs; commits `data/` to `main` |

## Site structure (as observed Nov 2026)

- Index page lists year sections (`2026年统计数据`, `2025年统计数据`, ...), each followed by
  category links (`社会融资规模`, `货币统计概览`, `金融机构信贷收支统计`, ...).
- Category pages list tables in `<tr>` rows: title (Chinese + English) plus `htm` / `xls` / `pdf`
  links pointing at `/diaochatongjisi/attachDir/YYYY/MM/<timestamp>.xls(x)`. The link text is just
  `xls`, so the **row text is the only description** of a file.
- Money Supply (`货币供应量`) lives under `货币统计概览`; AFRE flow/stock under `社会融资规模`.
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
5. **Reshape** deterministically (`reshape()`) into rows of
   `dataset, table_title, item, measure, period, value, unit, sheet, source_url, source_page, retrieved_at_utc`.
6. **Upsert** into the master on key `(dataset, item, measure, period)` (last write wins), then add
   the file URLs to `processed_urls.json`.

## Rules

- **LLM decides layout, code moves data.** Claude only returns row/column indices and labels;
  numeric values are always read by pandas from the sheet. Never ask the model to transcribe numbers.
- **Idempotent runs.** A file URL is added to `processed_urls.json` only after its data is written.
  Failed files are left out so the next run retries them. Re-running with nothing new is a no-op.
- **Write atomically.** The master is written to a temp file and renamed.
- **Be polite to pbc.gov.cn.** Keep the 2-5 s randomized delay between every page load and download;
  don't parallelize requests.
- **Periods are normalized to `YYYY-MM`** (or `YYYY` for annual). Excel may store `2026.10` as the
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
```
