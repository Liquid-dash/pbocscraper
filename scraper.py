"""PBoC (pbc.gov.cn) statistics crawler for Money Supply (M2) and Aggregate Financing (AFRE).

Pipeline:
  1. Playwright loads the Statistics & Analysis Department index page and collects the
     per-year category pages (e.g. 货币统计概览, 社会融资规模).
  2. Claude (claude-haiku-4-5, structured output) picks the relevant category pages and,
     on those pages, the relevant .xls/.xlsx attachments.
  3. Each new attachment is downloaded through the browser context; Claude detects the
     sheet layout (header/data/footer rows, orientation, unit note).
  4. Data is reshaped to long format and upserted into data/pboc_master_data.xlsx;
     processed attachment URLs are recorded in data/processed_urls.json.

Run `python scraper.py --dry-run` to exercise steps 1-4 with keyword/heuristic
substitutes for Claude and no writes to data/.
"""

from __future__ import annotations

import argparse
import io
import json
import logging
import os
import random
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal

import pandas as pd
from pydantic import BaseModel, Field

INDEX_URL = "http://www.pbc.gov.cn/diaochatongjisi/116219/116319/index.html"
STATS_PATH = "/diaochatongjisi/116219/116319/"
MODEL = "claude-haiku-4-5"

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
PROCESSED_PATH = DATA_DIR / "processed_urls.json"
MASTER_PATH = DATA_DIR / "pboc_master_data.xlsx"

KEY_COLUMNS = ["dataset", "item", "measure", "period"]
MASTER_COLUMNS = [
    "dataset", "table_title", "item", "measure", "period", "value", "unit",
    "sheet", "source_url", "source_page", "retrieved_at_utc",
]

Dataset = Literal["money_supply", "afre_flow", "afre_stock"]

log = logging.getLogger("pboc")


# --------------------------------------------------------------------------- schemas

class SelectedPage(BaseModel):
    url: str = Field(description="Exact URL copied from the candidate list")


class PageSelection(BaseModel):
    pages: list[SelectedPage]


class SelectedAttachment(BaseModel):
    url: str = Field(description="Exact .xls/.xlsx URL copied from the candidate list")
    dataset: Dataset


class AttachmentSelection(BaseModel):
    attachments: list[SelectedAttachment]


class SheetLayout(BaseModel):
    contains_data: bool = Field(description="False if the sheet has no statistical table")
    orientation: Literal["periods_in_columns", "periods_in_rows"]
    header_row: int = Field(
        description="0-based row index. periods_in_columns: the row holding period labels "
        "(e.g. 2026.01). periods_in_rows: the row holding the item/series names (prefer Chinese)."
    )
    data_start_row: int = Field(description="0-based index of the first data row")
    footer_start_row: int | None = Field(
        description="0-based index of the first footer/notes row (注/Note/source), or null if none"
    )
    measure_row: int | None = Field(
        description="periods_in_columns only: 0-based row of sub-column labels when each period "
        "spans several columns (e.g. 存量/增速, Stock/Growth Rate). Pick the row whose labels line "
        "up with the numeric columns (Chinese and English sub-header rows can be offset). "
        "Null when each period has a single column."
    )
    label_columns: list[int] = Field(
        description="periods_in_columns only: 0-based columns holding item names "
        "(may be several for indented hierarchies). Empty list otherwise."
    )
    period_column: int | None = Field(
        description="periods_in_rows only: 0-based column holding period labels. Null otherwise."
    )
    unit_note: str | None = Field(description="Unit text, e.g. '亿元人民币 / 100 Million Yuan'")
    table_title: str | None


# --------------------------------------------------------------------------- helpers

@dataclass
class Link:
    text: str
    url: str
    context: str = ""  # year heading for category pages, row title for attachments


def polite_delay() -> None:
    time.sleep(random.uniform(2.0, 5.0))


def norm_period(value) -> str | None:
    """'2026.01' / '2026.1' / '2026年1月' -> '2026-01'; '2026' -> '2026'.

    A numeric 2026.1 is ambiguous (Jan or Oct once Excel drops the trailing zero);
    it maps to January here and norm_period_sequence() corrects it from context.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return f"{value.year}-{value.month:02d}"
    s = str(value).strip()
    m = re.match(r"^(\d{4})\s*[.\-/年]\s*(\d{1,2})(?:\s*月)?$", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r"^(\d{4})(?:\.0)?\s*年?$", s)
    if m:
        return m.group(1)
    return None


def norm_period_sequence(values: list) -> list[str | None]:
    """norm_period over an ordered run of period labels, fixing numeric x.1 that means October."""
    out = [norm_period(v) for v in values]
    prev = None
    for i, (raw, p) in enumerate(zip(values, out)):
        if p is None:
            continue
        if isinstance(raw, float) and p.endswith("-01") and prev == f"{p[:4]}-09":
            out[i] = p = f"{p[:4]}-10"
        prev = p
    return out


def clean_label(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = re.sub(r"\s+", " ", str(value)).strip()
    return s or None


def sheet_to_grid_text(df: pd.DataFrame, head: int = 80, tail: int = 40) -> str:
    """Compact text rendering of non-empty cells, for the layout prompt."""
    idx = list(range(len(df)))
    if len(idx) > head + tail:
        idx = idx[:head] + idx[-tail:]
    lines, prev = [], -1
    for r in idx:
        if prev != -1 and r != prev + 1:
            lines.append("... (rows omitted) ...")
        prev = r
        cells = [f"c{c}={str(v)[:80]!r}" for c, v in enumerate(df.iloc[r].tolist())
                 if not (isinstance(v, float) and pd.isna(v))]
        lines.append(f"r{r}: " + (", ".join(cells) if cells else "<empty>"))
    return "\n".join(lines)


# --------------------------------------------------------------------------- state

def load_processed() -> set[str]:
    if PROCESSED_PATH.exists():
        return set(json.loads(PROCESSED_PATH.read_text(encoding="utf-8")))
    return set()


def save_processed(urls: set[str]) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    PROCESSED_PATH.write_text(json.dumps(sorted(urls), ensure_ascii=False, indent=2) + "\n",
                              encoding="utf-8")


def upsert_master(new: pd.DataFrame) -> int:
    DATA_DIR.mkdir(exist_ok=True)
    if MASTER_PATH.exists():
        old = pd.read_excel(MASTER_PATH, dtype={"period": str})
        merged = pd.concat([old, new], ignore_index=True)
    else:
        merged = new
    merged = (merged.drop_duplicates(subset=KEY_COLUMNS, keep="last")
                    .sort_values(KEY_COLUMNS)
                    .reset_index(drop=True))[MASTER_COLUMNS]
    tmp = MASTER_PATH.with_suffix(".tmp.xlsx")
    merged.to_excel(tmp, index=False, sheet_name="data")
    tmp.replace(MASTER_PATH)
    return len(merged)


# --------------------------------------------------------------------------- crawling

LINK_JS = """els => els.map(e => {
  const tr = e.closest('tr');
  return [e.innerText.trim(), e.href, tr ? tr.innerText.replace(/\\s+/g, ' ').trim() : ''];
})"""


def goto(page, url: str) -> None:
    resp = page.goto(url, wait_until="networkidle", timeout=90_000)
    if resp is None or not resp.ok:
        raise RuntimeError(f"GET {url} failed: {resp.status if resp else 'no response'}")


def collect_category_pages(page, years: int) -> list[Link]:
    """Category links under the latest `years` year headings on the index page."""
    goto(page, INDEX_URL)
    raw = page.eval_on_selector_all("a[href]", LINK_JS)
    out: list[Link] = []
    current_year: int | None = None
    year_seen: list[int] = []
    for text, href, _ in raw:
        if STATS_PATH not in href or not href.endswith("index.html"):
            continue
        m = re.match(r"^(\d{4})年统计数据$", text)
        if m:
            current_year = int(m.group(1))
            year_seen.append(current_year)
            continue
        if current_year is not None and text:
            out.append(Link(text=text, url=href, context=f"{current_year}年"))
    if not year_seen:
        raise RuntimeError("No year sections found on index page; layout may have changed")
    keep = {f"{y}年" for y in sorted(set(year_seen), reverse=True)[:years]}
    return [l for l in out if l.context in keep]


def collect_attachments(page, category: Link) -> list[Link]:
    goto(page, category.url)
    raw = page.eval_on_selector_all("a[href*='attachDir']", LINK_JS)
    return [Link(text=t, url=h, context=row) for t, h, row in raw
            if re.search(r"\.xlsx?$", h, re.I)]


def download(page, url: str) -> bytes:
    resp = page.request.get(url, timeout=90_000)
    if not resp.ok:
        raise RuntimeError(f"Download {url} failed: HTTP {resp.status}")
    return resp.body()


# --------------------------------------------------------------------------- Claude

class ClaudeSelector:
    def __init__(self) -> None:
        import anthropic
        self.anthropic = anthropic
        self.client = anthropic.Anthropic(max_retries=4)

    def _parse(self, prompt: str, schema: type[BaseModel]):
        resp = self.client.messages.parse(
            model=MODEL,
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
            output_format=schema,
        )
        if resp.stop_reason == "refusal" or resp.parsed_output is None:
            raise RuntimeError(f"No structured output (stop_reason={resp.stop_reason})")
        return resp.parsed_output

    def pick_pages(self, links: list[Link]) -> list[Link]:
        listing = "\n".join(f"- [{l.context}] {l.text}: {l.url}" for l in links)
        prompt = (
            "These are category pages from the People's Bank of China statistics site.\n"
            "Select every page likely to contain the Money Supply table (货币供应量, M0/M1/M2; "
            "usually under 货币统计概览) or Aggregate Financing to the Real Economy tables "
            "(社会融资规模, flow and stock). Exclude everything else.\n\n"
            f"Candidates:\n{listing}"
        )
        allowed = {l.url: l for l in links}
        picked = self._parse(prompt, PageSelection).pages
        return [allowed[p.url] for p in picked if p.url in allowed]

    def pick_attachments(self, links: list[Link]) -> list[tuple[Link, str]]:
        listing = "\n".join(f"- row title: {l.context!r} | link text: {l.text!r} | {l.url}"
                            for l in links)
        prompt = (
            "These are spreadsheet download links from PBoC statistics pages, each with the "
            "title of the table row it sits in. Select only:\n"
            "- money_supply: the Money Supply table (货币供应量 / M2, M1, M0)\n"
            "- afre_flow: Aggregate Financing to the Real Economy flow (社会融资规模增量)\n"
            "- afre_stock: Aggregate Financing to the Real Economy stock (社会融资规模存量)\n"
            "Ignore balance sheets, reserves, exchange rates, templates, calendars and anything "
            "else, including regional or quarterly AFRE breakdowns.\n\n"
            f"Candidates:\n{listing}"
        )
        allowed = {l.url: l for l in links}
        picked = self._parse(prompt, AttachmentSelection).attachments
        return [(allowed[a.url], a.dataset) for a in picked if a.url in allowed]

    def detect_layout(self, df: pd.DataFrame, title_hint: str) -> SheetLayout:
        prompt = (
            "Below is a spreadsheet sheet from the People's Bank of China, rendered as "
            "non-empty cells per row (rN = 0-based row, cN = 0-based column). "
            f"The download link sat next to the title: {title_hint!r}.\n\n"
            "Identify the table layout: where the header row is, where data starts, where "
            "footer notes (注/Note/data sources) start, whether periods (e.g. 2026.01) run "
            "across columns or down rows, which columns hold item names or periods, the unit "
            "note (单位/Unit) and the table title. Bilingual tables often put an English "
            "label on the row below each Chinese label; those English-only rows are not "
            "separate data rows.\n\n"
            f"{sheet_to_grid_text(df)}"
        )
        return self._parse(prompt, SheetLayout)


class HeuristicSelector:
    """Keyword/regex stand-in used by --dry-run (no API key needed)."""

    def pick_pages(self, links: list[Link]) -> list[Link]:
        return [l for l in links if re.search(r"货币统计概览|社会融资规模", l.text)]

    def pick_attachments(self, links: list[Link]) -> list[tuple[Link, str]]:
        out = []
        for l in links:
            if re.search(r"货币供应量|Money Supply", l.context):
                out.append((l, "money_supply"))
            elif re.search(r"社会融资规模增量", l.context):
                out.append((l, "afre_flow"))
            elif re.search(r"社会融资规模存量", l.context):
                out.append((l, "afre_stock"))
        return out

    def detect_layout(self, df: pd.DataFrame, title_hint: str) -> SheetLayout:
        is_period = df.map(lambda v: norm_period(v) is not None and len(norm_period(v)) == 7)
        row_counts, col_counts = is_period.sum(axis=1), is_period.sum(axis=0)
        footer = next((r for r in range(len(df))
                       if re.match(r"^\s*(注|Note)", str(df.iat[r, 0]))), None)
        unit = next((clean_label(v) for v in df.to_numpy().ravel()
                     if re.match(r"^\s*(单位|Unit)", str(v))), None)
        if row_counts.max() >= col_counts.max():
            hr = int(row_counts.idxmax())
            first_period_col = int(is_period.iloc[hr].to_numpy().argmax())
            label_cols = list(range(first_period_col))
            value_cols = df.columns[first_period_col:]
            # First row whose label columns hold text and value columns hold numbers.
            numeric = df[value_cols].apply(pd.to_numeric, errors="coerce").notna().any(axis=1)
            data_start = next((r for r in range(hr + 1, len(df)) if numeric.iat[r]), hr + 1)
            # Sub-column labels: the populated header row nearest the data, if any.
            measure_row = None
            if df.iloc[hr][value_cols].isna().any():
                between = range(data_start - 1, hr, -1)
                measure_row = next((r for r in between if df.iloc[r][value_cols].notna().sum()
                                    >= len(value_cols) - 1), None)
            return SheetLayout(contains_data=True, orientation="periods_in_columns",
                               header_row=hr, data_start_row=data_start, footer_start_row=footer,
                               measure_row=measure_row, label_columns=label_cols,
                               period_column=None, unit_note=unit, table_title=title_hint)
        pc = int(col_counts.idxmax())
        first = int(is_period.iloc[:, pc].to_numpy().argmax())
        # Item header: nearest row above data with labels in most value columns.
        hr = first - 1
        for r in range(first - 1, -1, -1):
            if df.iloc[r].drop(df.columns[pc]).notna().sum() >= (df.shape[1] - 1) / 2:
                hr = r
                break
        return SheetLayout(contains_data=True, orientation="periods_in_rows", header_row=hr,
                           data_start_row=first, footer_start_row=footer, measure_row=None,
                           label_columns=[], period_column=pc, unit_note=unit,
                           table_title=title_hint)


# --------------------------------------------------------------------------- reshaping

def reshape(df: pd.DataFrame, layout: SheetLayout) -> pd.DataFrame:
    """Return long-format rows with columns item, measure, period, value."""
    end = layout.footer_start_row if layout.footer_start_row is not None else len(df)
    end = max(min(end, len(df)), layout.data_start_row)
    rows = []
    if layout.orientation == "periods_in_columns":
        label_cols = set(layout.label_columns)
        header = df.iloc[layout.header_row].tolist()
        value_cols = [c for c in range(df.shape[1]) if c not in label_cols]
        seq = norm_period_sequence([header[c] for c in value_cols])
        col_period: dict[int, str] = {}
        col_measure: dict[int, str] = {}
        current = None
        for c, p in zip(value_cols, seq):
            if p:
                current = p
            elif layout.measure_row is None:
                continue  # no merged headers: blank header means not a data column
            if current is None:
                continue
            if layout.measure_row is not None:
                m = clean_label(df.iat[layout.measure_row, c])
                if not m:
                    continue
                col_measure[c] = m
            else:
                col_measure[c] = "value"
            col_period[c] = current
        for r in range(layout.data_start_row, end):
            labels = [clean_label(df.iat[r, c]) for c in layout.label_columns if c < df.shape[1]]
            item = next((l for l in labels if l), None)
            if not item:
                continue
            for c, p in col_period.items():
                rows.append((item, col_measure[c], p, df.iat[r, c]))
    else:
        pc = layout.period_column if layout.period_column is not None else 0
        items = {c: clean_label(v) for c, v in enumerate(df.iloc[layout.header_row].tolist())}
        items = {c: i for c, i in items.items() if i and c != pc}
        data_rows = list(range(layout.data_start_row, end))
        periods = norm_period_sequence([df.iat[r, pc] for r in data_rows])
        for r, p in zip(data_rows, periods):
            if not p:
                continue
            for c, item in items.items():
                rows.append((item, "value", p, df.iat[r, c]))
    out = pd.DataFrame(rows, columns=["item", "measure", "period", "value"])
    out["value"] = pd.to_numeric(out["value"], errors="coerce")
    return out.dropna(subset=["value"]).reset_index(drop=True)


def process_file(selector, content: bytes, link: Link, dataset: str,
                 source_page: str) -> pd.DataFrame:
    frames = []
    sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    for name, df in sheets.items():
        if df.dropna(how="all").empty:
            continue
        layout = selector.detect_layout(df, link.context)
        if not layout.contains_data:
            log.info("  sheet %s: no data table, skipped", name)
            continue
        long = reshape(df, layout)
        log.info("  sheet %s: %s, header r%d, measure r%s, data r%d-%s, unit=%r -> %d values",
                 name, layout.orientation, layout.header_row, layout.measure_row,
                 layout.data_start_row, layout.footer_start_row, layout.unit_note, len(long))
        dupes = long.duplicated(subset=["item", "measure", "period"]).sum()
        if dupes:
            log.warning("  sheet %s: %d duplicate (item, measure, period) keys; last wins",
                        name, dupes)
        long["dataset"] = dataset
        long["table_title"] = layout.table_title or link.context
        long["unit"] = layout.unit_note
        long["sheet"] = name
        long["source_url"] = link.url
        long["source_page"] = source_page
        frames.append(long)
    if not frames:
        return pd.DataFrame(columns=MASTER_COLUMNS)
    return pd.concat(frames, ignore_index=True)


# --------------------------------------------------------------------------- main

def run(dry_run: bool, years: int, max_files: int | None) -> int:
    from playwright.sync_api import sync_playwright

    if dry_run:
        selector = HeuristicSelector()
        log.info("DRY RUN: keyword/heuristic selection, nothing written to %s", DATA_DIR)
    else:
        if not os.environ.get("ANTHROPIC_API_KEY"):
            log.error("ANTHROPIC_API_KEY is not set (use --dry-run to test without it)")
            return 2
        selector = ClaudeSelector()

    processed = load_processed()
    retrieved_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    frames: list[pd.DataFrame] = []
    done_urls: list[str] = []
    failures = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(locale="zh-CN")
        page = context.new_page()
        try:
            categories = collect_category_pages(page, years)
            log.info("Index: %d category pages in latest %d year(s)", len(categories), years)
            pages = selector.pick_pages(categories)
            log.info("Selected pages: %s", [f"{l.context}{l.text}" for l in pages])

            todo: list[tuple[Link, str, str]] = []
            for cat in pages:
                polite_delay()
                atts = collect_attachments(page, cat)
                chosen = selector.pick_attachments(atts) if atts else []
                for link, dataset in chosen:
                    log.info("  %s %s -> %s%s", cat.context, link.context[:50], dataset,
                             " (already processed)" if link.url in processed else "")
                    if link.url not in processed:
                        todo.append((link, dataset, cat.url))
            if max_files is not None:
                todo = todo[:max_files]
            log.info("%d new file(s) to process", len(todo))

            for link, dataset, source_page in todo:
                polite_delay()
                log.info("Downloading %s", link.url)
                try:
                    content = download(page, link.url)
                    df = process_file(selector, content, link, dataset, source_page)
                except Exception as exc:  # keep going; file stays unprocessed for next run
                    failures += 1
                    log.exception("Failed on %s: %s", link.url, exc)
                    continue
                df["retrieved_at_utc"] = retrieved_at
                frames.append(df)
                done_urls.append(link.url)
        finally:
            browser.close()

    new = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=MASTER_COLUMNS)
    if dry_run:
        log.info("DRY RUN result: %d records from %d file(s), %d failure(s)",
                 len(new), len(done_urls), failures)
        if not new.empty:
            with pd.option_context("display.width", 200, "display.max_columns", 10):
                print(new[["dataset", "item", "measure", "period", "value", "unit"]]
                      .groupby("dataset").head(4).to_string(index=False))
                print(new.groupby("dataset")["period"].agg(["min", "max", "count"]))
        return 1 if failures else 0

    if not new.empty:
        total = upsert_master(new[MASTER_COLUMNS])
        log.info("Upserted %d records; master now has %d rows", len(new), total)
    if done_urls:
        save_processed(processed | set(done_urls))
    log.info("Done: %d file(s) processed, %d failure(s)", len(done_urls), failures)
    return 1 if failures else 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--dry-run", action="store_true",
                    help="Test crawl + parse without the Claude API and without writing data/")
    ap.add_argument("--years", type=int, default=1,
                    help="How many of the most recent year sections to crawl (default 1)")
    ap.add_argument("--max-files", type=int, default=None,
                    help="Cap on new files processed this run")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    sys.exit(run(args.dry_run, args.years, args.max_files))


if __name__ == "__main__":
    main()
