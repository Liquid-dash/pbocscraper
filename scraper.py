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
import unicodedata
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal

import pandas as pd
from pydantic import BaseModel, Field

INDEX_URL = "https://www.pbc.gov.cn/diaochatongjisi/116219/116319/index.html"
STATS_PATH = "/diaochatongjisi/116219/116319/"
MODEL = "claude-haiku-4-5"

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
PROCESSED_PATH = DATA_DIR / "processed_urls.json"
MASTER_PATH = DATA_DIR / "pboc_master_data.xlsx"
TRANSLATIONS_PATH = DATA_DIR / "translations.json"
WEB_JSON_PATH = DATA_DIR / "pboc_data.json"   # public, read by liquiditydashboard.net
WEB_CSV_PATH = DATA_DIR / "pboc_data.csv"
STATE_PATH = DATA_DIR / "state.json"

BACKFILL_YEARS = 20       # one-time history crawl on the first live run
BACKFILL_MAX_ATTEMPTS = 3  # stop re-running the backfill if some old files keep failing
TRANSLATE_BATCH = 100
CROSS_CHECK_RATIO = 0.5    # use the heuristic layout if the model's yields < half its values

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


class Translation(BaseModel):
    source: str = Field(description="The label exactly as given")
    english: str = Field(description="English-only label")


class TranslationBatch(BaseModel):
    translations: list[Translation]


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
    m = re.match(r"^(\d{4})\s*[.\-/年]?\s*(?:Q|第)\s*([1-4一二三四])\s*(?:季度)?$", s, re.I)
    if m:
        q = "1234"["一二三四".index(m.group(2))] if m.group(2) in "一二三四" else m.group(2)
        return f"{m.group(1)}-Q{q}"
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


def to_number(values) -> pd.Series:
    """Numeric conversion tolerant of text cells like '2431022.72\\xa0' or '1,234.5'."""
    s = pd.Series(values)
    s = s.map(lambda v: re.sub(r"[\s,]", "", v) if isinstance(v, str) else v)
    return pd.to_numeric(s, errors="coerce")


def html_tables(content: bytes) -> dict[str, pd.DataFrame]:
    """Grid DataFrames from an HTML page's <table>s, laid out like read_excel(header=None):
    a merged (colspan/rowspan) cell's text sits only at its top-left position. Older years
    (pre-2015) publish tables only as Excel-exported .htm pages, usually GB2312-encoded."""
    from lxml import html as lxml_html
    m = re.search(rb"charset\s*=\s*[\"']?([\w-]+)", content[:4096], re.I)
    charset = (m.group(1).decode() if m else "utf-8").lower()
    if charset in ("gb2312", "gbk", "gb18030"):
        charset = "gb18030"
    doc = lxml_html.fromstring(content.decode(charset, errors="replace"))
    out: dict[str, pd.DataFrame] = {}
    for t_idx, table in enumerate(doc.iter("table")):
        cells: dict[tuple[int, int], str] = {}
        taken: set[tuple[int, int]] = set()
        n_rows = n_cols = 0
        for r, tr in enumerate(table.xpath("./tr|./thead/tr|./tbody/tr")):
            c = 0
            for td in tr.xpath("./td|./th"):
                while (r, c) in taken:
                    c += 1
                cs, rs = int(td.get("colspan", 1) or 1), int(td.get("rowspan", 1) or 1)
                taken.update((r + i, c + j) for i in range(rs) for j in range(cs))
                text = re.sub(r"\s+", " ", td.text_content()).strip()
                if text:
                    cells[(r, c)] = text
                c += cs
                n_cols = max(n_cols, c)
            n_rows = r + 1
        if n_rows < 3 or not cells:
            continue
        grid = [[cells.get((r, c)) for c in range(n_cols)] for r in range(n_rows)]
        out[f"table{t_idx}"] = pd.DataFrame(grid, dtype=object)
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


def load_json(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default


def save_json(path: Path, obj) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
                   encoding="utf-8")
    tmp.replace(path)


CJK = re.compile(r"[㐀-鿿＀-￯]")
TRANSLATED_COLUMNS = ["table_title", "item", "measure", "unit"]


def translate_columns(df: pd.DataFrame, selector, cache: dict[str, str]) -> pd.DataFrame:
    """Replace Chinese labels with English. `cache` (data/translations.json) keeps every
    label's translation fixed across runs, since item/measure are part of the upsert key."""
    labels = sorted({v for c in TRANSLATED_COLUMNS if c in df
                     for v in df[c].dropna().astype(str) if CJK.search(v)})
    missing = [l for l in labels if l not in cache]
    for i in range(0, len(missing), TRANSLATE_BATCH):
        cache.update(selector.translate(missing[i:i + TRANSLATE_BATCH],
                                        sorted(set(cache.values()))))
    still = [l for l in missing if l not in cache]
    if still:
        log.warning("%d label(s) left untranslated, e.g. %s", len(still), still[:5])
    df = df.copy()
    for c in TRANSLATED_COLUMNS:
        if c in df:
            df[c] = df[c].map(lambda v: normalize_label(cache.get(v, v))
                              if isinstance(v, str) else v)
    return canonicalize_labels(df)


AFRE_TOTAL = {"afre_flow": "Aggregate Financing to the Real Economy (Flow)",
              "afre_stock": "Aggregate Financing to the Real Economy (Stock)"}


def normalize_label(s: str) -> str:
    """Deterministic cleanup so equivalent labels compare equal: full-width -> ASCII,
    'Growth Rate(%)' -> 'Growth Rate (%)', whitespace, and 'Of which:' prefixes."""
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\s*\(\s*", " (", s)
    s = re.sub(r"\s*\)", ")", s).strip()
    stripped = re.sub(r"^(of which|其中)\s*[:：,]?\s*", "", s, flags=re.I)
    return stripped or s


def canonicalize_labels(df: pd.DataFrame) -> pd.DataFrame:
    """Labels that differ only in case collapse to their most common spelling, and each
    AFRE dataset's headline series gets a single name (years label it differently)."""
    if "item" in df and "dataset" in df:
        is_total = df["item"].astype(str).str.match(
            r"(?i)^(aggregate financing to the real economy|afre)\b")
        for ds, name in AFRE_TOTAL.items():
            df.loc[is_total & (df["dataset"] == ds), "item"] = name
    for c in TRANSLATED_COLUMNS:
        if c not in df:
            continue
        counts = df[c].dropna().astype(str).value_counts()
        best: dict[str, str] = {}
        for label in sorted(counts.index, key=lambda l: (-counts[l], l)):
            best.setdefault(label.casefold(), label)
        df[c] = df[c].map(lambda v: best.get(v.casefold(), v) if isinstance(v, str) else v)
    return df


def merge_with_master(new: pd.DataFrame) -> pd.DataFrame:
    if MASTER_PATH.exists():
        old = pd.read_excel(MASTER_PATH, dtype={"period": str})
        return pd.concat([old, new], ignore_index=True)
    return new


def prepare_master(merged: pd.DataFrame) -> pd.DataFrame:
    """The master exactly as it will be written: one row per key (last wins), sorted."""
    return (merged.drop_duplicates(subset=KEY_COLUMNS, keep="last")
                  .sort_values(KEY_COLUMNS)
                  .reset_index(drop=True))[MASTER_COLUMNS]


def same_content(a: pd.DataFrame, b: pd.DataFrame) -> bool:
    """Content comparison that ignores dtype noise from the Excel round trip (e.g. a value
    column read back as int64, or str vs object columns)."""
    if a.shape != b.shape:
        return False
    a, b = a.reset_index(drop=True), b.reset_index(drop=True)
    for c in MASTER_COLUMNS:
        if c == "value":
            x, y = pd.to_numeric(a[c], errors="coerce"), pd.to_numeric(b[c], errors="coerce")
            if not ((x - y).abs().le(1e-9) | (x.isna() & y.isna())).all():
                return False
        elif not a[c].fillna("").astype(str).equals(b[c].fillna("").astype(str)):
            return False
    return True


def web_exports(final: pd.DataFrame) -> dict[Path, str]:
    """Web-friendly copies of the master for liquiditydashboard.net (served from the public
    repo via raw.githubusercontent.com). Deterministic text, so unchanged data gives
    byte-identical files and no commit."""
    df = final.copy()
    # float + rounding: identical text whether the frame came from memory or an Excel read
    # (which can return whole numbers as int64), and drops float noise from the source sheets.
    df["value"] = pd.to_numeric(df["value"], errors="coerce").astype(float).round(6)
    df = df.fillna({c: "" for c in MASTER_COLUMNS if c != "value"})
    series = []
    for (ds, item, measure), g in df.groupby(["dataset", "item", "measure"], sort=True):
        g = g.sort_values("period")
        units = sorted(g["unit"].dropna().astype(str).unique())
        series.append({
            "id": "/".join(re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
                           for s in (ds, item, measure)),
            "dataset": ds, "item": item, "measure": measure,
            "unit": units[0] if len(units) == 1 else units,
            "first": g["period"].iloc[0], "last": g["period"].iloc[-1],
            "data": [[p, None if pd.isna(v) else float(v)]
                     for p, v in zip(g["period"], g["value"])],
        })
    payload = {
        "source": "People's Bank of China, Statistics and Analysis Department",
        "source_index": INDEX_URL,
        "updated_at_utc": str(df["retrieved_at_utc"].dropna().max()),
        "datasets": {"money_supply": "Money supply (M0, M1, M2)",
                     "afre_flow": "Aggregate Financing to the Real Economy, flow (monthly increment)",
                     "afre_stock": "Aggregate Financing to the Real Economy, stock (outstanding)"},
        "series": series,
    }
    out = {WEB_JSON_PATH: json.dumps(payload, ensure_ascii=False, indent=1) + "\n",
           # Data columns only: the repeated source URLs would triple the size of a file
           # that is re-committed every month (provenance stays in the xlsx).
           WEB_CSV_PATH: df[["dataset", "item", "measure", "period", "value", "unit"]]
                         .to_csv(index=False, lineterminator="\n")}
    return out


def write_web_exports(final: pd.DataFrame) -> list[str]:
    """Write the web files that differ from what's on disk; returns the names written."""
    written = []
    for path, text in web_exports(final).items():
        if path.exists() and path.read_text(encoding="utf-8") == text:
            continue
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(text, encoding="utf-8", newline="\n")
        tmp.replace(path)
        written.append(path.name)
    return written


def write_master(merged: pd.DataFrame) -> int:
    DATA_DIR.mkdir(exist_ok=True)
    merged = prepare_master(merged)
    tmp = MASTER_PATH.with_suffix(".tmp.xlsx")
    merged.to_excel(tmp, index=False, sheet_name="data")
    tmp.replace(MASTER_PATH)
    return len(merged)


# --------------------------------------------------------------------------- crawling

LINK_JS = """els => els.map(e => {
  const tr = e.closest('tr');
  return [e.innerText.trim(), e.href, tr ? tr.innerText.replace(/\\s+/g, ' ').trim() : ''];
})"""


GOTO_ATTEMPTS = 3


def block_subresources(route) -> None:
    """Load only HTML documents. The link tables are static HTML, while the site's
    scripts/CSS take 30+ s to load (long enough to time out from GitHub runners)."""
    if route.request.resource_type == "document":
        route.continue_()
    else:
        route.abort()


def goto(page, url: str) -> None:
    for attempt in range(1, GOTO_ATTEMPTS + 1):
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            if resp is None or not resp.ok:
                raise RuntimeError(f"GET {url} failed: {resp.status if resp else 'no response'}")
            return
        except Exception as exc:
            if attempt == GOTO_ATTEMPTS:
                raise
            log.warning("GET %s failed (attempt %d/%d): %s", url, attempt, GOTO_ATTEMPTS,
                        str(exc).splitlines()[0])
            time.sleep(10 * attempt)


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
    # Recent files live under /attachDir/, older ones (pre-2019) under /fileDir/resource/cms/.
    # Before 2015 tables are only published as .htm pages; use one only when its row has no
    # .xls(x) link.
    raw = [(t, https(h), row) for t, h, row in page.eval_on_selector_all("a[href]", LINK_JS)]
    sheets = [Link(text=t, url=h, context=row) for t, h, row in raw
              if re.search(r"\.xlsx?$", h, re.I)]
    rows_with_sheet = {l.context for l in sheets}
    pages = [Link(text=t, url=h, context=row) for t, h, row in raw
             if re.search(r"/fileDir/.+\.htm$", h, re.I) and row not in rows_with_sheet]
    return sheets + pages


def https(url: str) -> str:
    """Canonical https:// form of a pbc.gov.cn URL, so processed_urls.json has no http/https
    duplicates. download() still falls back to http:// when https:// 404s."""
    return re.sub(r"^http://(www\.)?pbc\.gov\.cn/", "https://www.pbc.gov.cn/", url)


def download(page, url: str) -> bytes:
    """The old /eportal/ files 404 over http:// from some places/times and over https://
    from others (seen both ways on GitHub runners and locally), so each attempt tries both
    schemes, and a plain browser navigation is the last resort."""
    alt = re.sub(r"^https://", "http://", url) if url.startswith("https://") else https(url)
    status = None
    for attempt in range(1, GOTO_ATTEMPTS + 1):
        for candidate in dict.fromkeys([url, alt]):
            resp = page.request.get(candidate, timeout=90_000)
            if resp.ok:
                return resp.body()
            status = resp.status
        log.warning("Download %s: HTTP %s over http and https (attempt %d/%d)", url, status,
                    attempt, GOTO_ATTEMPTS)
        if attempt < GOTO_ATTEMPTS:
            time.sleep(10 * attempt)
    for candidate in dict.fromkeys([url, alt]):
        try:
            resp = page.goto(candidate, wait_until="domcontentloaded", timeout=60_000)
            if resp is not None and resp.ok:
                return resp.body()
        except Exception as exc:
            log.warning("Download %s via navigation failed: %s", candidate,
                        str(exc).splitlines()[0])
    raise RuntimeError(f"Download {url} failed: HTTP {status}")


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
            "else, including regional or quarterly AFRE breakdowns. Older years title the "
            "flow table just 社会融资规模统计表 (no 增量); that is afre_flow.\n\n"
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
            "separate data rows. If the sheet holds more than one table (e.g. a supplementary "
            "'adjusted' table or a growth-rate table below the notes), describe only the first, "
            "main table and set footer_start_row to where it ends.\n\n"
            f"{sheet_to_grid_text(df)}"
        )
        return self._parse(prompt, SheetLayout)

    def translate(self, labels: list[str], known: list[str]) -> dict[str, str]:
        listing = "\n".join(f"- {l}" for l in labels)
        vocab = "\n".join(f"- {k}" for k in known)
        prompt = (
            "Translate these labels from People's Bank of China statistical tables (money "
            "supply and Aggregate Financing to the Real Economy) into concise English, using "
            "the terminology of the PBoC's own English releases (e.g. 货币和准货币（M2） -> "
            "'Money and quasi-money (M2)', 社会融资规模存量 -> 'AFRE (stock)', 存量 -> 'Stock', "
            "增速（%） -> 'Growth Rate (%)', 单位：亿元人民币 -> '100 million yuan'). If a label "
            "already contains an English version, return just that English part, tidied. "
            "Units: return only the unit, without a 'Unit:' prefix. Drop 'of which' / 其中 "
            "prefixes and footnote markers. The same series is labelled slightly differently "
            "across years, so translate equivalent labels identically, and when a label means "
            "the same as one of the existing English labels below, return that existing label "
            "exactly. Translate every label; copy each source exactly as given.\n\n"
            f"Existing English labels:\n{vocab or '(none yet)'}\n\n"
            f"Labels:\n{listing}"
        )
        wanted = set(labels)
        out = self._parse(prompt, TranslationBatch).translations
        return {t.source: t.english.strip() for t in out
                if t.source in wanted and t.english.strip() and not CJK.search(t.english)}


class HeuristicSelector:
    """Keyword/regex stand-in used by --dry-run (no API key needed)."""

    def pick_pages(self, links: list[Link]) -> list[Link]:
        return [l for l in links if re.search(r"货币统计概览|社会融资规模", l.text)]

    def pick_attachments(self, links: list[Link]) -> list[tuple[Link, str]]:
        out = []
        for l in links:
            if re.search(r"货币供应量|Money Supply", l.context):
                out.append((l, "money_supply"))
            elif re.search(r"地区", l.context):
                continue  # regional breakdowns
            elif re.search(r"社会融资规模存量", l.context):
                out.append((l, "afre_stock"))
            elif re.search(r"社会融资规模(增量)?统计表", l.context):  # 2012 omits 增量
                out.append((l, "afre_flow"))
        return out

    def detect_layout(self, df: pd.DataFrame, title_hint: str) -> SheetLayout:
        is_period = df.map(lambda v: norm_period(v) is not None and len(norm_period(v)) == 7)
        row_counts, col_counts = is_period.sum(axis=1), is_period.sum(axis=0)
        footer = next((r for r in range(len(df))
                       if re.match(r"^\s*(注|Note)", str(df.iat[r, 0]))), None)
        unit = next((clean_label(v) for v in df.to_numpy().ravel()
                     if re.match(r"^\s*(单位|Unit)", str(v))), None)
        # Sheets can hold a supplementary table below the main one (in either orientation),
        # so pick whichever orientation's table starts first, not the one with most periods.
        header_rows = [r for r in range(len(df)) if row_counts.iat[r] >= 3]
        pc_best = int(col_counts.idxmax())
        col_start = (int(is_period.iloc[:, pc_best].to_numpy().argmax())
                     if col_counts.max() >= 3 else None)
        use_columns = bool(header_rows) and (col_start is None or header_rows[0] <= col_start)
        if not header_rows and col_start is None:
            use_columns = row_counts.max() >= col_counts.max()
        if use_columns:
            hr = header_rows[0] if header_rows else int(row_counts.idxmax())
            first_period_col = int(is_period.iloc[hr].to_numpy().argmax())
            label_cols = list(range(first_period_col))
            value_cols = df.columns[first_period_col:]
            # First row whose label columns hold text and value columns hold numbers.
            numeric = df[value_cols].apply(to_number).notna().any(axis=1)
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
        # Prefer the Chinese label row: English sub-rows can be split over two rows
        # (e.g. 'Foreign currency-denominated loans' / '(RMB equivalent)').
        candidates = [r for r in range(max(0, first - 6), first)
                      if df.iloc[r].drop(df.columns[pc]).notna().sum() >= (df.shape[1] - 1) / 2]
        hr = max(candidates, key=lambda r: (
            sum(bool(CJK.search(str(v))) for v in df.iloc[r].dropna()), r),
            default=first - 1)
        return SheetLayout(contains_data=True, orientation="periods_in_rows", header_row=hr,
                           data_start_row=first, footer_start_row=footer, measure_row=None,
                           label_columns=[], period_column=pc, unit_note=unit,
                           table_title=title_hint)

    def translate(self, labels: list[str], known: list[str]) -> dict[str, str]:
        return {}  # dry run: only labels already in data/translations.json get translated


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
    out["value"] = to_number(out["value"]).to_numpy()
    return out.dropna(subset=["value"]).reset_index(drop=True)


def process_file(selector, content: bytes, link: Link, dataset: str,
                 source_page: str) -> pd.DataFrame:
    frames = []
    if re.search(r"\.html?$", link.url, re.I):
        sheets = html_tables(content)
    else:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None)
    for name, df in sheets.items():
        if df.dropna(how="all").empty:
            continue
        layout = selector.detect_layout(df, link.context)
        if not layout.contains_data:
            log.info("  sheet %s: no data table, skipped", name)
            continue
        long = reshape(df, layout)
        if not isinstance(selector, HeuristicSelector):
            # Cross-check: a model layout that misses most of the table (e.g. a header row
            # that labels only one column) loses data silently, so compare with the heuristic.
            try:
                fallback = HeuristicSelector().detect_layout(df, link.context)
                fallback_long = reshape(df, fallback)
            except Exception as exc:
                log.warning("  sheet %s: heuristic cross-check failed: %s", name, exc)
                fallback_long = pd.DataFrame()
            if len(long) < CROSS_CHECK_RATIO * len(fallback_long):
                log.warning("  sheet %s: model layout gave %d values vs heuristic %d; using "
                            "heuristic. Model layout: %s", name, len(long), len(fallback_long),
                            layout.model_dump_json())
                layout, long = fallback, fallback_long
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
    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if out.empty:
        # A selected file must yield data; failing keeps it out of processed_urls.json
        # so the next run retries it instead of silently skipping it forever.
        raise RuntimeError(f"No values extracted from {link.url}")
    return out


# --------------------------------------------------------------------------- main

def run(dry_run: bool, years: int, max_files: int | None, no_llm: bool = False) -> int:
    from playwright.sync_api import sync_playwright

    if dry_run:
        selector = HeuristicSelector()
        log.info("DRY RUN: keyword/heuristic selection, nothing written to %s", DATA_DIR)
    elif no_llm:
        selector = HeuristicSelector()
        log.info("NO-LLM live run: keyword/heuristic selection; labels missing from %s stay "
                 "untranslated until the next Claude run", TRANSLATIONS_PATH.name)
    else:
        if not os.environ.get("ANTHROPIC_API_KEY"):
            log.error("ANTHROPIC_API_KEY is not set (use --dry-run to test without it)")
            return 2
        selector = ClaudeSelector()

    state = load_json(STATE_PATH, {})
    backfilling = not dry_run and not state.get("backfill_complete")
    if backfilling:
        years = max(years, BACKFILL_YEARS)
        state["backfill_attempts"] = state.get("backfill_attempts", 0) + 1
        log.info("One-time backfill: crawling the latest %d year sections (attempt %d/%d)",
                 years, state["backfill_attempts"], BACKFILL_MAX_ATTEMPTS)

    processed = load_processed()
    retrieved_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    frames: list[pd.DataFrame] = []
    done_urls: list[str] = []
    failures = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(locale="zh-CN")
        context.route("**/*", block_subresources)
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
                if atts and not isinstance(selector, HeuristicSelector):
                    # Keyword matches the model missed (e.g. 2012's flow table, titled without
                    # 增量) are added; the model's dataset tag wins where both picked a file.
                    have = {l.url for l, _ in chosen}
                    extra = [(l, ds) for l, ds in HeuristicSelector().pick_attachments(atts)
                             if l.url not in have]
                    for l, ds in extra:
                        log.info("  keyword match added: %s -> %s", l.context[:50], ds)
                    chosen = chosen + extra
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
    translations = load_json(TRANSLATIONS_PATH, {})
    if dry_run:
        new = translate_columns(new, selector, translations)
        log.info("DRY RUN result: %d records from %d file(s), %d failure(s)",
                 len(new), len(done_urls), failures)
        if not new.empty:
            with pd.option_context("display.width", 200, "display.max_columns", 10):
                print(new[["dataset", "item", "measure", "period", "value", "unit"]]
                      .groupby("dataset").head(4).to_string(index=False))
                print(new.groupby("dataset")["period"].agg(["min", "max", "count"]))
        return 1 if failures else 0

    merged = merge_with_master(new[MASTER_COLUMNS])
    # Translating the merged frame also converts rows already in the master, and must
    # happen before de-duplication so old Chinese keys collapse onto the English ones.
    n_known = len(translations)
    translated = translate_columns(merged, selector, translations)
    # Rewrite only on a real content change: an unchanged run must leave the file untouched
    # so the workflow has nothing to commit (a re-saved .xlsx differs byte-wise).
    final = prepare_master(translated)
    old = (prepare_master(pd.read_excel(MASTER_PATH, dtype={"period": str}))
           if MASTER_PATH.exists() else None)
    if old is not None and same_content(final, old):
        log.info("Master unchanged (%d rows); not rewritten", len(final))
    elif not final.empty:
        total = write_master(final)
        log.info("Upserted %d new records; master now has %d rows", len(new), total)
    if not final.empty:
        written = write_web_exports(final)
        log.info("Web exports: %s", ", ".join(written) if written else "unchanged")
    if len(translations) != n_known:
        save_json(TRANSLATIONS_PATH, translations)
    if done_urls:
        save_processed(processed | set(done_urls))
    if backfilling:
        if failures == 0 or state["backfill_attempts"] >= BACKFILL_MAX_ATTEMPTS:
            state["backfill_complete"] = True
            state["backfill_years"] = years
            log.info("Backfill complete (%d failure(s)); later runs crawl only the latest "
                     "year unless --years says otherwise", failures)
        save_json(STATE_PATH, state)
    log.info("Done: %d file(s) processed, %d failure(s)", len(done_urls), failures)
    return 1 if failures else 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--dry-run", action="store_true",
                    help="Test crawl + parse without the Claude API and without writing data/")
    ap.add_argument("--years", type=int, default=2,
                    help="How many of the most recent year sections to crawl (default 2: December data "
                    "is published in January under the previous year)")
    ap.add_argument("--max-files", type=int, default=None,
                    help="Cap on new files processed this run")
    ap.add_argument("--no-llm", action="store_true",
                    help="Live run (writes data/) with heuristic selection instead of Claude; "
                    "no API key needed. For fetching files the runners can't reach.")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    sys.exit(run(args.dry_run, args.years, args.max_files, args.no_llm))


if __name__ == "__main__":
    main()
