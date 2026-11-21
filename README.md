# pbocscraper

Twice-weekly scraper for People's Bank of China statistics: **money supply (M0, M1, M2)**,
**Aggregate Financing to the Real Economy (AFRE, flow and stock)** and the **Medium-term Lending
Facility (MLF)**, translated to English.
Runs on GitHub Actions every Monday and Thursday at 08:00 UTC and commits any new data.

## Data files

| File | Contents |
|---|---|
| [`data/pboc_data.json`](data/pboc_data.json) | One entry per series with `[period, value]` pairs, ready for charting |
| [`data/pboc_data.csv`](data/pboc_data.csv) | Long format: `dataset, item, measure, period, value, unit` |
| [`data/pboc_master_data.xlsx`](data/pboc_master_data.xlsx) | Full master incl. source file URLs and retrieval times |

Coverage: money supply 2008-01 onwards, AFRE flow 2012-01 onwards, AFRE stock 2015 (quarterly)
then monthly from 2016-01. MLF: month-end outstanding balance from 2014-09 (published by the PBoC
until 2025-05; later months are derived as last published balance + cumulative net injection and
flagged with `derived_from` in the JSON), and monthly injection / withdrawal / net injection from
2025-05. Periods are `YYYY-MM` (`YYYY-Qn` for the 2015 quarterly stock).
Units: money supply, AFRE flow and MLF in 100 million yuan; AFRE stock in trillion yuan; growth rates in %.

### Using the JSON from a website

```js
const URL = "https://raw.githubusercontent.com/Liquid-dash/pbocscraper/main/data/pboc_data.json";
const { updated_at_utc, series } = await fetch(URL).then(r => r.json());
const m2 = series.find(s => s.id === "money-supply/money-and-quasi-money-m2/value");
// m2.data -> [["2008-01", 417846.17], ..., ["2026-10", 3590550.61]], m2.unit -> "100 million yuan"
```

raw.githubusercontent.com allows cross-origin requests and caches for about 5 minutes. Series
`id`s are stable (`dataset/item/measure`, slugified); the JSON's `series` list shows them all.

Source: [PBoC Statistics & Analysis Department](http://www.pbc.gov.cn/diaochatongjisi/116219/116319/index.html).
