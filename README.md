# US High-Frequency Economic Data

A single self-contained HTML dashboard tracking US activity, labor, credit,
and inflation off weekly/daily source data — plus a Python script that
refreshes it from the original public sources on demand.

**[Open the live dashboard](./us_high_frequency_dashboard.html)** — download
and open the file in any browser, no server or build step required.

## What's in it

**11 indicator panels** — Fed Weekly Economic Index, Atlanta Fed GDPNow, US
Rail Traffic, Initial & Continued Jobless Claims, Withheld Employment Taxes,
Bank Loan Growth, Chicago Fed NFCI, St. Louis Fed Financial Stress, 30-Yr
Mortgage Rate, Cleveland Fed CPI Nowcast.

**4 seasonal comparison panels** — TSA Airport Throughput, Weekly Business
Applications, US Rail Traffic, and Initial Jobless Claims, each plotting the
current year against the 2021–2025 seasonal average with a min–max range
band.

| Source | Series |
|---|---|
| FRED (St. Louis Fed) | WEI, ICSA/ICNSA, CCSA, TOTLL, NFCI, MORTGAGE30US, STLFSI4, rail carloads/intermodal (SA + NSA) |
| Atlanta Fed | GDPNow daily nowcast |
| Cleveland Fed | CPI nowcast |
| US Treasury (Fiscal Data API) | Withheld employment tax deposits (DTS Table IV + II, spliced) |
| US Census Bureau | Weekly Business Formation Statistics |
| TSA | Daily checkpoint throughput |

All sources are public with no paid subscription required. A few series from
the original reference layout (Redbook, Citi Economic Surprise Index, IBES
forward EPS, Truflation) are proprietary with no public API and are
deliberately **not** reproduced — see the notes in the dashboard footer.

## Quick start

```bash
git clone <this-repo-url>
cd <repo>
pip install -r requirements.txt

export FRED_API_KEY=your_key_here   # see below
python rebuild.py
```

This fetches fresh data from every source above and rewrites
`us_high_frequency_dashboard.html` in place — open it in a browser to view.

### Get a FRED API key

The script needs one free key from the St. Louis Fed:

1. Create an account at <https://fred.stlouisfed.org>
2. Request a key at <https://fred.stlouisfed.org/docs/api/api_key.html>
   (instant, no approval wait)
3. Set it as an environment variable before running the script:
   ```bash
   export FRED_API_KEY=your_key_here      # macOS/Linux
   setx FRED_API_KEY your_key_here        # Windows
   ```

No other source in the pipeline requires a key.

## Automated refresh (optional)

`.github/workflows/refresh.yml` runs `rebuild.py` on a schedule (06:00 UTC,
Tue–Sat) and commits the updated HTML back to the repo — so if you enable
GitHub Pages on this repo, the published dashboard stays current on its own.

To enable it:

1. **Settings → Secrets and variables → Actions → New repository secret**
   Name: `FRED_API_KEY`, value: your key from above.
2. **Settings → Actions → General → Workflow permissions** → set to
   *Read and write permissions* (so the workflow can commit the refreshed file).
3. Optionally trigger it immediately from the **Actions** tab →
   *Refresh dashboard* → *Run workflow*, instead of waiting for the schedule.

### Publishing with GitHub Pages

**Settings → Pages → Source** → deploy from the branch this repo lives on,
root folder. The dashboard will be served at
`https://<username>.github.io/<repo>/us_high_frequency_dashboard.html`.

## Repo layout

```
us_high_frequency_dashboard.html   the dashboard — self-contained, no external data calls
rebuild.py                         fetches all sources, recomputes, rewrites the HTML above
requirements.txt                   Python dependencies for rebuild.py
.github/workflows/refresh.yml      optional scheduled auto-refresh
```

`data.json` is not stored in the repo — it's an intermediate the script
regenerates each run and injects directly into the HTML.

## Notes on methodology

- **Seasonal panels** compare the current year against the mean and
  min–max range of the five most recently completed calendar years
  (2021–2025). That baseline includes 2021, when several series (initial
  claims especially) were still running at pandemic-distorted levels —
  this widens the range and lifts the average in the affected months. Kept
  for consistency across panels rather than special-cased.
- **GDPNow** resets each quarter and is naturally most volatile in the
  first few weeks of a new quarter, producing a sawtooth pattern.
- **Rail traffic** appears twice with different series: the indicator-grid
  panel uses the seasonally-adjusted series (a clean level/trend read);
  the seasonal-comparison panel uses the non-adjusted series, since SA data
  has the seasonal signal removed by construction and can't be compared
  against a seasonal norm.
- **CPI Nowcast** chains the Cleveland Fed's monthly headline-CPI model
  finals into a year-over-year rate; the raw single-month annualized figure
  is too noisy to read on its own.

## License / usage

Provided for research purposes. Not investment advice. All underlying data
belongs to its respective source agency — see the dashboard footer for full
attribution.
