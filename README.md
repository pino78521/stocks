# Pino Market Dashboard (Auto-Updating)

## What this repo does
- Stores your Excel dashboard
- Runs a GitHub Action (automation) on a schedule (and on-demand)
- Pulls daily prices and updates:
  - 50-day SMA, 200-day SMA
  - 50/200 spread (%)
  - Price vs 50-day (%)
  - RSI (14)
  - Sector heat (using sector ETFs)
  - Action label (your rule-based signal)

## Sheets updated
- **Dashboard**: the 11 sector leaders + sector heat (sector ETF trend)
- **Watchlist**: add any tickers you want, optional sector ETF, and it will compute the same metrics

## Data sources
- **Stooq** (free, no key) for daily OHLC data (used for all calculations)
- **yfinance** (optional, best-effort) for dividend yield

## How to use
1. Upload everything in this folder to a GitHub repo.
2. Go to the **Actions** tab and enable Actions if asked.
3. Run the workflow once: **Update Pino Market Dashboard** → **Run workflow**
4. Download the updated Excel file: `Pino_Market_Dashboard_Auto.xlsx`

## Change the schedule
Edit `.github/workflows/update-dashboard.yml`.
GitHub uses **UTC** for schedules.
