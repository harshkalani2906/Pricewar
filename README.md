# PRICE WAR - Nat Habit price parity tracker

Live Nat Habit selling prices on 10 channels, compared with the aligned price (Event or BAU).

- **Marketplaces** (one all-India price): Amazon, Flipkart, Myntra, Nykaa, FirstCry
- **Quick commerce** (city by city): Blinkit, Zepto, Instamart, BigBasket, Flipkart Minutes

## Folders

| Folder / file | What it is |
|---|---|
| `1_Input/` | The Excel files you edit: Links, Product Master, Pincodes, Aligned Price, SKU Priority |
| `_engine/` | The scripts (no need to open) |
| `3_Dashboard/price-pulse.html` | The dashboard - download and open in Chrome |
| `*.bat` | Buttons to double-click on the Windows laptop |

## Buttons

- **FIRST TIME SETUP.bat** - install once
- **LOG IN TO APPS (once).bat** - log in to Instamart, BigBasket, Flipkart Minutes
- **RUN PRICE TRACKER.bat** - check all prices now
- **RUN AMAZON NOW.bat** - check Amazon only
- **SCHEDULE DAILY RUN.bat** - automatic: 07:30 daily (80% revenue SKUs) + every 2 hours 09:30-21:30 (Top 50)
- **STOP DAILY RUN.bat** - stop automatic runs
- **OPEN DASHBOARD.bat** - rebuild and open the dashboard
- **PUSH TO GITHUB.bat** - upload to this repository (after the first time it uploads by itself after every run)

## Needs these folders next to it (same Desktop)

Price Check, Amazon, Flipkart, Myntra, Nykaa, FirstCry, SLA TRACKER\SLA Tracker v2 (not in this repository).

Not uploaded on purpose: app logins (`_engine/state`), results and logs (`2_Output`).
