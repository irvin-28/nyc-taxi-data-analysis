# NYC Taxi Trips: Data Cleaning and Excel Dashboard

Cleaned a public NYC taxi trip sample (6,433 raw trips, March 2019) with Python and built an Excel dashboard where every number is a live formula.

## What I did
- **Cleaned the data in Python (pandas):** parsed dates, removed duplicates, filled missing zone, borough and payment values, removed invalid trips (zero distance, zero passengers, impossible durations and speeds).
- **Added new columns:** trip minutes, tip %, pickup hour and weekday.
- **Built an Excel dashboard:** KPIs, summary tables (SUMIFS, COUNTIFS, AVERAGEIFS) and four charts.
- **Documented every cleaning step** in the Cleaning Log sheet.

## Key findings
- 6,255 clean trips, $115.5K total revenue, average fare $13.04.
- Manhattan makes up about 82% of pickups, but Queens trips have the highest average fare (about $25) because they are longer.
- Credit card riders tip about 24% on average. Cash tips are not recorded in the data, so tip % is credit card only.
- About 28% of trips are paid in cash.
- Busiest pickup hours are 5 to 8 PM, and Friday and Saturday are the busiest days.

## Files
| File | What it is |
|---|---|
| `taxi_analysis.py` | Cleaning, analysis and dashboard build script |
| `taxis_raw.csv` | Original dataset |
| `taxis_clean.csv` | Cleaned dataset |
| `NYC_Taxi_Dashboard.xlsx` | Excel dashboard (Dashboard, Data, Cleaning Log) |

## Run it
```
pip install pandas openpyxl
python taxi_analysis.py
```

Data source: seaborn-data `taxis` dataset (public NYC TLC trip sample).
