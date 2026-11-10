import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

# ---------- 1. LOAD ----------
df = pd.read_csv("taxis_raw.csv")
log = [("Raw rows loaded", len(df), "")]

# ---------- 2. CLEAN ----------
df["pickup"] = pd.to_datetime(df["pickup"])
df["dropoff"] = pd.to_datetime(df["dropoff"])

n = len(df)
df = df.drop_duplicates()
log.append(("Duplicate rows removed", n - len(df), "Exact duplicate trips"))

for col in ["payment", "pickup_zone", "dropoff_zone", "pickup_borough", "dropoff_borough"]:
    missing = int(df[col].isna().sum())
    df[col] = df[col].fillna("Unknown")
    log.append((f"Missing '{col}' filled", missing, "Replaced with 'Unknown'"))

df["payment"] = df["payment"].str.strip().str.title()
df["color"] = df["color"].str.strip().str.title()

df["trip_minutes"] = (df["dropoff"] - df["pickup"]).dt.total_seconds() / 60

n = len(df)
df = df[(df["distance"] > 0) & (df["passengers"] > 0)]
log.append(("Rows removed: zero distance or zero passengers", n - len(df), "Not valid completed trips"))

n = len(df)
df = df[(df["trip_minutes"] >= 1) & (df["trip_minutes"] <= 180)]
log.append(("Rows removed: trip under 1 min or over 180 min", n - len(df), "Likely meter errors"))

df["speed_mph"] = df["distance"] / (df["trip_minutes"] / 60)
n = len(df)
df = df[df["speed_mph"] <= 80]
log.append(("Rows removed: speed above 80 mph", n - len(df), "Physically unrealistic"))

# ---------- 3. FEATURES ----------
df["tip_pct"] = (df["tip"] / df["fare"]).round(4)
df["hour"] = df["pickup"].dt.hour
df["weekday"] = df["pickup"].dt.day_name()
df["trip_minutes"] = df["trip_minutes"].round(1)
df = df.drop(columns=["speed_mph"])
df = df.sort_values("pickup").reset_index(drop=True)
log.append(("Clean rows kept", len(df), "Used in the dashboard"))

df.to_csv("taxis_clean.csv", index=False)

# ---------- 4. QUICK INSIGHTS (printed for the README) ----------
cc = df[df["payment"] == "Credit Card"]
print("Rows:", len(df))
print("Avg fare:", round(df["fare"].mean(), 2))
print("Avg tip % (credit card only):", round(cc["tip_pct"].mean() * 100, 1))
print("Cash share of trips:", round((df["payment"] == "Cash").mean() * 100, 1))
print(df.groupby("pickup_borough")["fare"].agg(["count", "mean"]).round(2))
print(df.groupby("hour")["fare"].count().sort_values(ascending=False).head(3))
print(df.groupby("weekday")["fare"].count().sort_values(ascending=False).head(3))
print(df.groupby("payment")["tip_pct"].mean().round(3))

# ---------- 5. EXCEL DASHBOARD ----------
ARIAL = "Arial"
head_font = Font(name=ARIAL, bold=True, color="FFFFFF")
head_fill = PatternFill("solid", start_color="1F3864")
bold = Font(name=ARIAL, bold=True)
norm = Font(name=ARIAL)

wb = Workbook()
ws = wb.active
ws.title = "Dashboard"
wd = wb.create_sheet("Data")
wl = wb.create_sheet("Cleaning Log")

# Data sheet
cols = list(df.columns)
wd.append(cols)
for row in df.itertuples(index=False):
    wd.append([v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in row])
for c in wd[1]:
    c.font = head_font
    c.fill = head_fill
last = len(df) + 1
wd.freeze_panes = "A2"
col_idx = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}
for name in ["pickup", "dropoff"]:
    for r in range(2, last + 1):
        wd[f"{col_idx[name]}{r}"].number_format = "yyyy-mm-dd hh:mm"
for r in range(2, last + 1):
    wd[f"{col_idx['tip_pct']}{r}"].number_format = "0.0%"
for i, name in enumerate(cols):
    wd.column_dimensions[get_column_letter(i + 1)].width = 18 if name not in ("pickup_zone", "dropoff_zone") else 30


def rng(name):
    c = col_idx[name]
    return f"Data!${c}$2:${c}${last}"


# Dashboard title
ws["A1"] = "NYC Taxi Trips - Dashboard (March 2019 sample)"
ws["A1"].font = Font(name=ARIAL, bold=True, size=16, color="1F3864")
ws["A2"] = "Source: seaborn-data 'taxis' dataset (public NYC TLC sample). All figures are live formulas on the Data sheet."
ws["A2"].font = Font(name=ARIAL, italic=True, size=9)

# KPIs
kpis = [
    ("Total trips", f"=COUNT({rng('fare')})", "#,##0"),
    ("Total revenue ($)", f"=SUM({rng('total')})", "$#,##0"),
    ("Average fare ($)", f"=AVERAGE({rng('fare')})", "$#,##0.00"),
    ("Avg distance (miles)", f"=AVERAGE({rng('distance')})", "0.00"),
    ("Avg trip time (min)", f"=AVERAGE({rng('trip_minutes')})", "0.0"),
    ("Avg tip % (credit card)", f"=AVERAGEIFS({rng('tip_pct')},{rng('payment')},\"Credit Card\")", "0.0%"),
    ("Cash trips share", f"=COUNTIFS({rng('payment')},\"Cash\")/COUNT({rng('fare')})", "0.0%"),
]
for i, (label, formula, fmt) in enumerate(kpis):
    r = 4 + i
    ws[f"A{r}"] = label
    ws[f"A{r}"].font = bold
    ws[f"B{r}"] = formula
    ws[f"B{r}"].number_format = fmt
    ws[f"B{r}"].font = norm
ws["A11"] = "Note: cash tips are not recorded in the data, so tip % is calculated for credit card trips only."
ws["A11"].font = Font(name=ARIAL, italic=True, size=9)


def table(top, title, header, labels, formulas_fn, fmts):
    ws[f"A{top}"] = title
    ws[f"A{top}"].font = Font(name=ARIAL, bold=True, size=12, color="1F3864")
    for j, h in enumerate(header):
        c = ws.cell(row=top + 1, column=1 + j, value=h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")
    for i, lab in enumerate(labels):
        r = top + 2 + i
        ws.cell(row=r, column=1, value=lab).font = norm
        for j, f in enumerate(formulas_fn(r)):
            c = ws.cell(row=r, column=2 + j, value=f)
            c.number_format = fmts[j]
            c.font = norm
    return top + 2, top + 1 + len(labels)


boroughs = sorted(df["pickup_borough"].unique())
b0, b1 = table(
    13, "Trips by pickup borough", ["Borough", "Trips", "Avg fare ($)", "Revenue ($)"], boroughs,
    lambda r: [
        f"=COUNTIFS({rng('pickup_borough')},A{r})",
        f"=IFERROR(AVERAGEIFS({rng('fare')},{rng('pickup_borough')},A{r}),0)",
        f"=SUMIFS({rng('total')},{rng('pickup_borough')},A{r})",
    ],
    ["#,##0", "$#,##0.00", "$#,##0"],
)

pay = sorted(df["payment"].unique())
p_top = b1 + 3
p0, p1 = table(
    p_top, "Payment type", ["Payment", "Trips", "Avg fare ($)", "Avg tip %"], pay,
    lambda r: [
        f"=COUNTIFS({rng('payment')},A{r})",
        f"=IFERROR(AVERAGEIFS({rng('fare')},{rng('payment')},A{r}),0)",
        f"=IFERROR(AVERAGEIFS({rng('tip_pct')},{rng('payment')},A{r}),0)",
    ],
    ["#,##0", "$#,##0.00", "0.0%"],
)

days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
d_top = p1 + 3
d0, d1 = table(
    d_top, "Trips by weekday", ["Weekday", "Trips", "Avg fare ($)", "Revenue ($)"], days,
    lambda r: [
        f"=COUNTIFS({rng('weekday')},A{r})",
        f"=IFERROR(AVERAGEIFS({rng('fare')},{rng('weekday')},A{r}),0)",
        f"=SUMIFS({rng('total')},{rng('weekday')},A{r})",
    ],
    ["#,##0", "$#,##0.00", "$#,##0"],
)

h_top = d1 + 3
h0, h1 = table(
    h_top, "Trips by pickup hour", ["Hour", "Trips", "Avg fare ($)", "Avg distance (mi)"], list(range(24)),
    lambda r: [
        f"=COUNTIFS({rng('hour')},A{r})",
        f"=IFERROR(AVERAGEIFS({rng('fare')},{rng('hour')},A{r}),0)",
        f"=IFERROR(AVERAGEIFS({rng('distance')},{rng('hour')},A{r}),0)",
    ],
    ["#,##0", "$#,##0.00", "0.00"],
)

for col, w in zip("ABCD", [30, 14, 16, 18]):
    ws.column_dimensions[col].width = w

# Charts
ch1 = BarChart()
ch1.title = "Trips by pickup borough"
ch1.add_data(Reference(ws, min_col=2, min_row=b0 - 1, max_row=b1), titles_from_data=True)
ch1.set_categories(Reference(ws, min_col=1, min_row=b0, max_row=b1))
ch1.legend = None
ch1.height, ch1.width = 7.5, 14
ws.add_chart(ch1, "F4")

ch2 = BarChart()
ch2.title = "Average tip % by payment type"
ch2.add_data(Reference(ws, min_col=4, min_row=p0 - 1, max_row=p1), titles_from_data=True)
ch2.set_categories(Reference(ws, min_col=1, min_row=p0, max_row=p1))
ch2.legend = None
ch2.height, ch2.width = 7.5, 14
ws.add_chart(ch2, "F20")

ch3 = BarChart()
ch3.title = "Trips by weekday"
ch3.add_data(Reference(ws, min_col=2, min_row=d0 - 1, max_row=d1), titles_from_data=True)
ch3.set_categories(Reference(ws, min_col=1, min_row=d0, max_row=d1))
ch3.legend = None
ch3.height, ch3.width = 7.5, 14
ws.add_chart(ch3, "F36")

ch4 = LineChart()
ch4.title = "Trips by pickup hour"
ch4.add_data(Reference(ws, min_col=2, min_row=h0 - 1, max_row=h1), titles_from_data=True)
ch4.set_categories(Reference(ws, min_col=1, min_row=h0, max_row=h1))
ch4.legend = None
ch4.height, ch4.width = 7.5, 14
ws.add_chart(ch4, "F52")

# Cleaning log sheet
wl.append(["Step", "Rows", "Note"])
for c in wl[1]:
    c.font = head_font
    c.fill = head_fill
for item in log:
    wl.append(list(item))
wl.append([])
wl.append(["Row counts above are written by taxi_analysis.py when it runs."])
wl.column_dimensions["A"].width = 52
wl.column_dimensions["B"].width = 10
wl.column_dimensions["C"].width = 40
for row in wl.iter_rows(min_row=2):
    for c in row:
        c.font = norm

wb.save("NYC_Taxi_Dashboard.xlsx")
print("Saved NYC_Taxi_Dashboard.xlsx")
