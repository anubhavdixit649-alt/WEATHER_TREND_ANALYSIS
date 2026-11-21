"""
Weather Climate Search - interactive program (Python + Pandas + NumPy)

Lets you search cities, check temperature / rainfall / wind, see the climate
of a city and compare cities.

The weather data is built into this file (it is a copy of the 'Weather Data'
sheet of Weather_Analysis_Results.xlsx), so no Excel file, upload or internet
is needed. Just run it:
    python weather_climate_search.py

Everything shown is calculated from that built-in dataset. The only extra column is
"Condition", a simple label this program creates from rainfall, temperature
and wind (the rules are at the top of the file, you can change them).
"""

import io
import numpy as np
import pandas as pd

# ----------------------------------------------------------------------
# SETTINGS
# ----------------------------------------------------------------------
# Rules for the "Condition" label (change them if you like)
HEAVY_RAIN_MM = 25      # rainfall at or above this  -> Heavy Rain
MODERATE_RAIN_MM = 5    # rainfall at or above this  -> Moderate Rain
HOT_DAY_MAX_C = 35      # max temperature at or above this -> Hot
WINDY_KMH = 15          # wind speed at or above this -> Windy
WARM_AVG_C = 28         # average temperature at or above this -> Warm
COOL_AVG_C = 20         # average temperature below this -> Cool

NUMBER_COLUMNS = ["Max_Temperature_C", "Min_Temperature_C",
                  "Average_Temperature_C", "Rainfall_mm", "Wind_Speed_kmh"]

TABLE_COLUMNS = ["Date", "City", "Max_Temperature_C", "Min_Temperature_C",
                 "Average_Temperature_C", "Rainfall_mm", "Wind_Speed_kmh",
                 "Condition"]

HEADERS = {
    "Date": "Date", "City": "City",
    "Max_Temperature_C": "Max Temp (C)",
    "Min_Temperature_C": "Min Temp (C)",
    "Average_Temperature_C": "Avg Temp (C)",
    "Rainfall_mm": "Rain (mm)",
    "Wind_Speed_kmh": "Wind (km/h)",
    "Condition": "Condition",
}

CITY_TABLE_HEADERS = {
    "avg_temp": "Avg Temp (C)", "highest": "Highest Temp (C)",
    "lowest": "Lowest Temp (C)", "total_rain": "Total Rain (mm)",
    "rainy_days": "Rainy Days", "avg_wind": "Avg Wind (km/h)",
    "max_wind": "Max Wind (km/h)",
}


# ----------------------------------------------------------------------
# EMBEDDED DATASET
# ----------------------------------------------------------------------
# Copy of the "Weather Data" sheet of Weather_Analysis_Results.xlsx
# (5 cities x 20 days, 2026-10-27 to 2026-11-15). Columns: date, city, max / min /
# average temperature (C), rainfall (mm), wind speed (km/h).
DATA_CSV = """
Date,City,Max_Temperature_C,Min_Temperature_C,Average_Temperature_C,Rainfall_mm,Wind_Speed_kmh
2026-10-27,Bengaluru,26.9,17.4,22.0,0.0,11.4
2026-10-28,Bengaluru,28.0,17.6,22.6,0.0,15.9
2026-10-29,Bengaluru,28.8,18.5,23.3,0.0,11.4
2026-10-30,Bengaluru,30.3,18.4,24.5,0.0,10.7
2026-10-31,Bengaluru,31.0,20.5,25.4,0.0,12.7
2026-11-01,Bengaluru,29.5,20.6,24.6,0.9,12.3
2026-11-02,Bengaluru,26.5,20.6,22.9,15.3,11.8
2026-11-03,Bengaluru,28.6,19.5,23.6,0.8,6.6
2026-11-04,Bengaluru,28.1,19.8,23.0,3.7,8.1
2026-11-05,Bengaluru,27.9,20.4,23.0,6.8,11.9
2026-11-06,Bengaluru,27.3,20.0,22.7,2.9,12.9
2026-11-07,Bengaluru,28.1,19.9,23.0,2.7,14.3
2026-11-08,Bengaluru,28.3,20.1,23.5,0.8,14.3
2026-11-09,Bengaluru,28.7,19.6,23.6,0.0,14.5
2026-11-10,Bengaluru,28.8,19.1,23.4,0.0,11.5
2026-11-11,Bengaluru,29.0,18.6,23.5,0.0,10.7
2026-11-12,Bengaluru,28.2,18.2,22.8,0.0,10.7
2026-11-13,Bengaluru,28.7,17.2,22.8,0.0,13.0
2026-11-14,Bengaluru,28.6,17.8,23.4,0.0,10.6
2026-11-15,Bengaluru,27.0,19.0,22.9,0.0,8.4
2026-10-27,Delhi,29.1,18.5,23.7,0.0,7.2
2026-10-28,Delhi,29.2,18.2,24.0,0.0,4.8
2026-10-29,Delhi,24.7,19.3,21.4,46.5,13.1
2026-10-30,Delhi,21.2,18.8,20.1,50.3,8.8
2026-10-31,Delhi,27.1,18.6,22.6,0.0,8.5
2026-11-01,Delhi,27.6,19.4,23.3,0.0,7.0
2026-11-02,Delhi,28.3,19.4,23.7,0.0,6.1
2026-11-03,Delhi,28.5,18.2,22.9,0.0,10.5
2026-11-04,Delhi,28.4,17.8,23.0,0.0,7.7
2026-11-05,Delhi,28.6,18.3,23.2,0.0,6.8
2026-11-06,Delhi,28.6,18.6,23.2,0.0,9.5
2026-11-07,Delhi,28.6,18.6,23.2,0.0,7.1
2026-11-08,Delhi,28.5,18.7,23.4,0.0,7.6
2026-11-09,Delhi,28.2,18.3,23.0,0.0,7.7
2026-11-10,Delhi,28.3,18.9,23.1,0.0,6.0
2026-11-11,Delhi,27.9,17.5,22.6,0.0,6.7
2026-11-12,Delhi,28.4,18.3,23.4,0.0,5.4
2026-11-13,Delhi,29.5,18.6,23.8,0.0,6.0
2026-11-14,Delhi,28.9,19.5,23.7,0.0,6.0
2026-11-15,Delhi,27.2,19.0,22.9,0.0,8.3
2026-10-27,Jaipur,31.9,19.7,25.4,0.0,11.9
2026-10-28,Jaipur,30.9,19.0,24.6,0.0,10.8
2026-10-29,Jaipur,27.8,18.5,22.8,4.2,15.9
2026-10-30,Jaipur,24.3,18.2,20.6,2.6,18.6
2026-10-31,Jaipur,28.3,17.6,22.5,0.0,11.2
2026-11-01,Jaipur,28.2,19.0,23.4,0.0,5.9
2026-11-02,Jaipur,28.8,19.4,23.8,0.0,7.9
2026-11-03,Jaipur,30.0,18.9,23.8,0.0,7.5
2026-11-04,Jaipur,31.0,19.4,24.5,0.0,7.6
2026-11-05,Jaipur,31.5,18.5,24.6,0.0,8.5
2026-11-06,Jaipur,31.5,19.0,24.8,0.0,11.6
2026-11-07,Jaipur,31.1,18.8,24.6,0.0,10.2
2026-11-08,Jaipur,31.7,19.6,25.0,0.0,10.0
2026-11-09,Jaipur,31.5,18.7,24.6,0.0,12.9
2026-11-10,Jaipur,29.1,18.8,23.6,0.0,10.9
2026-11-11,Jaipur,29.4,17.7,23.6,0.0,6.5
2026-11-12,Jaipur,30.2,19.2,24.4,0.0,8.1
2026-11-13,Jaipur,31.0,20.1,24.9,0.0,8.5
2026-11-14,Jaipur,30.7,20.0,25.2,0.0,7.9
2026-11-15,Jaipur,27.7,19.5,23.4,0.0,14.7
2026-10-27,Mumbai,36.8,26.0,31.2,0.0,16.8
2026-10-28,Mumbai,35.5,25.4,30.1,0.0,16.1
2026-10-29,Mumbai,34.4,26.7,30.1,1.3,11.5
2026-10-30,Mumbai,31.8,23.4,28.0,4.1,10.3
2026-10-31,Mumbai,32.2,25.1,28.5,0.0,13.5
2026-11-01,Mumbai,33.2,26.9,30.0,0.0,10.0
2026-11-02,Mumbai,34.3,26.7,30.0,0.3,11.7
2026-11-03,Mumbai,34.9,26.1,30.4,1.3,12.0
2026-11-04,Mumbai,35.9,26.3,30.8,0.3,12.9
2026-11-05,Mumbai,36.2,26.1,30.5,0.0,12.2
2026-11-06,Mumbai,36.6,26.3,31.2,0.0,12.7
2026-11-07,Mumbai,36.4,26.1,30.9,0.0,13.0
2026-11-08,Mumbai,36.6,25.1,30.6,0.0,13.5
2026-11-09,Mumbai,36.0,25.4,30.6,0.0,15.6
2026-11-10,Mumbai,35.9,24.4,29.9,0.0,12.1
2026-11-11,Mumbai,35.6,24.1,29.9,0.0,11.7
2026-11-12,Mumbai,35.2,23.7,29.3,0.0,12.4
2026-11-13,Mumbai,34.9,23.3,29.1,0.0,11.3
2026-11-14,Mumbai,34.6,23.6,28.6,0.0,13.7
2026-11-15,Mumbai,33.7,22.8,28.1,0.0,12.5
2026-10-27,Prayagraj,30.7,21.1,25.8,0.0,8.5
2026-10-28,Prayagraj,30.2,22.7,26.3,0.0,9.9
2026-10-29,Prayagraj,25.5,21.0,23.3,4.2,18.2
2026-10-30,Prayagraj,23.2,20.7,21.8,9.0,18.8
2026-10-31,Prayagraj,25.4,21.4,23.2,0.4,8.2
2026-11-01,Prayagraj,27.7,22.6,24.5,0.8,5.2
2026-11-02,Prayagraj,28.7,21.2,24.7,0.0,3.2
2026-11-03,Prayagraj,30.1,21.7,25.3,0.0,7.1
2026-11-04,Prayagraj,30.2,20.3,24.6,0.0,8.1
2026-11-05,Prayagraj,29.9,19.0,24.0,0.0,6.4
2026-11-06,Prayagraj,29.7,18.6,23.9,0.0,7.6
2026-11-07,Prayagraj,30.1,19.8,24.3,0.0,10.1
2026-11-08,Prayagraj,30.0,19.3,23.9,0.0,6.1
2026-11-09,Prayagraj,28.8,19.2,23.7,0.0,4.8
2026-11-10,Prayagraj,28.6,19.0,23.7,0.0,4.8
2026-11-11,Prayagraj,28.7,19.6,23.8,0.0,5.5
2026-11-12,Prayagraj,28.9,18.7,23.5,0.0,5.5
2026-11-13,Prayagraj,29.5,19.1,23.8,0.0,7.9
2026-11-14,Prayagraj,29.7,19.1,23.7,0.0,12.0
2026-11-15,Prayagraj,29.9,18.5,23.9,0.0,6.2
"""


# ----------------------------------------------------------------------
# LOADING AND CLEANING THE DATA
# ----------------------------------------------------------------------
def load_data():
    """Read the built-in dataset, clean it and add the Condition column."""
    df = pd.read_csv(io.StringIO(DATA_CSV.strip()))
    start_rows = len(df)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["City"] = df["City"].astype(str).str.strip().str.title()
    for column in NUMBER_COLUMNS:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df.dropna(subset=["Date"] + NUMBER_COLUMNS)
    df = df.drop_duplicates(subset=["Date", "City"])
    df = df.sort_values(["City", "Date"]).reset_index(drop=True)

    df["Month"] = df["Date"].dt.month_name()
    df["Month_No"] = df["Date"].dt.month

    # Simple daily label made from the numbers
    rain = df["Rainfall_mm"].values
    conditions = [
        rain >= HEAVY_RAIN_MM,
        rain >= MODERATE_RAIN_MM,
        rain > 0,
        df["Max_Temperature_C"].values >= HOT_DAY_MAX_C,
        df["Wind_Speed_kmh"].values >= WINDY_KMH,
        df["Average_Temperature_C"].values >= WARM_AVG_C,
        df["Average_Temperature_C"].values >= COOL_AVG_C,
    ]
    labels = ["Heavy Rain", "Moderate Rain", "Light Rain",
              "Hot", "Windy", "Warm", "Pleasant"]
    df["Condition"] = np.select(conditions, labels, default="Cool")

    removed = start_rows - len(df)
    print(f"Loaded {len(df)} records from the built-in dataset"
          f" ({removed} bad or duplicate rows removed).")
    return df


# ----------------------------------------------------------------------
# INPUT HELPERS
# ----------------------------------------------------------------------
def resolve_city(text):
    """Turn what the user typed (name, part of a name, or list number) into a city."""
    text = text.strip().lower()
    if text.isdigit() and 1 <= int(text) <= len(cities):
        return cities[int(text) - 1]
    for city in cities:
        if city.lower() == text:
            return city
    if text != "":
        starts = [c for c in cities if c.lower().startswith(text)]
        if len(starts) == 1:
            return starts[0]
    return None


def print_city_list():
    print("Cities: " + ", ".join(f"{i}. {c}" for i, c in enumerate(cities, start=1)))


def ask_city(allow_all=False):
    """Ask for one city. Returns the city name, 'ALL' (if allowed) or None (cancel)."""
    print_city_list()
    if allow_all:
        prompt = "Enter city name or number (Enter = all cities): "
    else:
        prompt = "Enter city name or number (Enter = cancel): "
    while True:
        text = input(prompt).strip()
        if text == "":
            return "ALL" if allow_all else None
        city = resolve_city(text)
        if city is not None:
            return city
        print("City not found. Please try again.")


def ask_many_cities():
    """Ask for two or more cities (comma separated) or ALL."""
    print_city_list()
    text = input("Enter two or more cities separated by commas (or ALL): ").strip()
    if text.lower() == "all":
        return list(cities)
    chosen = []
    for part in text.split(","):
        if part.strip() == "":
            continue
        city = resolve_city(part)
        if city is None:
            print(f"'{part.strip()}' not found, skipping it.")
        elif city not in chosen:
            chosen.append(city)
    return chosen


def ask_number(prompt):
    while True:
        text = input(prompt).strip()
        try:
            return float(text)
        except ValueError:
            print("Please enter a valid number.")


def ask_date(prompt="Enter date (YYYY-MM-DD, e.g. 2026-11-05): "):
    while True:
        text = input(prompt).strip()
        for date_format in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
            try:
                return pd.to_datetime(text, format=date_format)
            except ValueError:
                pass
        print("Invalid date. Use YYYY-MM-DD (or DD-MM-YYYY).")


def ask_month():
    """Ask for a month (number or name). Only months that exist in the data are accepted."""
    available = data[["Month_No", "Month"]].drop_duplicates().sort_values("Month_No")
    names = list(available["Month"])
    print("Months in the data: " + ", ".join(names))
    while True:
        text = input("Enter month (name or number, Enter = cancel): ").strip().lower()
        if text == "":
            return None
        for number, name in zip(available["Month_No"], available["Month"]):
            if text == str(number) or (len(text) >= 3 and name.lower().startswith(text)):
                return number
        print("No data for that month. Choose from:", ", ".join(names))


def rows_for(city):
    """All rows for a city, or every row when city is 'ALL'."""
    if city == "ALL":
        return data
    return data[data["City"] == city]


# ----------------------------------------------------------------------
# DISPLAY HELPERS
# ----------------------------------------------------------------------
def show_table(rows, limit=25):
    """Print records as a simple text table."""
    if rows.empty:
        print("No records found.")
        return
    total = len(rows)
    print(f"\nRecords found: {total}")
    count = total
    if total > limit:
        answer = input(f"How many rows to show? (Enter = {limit}, 'all' = everything): ").strip().lower()
        if answer == "all":
            count = total
        elif answer.isdigit() and int(answer) > 0:
            count = int(answer)
        else:
            count = limit
    out = rows[TABLE_COLUMNS].head(count).copy()
    out["Date"] = out["Date"].dt.strftime("%Y-%m-%d")
    out = out.rename(columns=HEADERS)
    print(out.to_string(index=False))
    if total > count:
        print(f"... showing {count} of {total} records.")


def city_table(rows):
    """One summary row per city (used by compare and rank)."""
    table = rows.groupby("City").agg(
        avg_temp=("Average_Temperature_C", "mean"),
        highest=("Max_Temperature_C", "max"),
        lowest=("Min_Temperature_C", "min"),
        total_rain=("Rainfall_mm", "sum"),
        rainy_days=("Rainfall_mm", lambda s: int((s > 0).sum())),
        avg_wind=("Wind_Speed_kmh", "mean"),
        max_wind=("Wind_Speed_kmh", "max"),
    )
    return table.round(2)


def climate_description(c):
    """A short plain-English description made from the numbers."""
    avg_temp = c["Average_Temperature_C"].mean()
    rainy_share = (c["Rainfall_mm"] > 0).mean() * 100
    avg_wind = c["Wind_Speed_kmh"].mean()

    temp_word = np.select([avg_temp >= 30, avg_temp >= 26, avg_temp >= 20],
                          ["Hot", "Warm", "Mild"], default="Cool")
    rain_word = np.select([rainy_share == 0, rainy_share <= 20, rainy_share <= 40],
                          ["dry", "mostly dry", "occasionally rainy"], default="often rainy")
    wind_word = np.select([avg_wind >= 12, avg_wind >= 8],
                          ["breezy", "moderate winds"], default="light winds")
    return f"{temp_word} and {rain_word}, with {wind_word}"


# ----------------------------------------------------------------------
# MENU ACTIONS
# ----------------------------------------------------------------------
def action_list_cities():
    print("\nAvailable cities:")
    for i, city in enumerate(cities, start=1):
        days = int((data["City"] == city).sum())
        print(f" {i}. {city}  ({days} days of data)")
    print(f"\nData covers {data['Date'].min():%Y-%m-%d} to {data['Date'].max():%Y-%m-%d}.")


def action_climate_overview():
    city = ask_city()
    if city is None:
        return
    c = data[data["City"] == city]

    hottest = c.loc[c["Max_Temperature_C"].idxmax()]
    coldest = c.loc[c["Min_Temperature_C"].idxmin()]
    wettest = c.loc[c["Rainfall_mm"].idxmax()]
    windiest = c.loc[c["Wind_Speed_kmh"].idxmax()]
    rainy_days = int((c["Rainfall_mm"] > 0).sum())

    print(f"\n=== CLIMATE OVERVIEW: {city.upper()} ===")
    print(f"Period : {c['Date'].min():%Y-%m-%d} to {c['Date'].max():%Y-%m-%d} ({len(c)} days)")
    print(f"Climate: {climate_description(c)}")
    print("(description is generated by this program from the numbers)\n")

    lines = [
        ("Average temperature (C)", round(c["Average_Temperature_C"].mean(), 2)),
        ("Average of daily max (C)", round(c["Max_Temperature_C"].mean(), 2)),
        ("Average of daily min (C)", round(c["Min_Temperature_C"].mean(), 2)),
        ("Highest temperature (C)", f"{hottest['Max_Temperature_C']} on {hottest['Date']:%Y-%m-%d}"),
        ("Lowest temperature (C)", f"{coldest['Min_Temperature_C']} on {coldest['Date']:%Y-%m-%d}"),
        ("Total rainfall (mm)", round(c["Rainfall_mm"].sum(), 1)),
        ("Rainy days", f"{rainy_days} of {len(c)}"),
        ("Wettest day", f"{wettest['Rainfall_mm']} mm on {wettest['Date']:%Y-%m-%d}"),
        ("Average wind speed (km/h)", round(c["Wind_Speed_kmh"].mean(), 2)),
        ("Windiest day", f"{windiest['Wind_Speed_kmh']} km/h on {windiest['Date']:%Y-%m-%d}"),
    ]
    for label, value in lines:
        print(f"{label:<28}: {value}")

    print("\n--- Days by condition ---")
    counts = c["Condition"].value_counts()
    print(counts.rename_axis("Condition").reset_index(name="Days").to_string(index=False))

    print("\n--- Month by month ---")
    monthly = c.groupby("Month_No").agg(
        Month=("Month", "first"),
        Days=("Date", "count"),
        Avg_Temp=("Average_Temperature_C", "mean"),
        Highest=("Max_Temperature_C", "max"),
        Lowest=("Min_Temperature_C", "min"),
        Total_Rain=("Rainfall_mm", "sum"),
        Avg_Wind=("Wind_Speed_kmh", "mean"),
    ).round(2)
    print(monthly.to_string(index=False))


def action_all_records():
    city = ask_city()
    if city is not None:
        show_table(rows_for(city))


def action_city_and_date():
    city = ask_city()
    if city is None:
        return
    date = ask_date()
    result = data[(data["City"] == city) & (data["Date"] == date)]
    if result.empty:
        print(f"No record for {city} on {date:%Y-%m-%d}. "
              f"Data covers {data['Date'].min():%Y-%m-%d} to {data['Date'].max():%Y-%m-%d}.")
    else:
        show_table(result)


def action_city_and_month():
    city = ask_city()
    if city is None:
        return
    month = ask_month()
    if month is None:
        return
    show_table(data[(data["City"] == city) & (data["Month_No"] == month)])


def action_date_range():
    city = ask_city(allow_all=True)
    start = ask_date("Enter start date (YYYY-MM-DD): ")
    end = ask_date("Enter end date   (YYYY-MM-DD): ")
    if start > end:
        start, end = end, start
    rows = rows_for(city)
    show_table(rows[rows["Date"].between(start, end)])


def filter_by_range(column, label):
    city = ask_city(allow_all=True)
    low = ask_number(f"Enter minimum {label}: ")
    high = ask_number(f"Enter maximum {label}: ")
    if low > high:
        low, high = high, low
    rows = rows_for(city)
    show_table(rows[rows[column].between(low, high)])


def action_temperature_filter():
    print("Which temperature?  1 = Average   2 = Daily maximum   3 = Daily minimum")
    choice = input("Choice (Enter = 1): ").strip()
    column = {"2": "Max_Temperature_C", "3": "Min_Temperature_C"}.get(choice, "Average_Temperature_C")
    filter_by_range(column, "temperature (C)")


def action_rainfall_filter():
    filter_by_range("Rainfall_mm", "rainfall (mm)")


def action_wind_filter():
    filter_by_range("Wind_Speed_kmh", "wind speed (km/h)")


def action_condition_search():
    available = sorted(data["Condition"].unique())
    print("Conditions: " + ", ".join(f"{i}. {c}" for i, c in enumerate(available, start=1)))
    text = input("Enter condition name or number (Enter = cancel): ").strip().lower()
    if text == "":
        return
    chosen = None
    if text.isdigit() and 1 <= int(text) <= len(available):
        chosen = available[int(text) - 1]
    else:
        for name in available:
            if name.lower() == text:
                chosen = name
    if chosen is None:
        print("Condition not found.")
        return
    city = ask_city(allow_all=True)
    rows = rows_for(city)
    show_table(rows[rows["Condition"] == chosen])


def action_extreme_days():
    options = {
        "1": ("Hottest days", "Max_Temperature_C", False),
        "2": ("Coldest days", "Min_Temperature_C", True),
        "3": ("Wettest days", "Rainfall_mm", False),
        "4": ("Windiest days", "Wind_Speed_kmh", False),
    }
    print("1 = Hottest days   2 = Coldest days   3 = Wettest days   4 = Windiest days")
    choice = input("Choice (Enter = cancel): ").strip()
    if choice not in options:
        return
    title, column, ascending = options[choice]
    city = ask_city(allow_all=True)
    text = input("How many days to show? (Enter = 5): ").strip()
    count = int(text) if text.isdigit() and int(text) > 0 else 5

    rows = rows_for(city).sort_values(column, ascending=ascending).head(count)
    scope = "all cities" if city == "ALL" else city
    print(f"\n{title} - {scope}")
    show_table(rows, limit=max(count, 25))


def action_compare():
    chosen = ask_many_cities()
    if len(chosen) < 2:
        print("Please choose at least two valid cities.")
        return
    table = city_table(data[data["City"].isin(chosen)])
    table = table.rename(columns=CITY_TABLE_HEADERS)
    print("\n=== CITY COMPARISON ===")
    print(table.to_string())


def action_rank():
    options = {
        "1": ("Average temperature (hottest first)", "avg_temp", False),
        "2": ("Highest temperature (hottest first)", "highest", False),
        "3": ("Lowest temperature (coldest first)", "lowest", True),
        "4": ("Total rainfall (wettest first)", "total_rain", False),
        "5": ("Average wind speed (windiest first)", "avg_wind", False),
    }
    for key, (title, _, _) in options.items():
        print(f" {key} = {title}")
    choice = input("Choice (Enter = cancel): ").strip()
    if choice not in options:
        return
    title, column, ascending = options[choice]
    table = city_table(data).sort_values(column, ascending=ascending)
    table.insert(0, "Rank", np.arange(1, len(table) + 1))
    table = table.rename(columns=CITY_TABLE_HEADERS)
    print(f"\n=== CITIES RANKED BY: {title} ===")
    print(table.to_string())


# ----------------------------------------------------------------------
# MAIN MENU
# ----------------------------------------------------------------------
MENU = [
    ("1", "List all cities and data period", action_list_cities),
    ("2", "Climate overview of a city (temperature, rain, wind)", action_climate_overview),
    ("3", "Show all records of a city", action_all_records),
    ("4", "Search by city and date", action_city_and_date),
    ("5", "Search by city and month", action_city_and_month),
    ("6", "Search by date range", action_date_range),
    ("7", "Filter by temperature range", action_temperature_filter),
    ("8", "Filter by rainfall range", action_rainfall_filter),
    ("9", "Filter by wind speed range", action_wind_filter),
    ("10", "Search by weather condition (Rain, Hot, Windy ...)", action_condition_search),
    ("11", "Hottest / coldest / wettest / windiest days", action_extreme_days),
    ("12", "Compare cities", action_compare),
    ("13", "Rank all cities", action_rank),
]


def show_menu():
    print("\n" + "=" * 62)
    print("                    WEATHER SEARCH")
    print("=" * 62)
    for key, text, _ in MENU:
        print(f" {key:>2}. {text}")
    print("  0. Exit")
    print("=" * 62)


if __name__ == "__main__":
    data = load_data()
    cities = sorted(data["City"].unique())

    try:
        while True:
            show_menu()
            choice = input("Enter your choice (0-13): ").strip()
            if choice == "0":
                print("Goodbye!")
                break
            for key, _, action in MENU:
                if choice == key:
                    action()
                    break
            else:
                print("Invalid choice. Please enter a number from 0 to 13.")
                continue
            input("\nPress Enter to return to the menu...")
    except (KeyboardInterrupt, EOFError):
        print("\nProgram stopped.")
