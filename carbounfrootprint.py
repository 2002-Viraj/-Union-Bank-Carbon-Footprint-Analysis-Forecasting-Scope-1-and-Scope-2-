import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# =========================
# 1. FILE PATHS
# =========================
file_2024 = r"C:\Users\viraj\OneDrive\Desktop\FS 3002 - Service Learning\Carboun printer forcast using python\Carbon Footprint-final 2024.xlsx"
file_2025 = r"C:\Users\viraj\OneDrive\Desktop\FS 3002 - Service Learning\Carboun printer forcast using python\Carbon Footprint-final 2025.xlsx"

# =========================
# 2. READ SUMMARY SHEETS
# =========================
summary_2024 = pd.read_excel(file_2024, sheet_name="Summary", header=None)
summary_2025 = pd.read_excel(file_2025, sheet_name="Summary", header=None)

sources_2024 = pd.DataFrame({
    "Source": [
        "Onsite Diesel Generators",
        "Refrigerant Leakage",
        "Fire Extinguishers",
        "Company Vehicles",
        "Electricity"
    ],
    "Emission_tCO2e": [
        summary_2024.iloc[2, 2],
        summary_2024.iloc[3, 2],
        summary_2024.iloc[4, 2],
        summary_2024.iloc[5, 2],
        summary_2024.iloc[7, 2]
    ]
})

sources_2025 = pd.DataFrame({
    "Source": [
        "Onsite Diesel Generators",
        "Refrigerant Leakage",
        "Fire Extinguishers",
        "Company Vehicles",
        "Electricity"
    ],
    "Emission_tCO2e": [
        summary_2025.iloc[2, 2],
        summary_2025.iloc[3, 2],
        summary_2025.iloc[4, 2],
        summary_2025.iloc[5, 2],
        summary_2025.iloc[7, 2]
    ]
})

sources_2024["Emission_tCO2e"] = pd.to_numeric(sources_2024["Emission_tCO2e"], errors="coerce")
sources_2025["Emission_tCO2e"] = pd.to_numeric(sources_2025["Emission_tCO2e"], errors="coerce")

comparison = sources_2024.merge(
    sources_2025,
    on="Source",
    suffixes=("_2024", "_2025")
)

comparison["Change"] = comparison["Emission_tCO2e_2025"] - comparison["Emission_tCO2e_2024"]
comparison["Percent_Change"] = (comparison["Change"] / comparison["Emission_tCO2e_2024"]) * 100

total_2024 = sources_2024["Emission_tCO2e"].sum()
total_2025 = sources_2025["Emission_tCO2e"].sum()

comparison["Share_2024_%"] = (comparison["Emission_tCO2e_2024"] / total_2024) * 100
comparison["Share_2025_%"] = (comparison["Emission_tCO2e_2025"] / total_2025) * 100

print("\n=== 2024 vs 2025 Comparison ===")
print(comparison)

comparison.to_excel("Carbon_Comparison_2024_2025.xlsx", index=False)

# =========================
# 3. READ ELECTRICITY DATA
# =========================
elec_2024 = pd.read_excel(file_2024, sheet_name="Eletricity connsumtion actual ", header=None)
elec_2025 = pd.read_excel(file_2025, sheet_name="Electricity- 2025", header=None)

months = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

monthly_2024 = elec_2024.iloc[3:, 3:15].copy()
monthly_2024.columns = months
monthly_2024 = monthly_2024.apply(pd.to_numeric, errors="coerce")
monthly_totals_2024 = monthly_2024.sum(axis=0)

monthly_2025 = elec_2025.iloc[1:, 3:15].copy()
monthly_2025.columns = months
monthly_2025 = monthly_2025.apply(pd.to_numeric, errors="coerce")
monthly_totals_2025 = monthly_2025.sum(axis=0)

monthly_series = pd.concat([monthly_totals_2024, monthly_totals_2025], axis=0)
monthly_series.index = pd.date_range(start="2024-01-01", periods=24, freq="MS")
monthly_series = monthly_series.astype(float)

print("\n=== Actual Monthly Electricity Series (2024-2025) ===")
print(monthly_series)

# =========================
# 4. PLOT HISTORICAL TREND
# =========================
plt.figure(figsize=(12, 5))
plt.plot(monthly_series.index, monthly_series.values, marker="o")
plt.title("Monthly Electricity Consumption (2024-2025)")
plt.xlabel("Month")
plt.ylabel("Electricity Consumption (kWh)")
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# =========================
# 5. FORECAST FOR 2026 ONLY
# =========================
forecast_months = 12

model = ExponentialSmoothing(
    monthly_series,
    trend="add",
    seasonal="add",
    seasonal_periods=12
).fit()

forecast_kwh = model.forecast(forecast_months)

forecast_df = pd.DataFrame({
    "Date": forecast_kwh.index,
    "Forecast_kWh": forecast_kwh.values
})

forecast_df["Month"] = forecast_df["Date"].dt.strftime("%Y-%m")

print("\n=== 2026 Monthly Forecast ===")
print(forecast_df[["Month", "Forecast_kWh"]])

forecast_df[["Month", "Forecast_kWh"]].to_excel("Electricity_Forecast_2026.xlsx", index=False)

# =========================
# 5A. PLOT ACTUAL + FORECAST WITH CONNECTOR LINE
# =========================
plt.figure(figsize=(15, 6))

# actual line
plt.plot(
    monthly_series.index,
    monthly_series.values,
    marker="o",
    linewidth=2,
    label="Actual (2024-2025)"
)

# connector line from 2025-12 to 2026-01
plt.plot(
    [monthly_series.index[-1], forecast_kwh.index[0]],
    [monthly_series.iloc[-1], forecast_kwh.iloc[0]],
    linestyle="--",
    linewidth=2,
    color="black"
)

# forecast line
plt.plot(
    forecast_kwh.index,
    forecast_kwh.values,
    marker="o",
    linestyle="--",
    linewidth=2,
    label="Forecast (2026)"
)

plt.title("Monthly Electricity Consumption (2024-2025) with 2026 Forecast")
plt.xlabel("Month")
plt.ylabel("Electricity Consumption (kWh)")
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# =========================
# 6. CONVERT FORECAST TO EMISSIONS
# =========================
grid_emission_factor = 0.4173  # kg CO2e per kWh

forecast_df["Forecast_Electricity_tCO2e"] = (
    forecast_df["Forecast_kWh"] * grid_emission_factor
) / 1000

forecast_df["Year"] = forecast_df["Date"].dt.year

annual_electricity_forecast = forecast_df.groupby("Year")[["Forecast_kWh", "Forecast_Electricity_tCO2e"]].sum().reset_index()

print("\n=== Annual Electricity Forecast ===")
print(annual_electricity_forecast)

annual_electricity_forecast.to_excel("Annual_Electricity_Forecast_2026.xlsx", index=False)

# =========================
# 7. BASELINE TOTAL CARBON FOOTPRINT
# =========================
baseline_2025 = dict(zip(sources_2025["Source"], sources_2025["Emission_tCO2e"]))

diesel_base = baseline_2025["Onsite Diesel Generators"]
refrigerant_base = baseline_2025["Refrigerant Leakage"]
fire_base = baseline_2025["Fire Extinguishers"]
vehicle_base = baseline_2025["Company Vehicles"]

baseline_total = annual_electricity_forecast.copy()
baseline_total["Diesel_tCO2e"] = diesel_base
baseline_total["Refrigerant_tCO2e"] = refrigerant_base
baseline_total["Fire_tCO2e"] = fire_base
baseline_total["Vehicle_tCO2e"] = vehicle_base

baseline_total["Total_Baseline_tCO2e"] = (
    baseline_total["Forecast_Electricity_tCO2e"] +
    baseline_total["Diesel_tCO2e"] +
    baseline_total["Refrigerant_tCO2e"] +
    baseline_total["Fire_tCO2e"] +
    baseline_total["Vehicle_tCO2e"]
)

# =========================
# 8. SCENARIO ANALYSIS
# =========================
scenario_df = baseline_total.copy()

scenario_df["Scenario1_Total_tCO2e"] = (
    scenario_df["Forecast_Electricity_tCO2e"] * 0.95 +
    diesel_base +
    refrigerant_base +
    fire_base +
    vehicle_base
)

scenario_df["Scenario2_Total_tCO2e"] = (
    scenario_df["Forecast_Electricity_tCO2e"] +
    diesel_base +
    (refrigerant_base * 0.90) +
    fire_base +
    vehicle_base
)

scenario_df["Scenario3_Total_tCO2e"] = (
    scenario_df["Forecast_Electricity_tCO2e"] +
    (diesel_base * 0.90) +
    refrigerant_base +
    fire_base +
    (vehicle_base * 0.92)
)

scenario_df["Scenario4_Total_tCO2e"] = (
    scenario_df["Forecast_Electricity_tCO2e"] * 0.92 +
    (diesel_base * 0.88) +
    (refrigerant_base * 0.85) +
    fire_base +
    (vehicle_base * 0.90)
)

scenario_df["Scenario1_Saving"] = scenario_df["Total_Baseline_tCO2e"] - scenario_df["Scenario1_Total_tCO2e"]
scenario_df["Scenario2_Saving"] = scenario_df["Total_Baseline_tCO2e"] - scenario_df["Scenario2_Total_tCO2e"]
scenario_df["Scenario3_Saving"] = scenario_df["Total_Baseline_tCO2e"] - scenario_df["Scenario3_Total_tCO2e"]
scenario_df["Scenario4_Saving"] = scenario_df["Total_Baseline_tCO2e"] - scenario_df["Scenario4_Total_tCO2e"]

print("\n=== Scenario Analysis ===")
print(scenario_df[[
    "Year",
    "Total_Baseline_tCO2e",
    "Scenario1_Total_tCO2e", "Scenario1_Saving",
    "Scenario2_Total_tCO2e", "Scenario2_Saving",
    "Scenario3_Total_tCO2e", "Scenario3_Saving",
    "Scenario4_Total_tCO2e", "Scenario4_Saving"
]])

scenario_df.to_excel("Carbon_Forecast_Scenarios_2026.xlsx", index=False)

# =========================
# 9. CORRECT BASELINE VS SCENARIO PLOT
# =========================
scenario_plot_df = pd.DataFrame({
    "Scenario": [
        "Baseline 2026",
        "Scenario 1\nElectricity -5%",
        "Scenario 2\nRefrigerant -10%",
        "Scenario 3\nDiesel -10% & Vehicle -8%",
        "Scenario 4\nCombined Strategy"
    ],
    "Total_tCO2e": [
        float(scenario_df["Total_Baseline_tCO2e"].iloc[0]),
        float(scenario_df["Scenario1_Total_tCO2e"].iloc[0]),
        float(scenario_df["Scenario2_Total_tCO2e"].iloc[0]),
        float(scenario_df["Scenario3_Total_tCO2e"].iloc[0]),
        float(scenario_df["Scenario4_Total_tCO2e"].iloc[0])
    ]
})

print("\n=== Baseline vs Scenario Values ===")
print(scenario_plot_df)

plt.figure(figsize=(13, 6))
plt.plot(
    scenario_plot_df["Scenario"],
    scenario_plot_df["Total_tCO2e"],
    marker="o",
    linewidth=2
)
plt.title("2026 Baseline vs Scenarios")
plt.xlabel("Scenario")
plt.ylabel("Total Carbon Footprint (tCO2e)")
plt.grid(True)
plt.xticks(rotation=15)
plt.tight_layout()
plt.show()

# =========================
# 10. BEST SCENARIO
# =========================
scenario_values = {
    "Scenario 1": float(scenario_df["Scenario1_Total_tCO2e"].iloc[0]),
    "Scenario 2": float(scenario_df["Scenario2_Total_tCO2e"].iloc[0]),
    "Scenario 3": float(scenario_df["Scenario3_Total_tCO2e"].iloc[0]),
    "Scenario 4": float(scenario_df["Scenario4_Total_tCO2e"].iloc[0])
}

best_name = min(scenario_values, key=scenario_values.get)
best_value = scenario_values[best_name]

best_scenario_df = pd.DataFrame({
    "Year": [2026],
    "Best_Scenario": [best_name],
    "Lowest_Total_tCO2e": [best_value]
})

best_scenario_df.to_excel("Best_Scenario_2026.xlsx", index=False)

print("\n=== Best Scenario ===")
print(best_scenario_df)

# ==========================================
# 11. GIVEN DATA
# ==========================================
ghg_2024 = 1503.72
ghg_2025 = 1540.46

# ==========================================
# 12. CALCULATE CHANGE
# ==========================================
yearly_change = ghg_2025 - ghg_2024

# ==========================================
# 13. FORECAST 2026
# ==========================================
ghg_2026 = ghg_2025 + yearly_change

# ==========================================
# 14. CREATE DATAFRAME
# ==========================================
actual_years = [2024, 2025]
actual_values = [ghg_2024, ghg_2025]

forecast_years = [2025, 2026]   # include 2025 for connector
forecast_values = [ghg_2025, ghg_2026]

# ==========================================
# 15. PLOT GRAPH
# ==========================================
plt.figure(figsize=(10, 5))

# 🔵 Actual line (solid blue)
plt.plot(
    actual_years,
    actual_values,
    marker="o",
    linewidth=2,
    label="Actual",
)

# 🔴 Forecast line (dashed red)
plt.plot(
    forecast_years,
    forecast_values,
    marker="o",
    linestyle="--",
    linewidth=2,
    label="Forecast",
)

# ==========================================
# 16. VALUE LABELS
# ==========================================
for x, y in zip(actual_years, actual_values):
    plt.text(x, y + 2, round(y, 2), ha='center')

for x, y in zip([2026], [ghg_2026]):
    plt.text(x, y + 2, round(y, 2), ha='center')

# ==========================================
# 17. GRAPH SETTINGS
# ==========================================
plt.title("Total GHG Emission Trend with 2026 Forecast")
plt.xlabel("Year")
plt.ylabel("Total GHG Emission (tCO2e)")
plt.grid(True)
plt.legend()
plt.tight_layout()
plt.show()

# =========================================================
# 18. FILE PATHS
# =========================================================
file_2024 = r"C:\Users\viraj\OneDrive\Desktop\FS 3002 - Service Learning\Carboun printer forcast using python\Carbon Footprint-final 2024.xlsx"
file_2025 = r"C:\Users\viraj\OneDrive\Desktop\FS 3002 - Service Learning\Carboun printer forcast using python\Carbon Footprint-final 2025.xlsx"

# =========================================================
# 19. READ ELECTRICITY SHEETS
# =========================================================
e24 = pd.read_excel(file_2024, sheet_name="Eletricity connsumtion actual ", header=None)
e25 = pd.read_excel(file_2025, sheet_name="Electricity- 2025", header=None)

months = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

# =========================================================
# 20. EXTRACT 2024 DATA
#    Offices & Branches = rows 3 to 66
#    Off-site ATMs      = rows 69 to 108
# =========================================================
offices_2024 = e24.iloc[3:67, [0, 14]].copy()
offices_2024.columns = ["Location", "Total_kWh"]

atms_2024 = e24.iloc[69:109, [0, 14]].copy()
atms_2024.columns = ["Location", "Total_kWh"]

offices_2024["Total_kWh"] = pd.to_numeric(offices_2024["Total_kWh"], errors="coerce")
atms_2024["Total_kWh"] = pd.to_numeric(atms_2024["Total_kWh"], errors="coerce")

# =========================================================
# 21. EXTRACT 2025 DATA
#    Offices & Branches = rows 1 to 61
#    Off-site ATMs      = rows 65 to 104
# =========================================================
offices_2025 = e25.iloc[1:62, [0, 14]].copy()
offices_2025.columns = ["Location", "Total_kWh"]

atms_2025 = e25.iloc[65:105, [0, 14]].copy()
atms_2025.columns = ["Location", "Total_kWh"]

offices_2025["Total_kWh"] = pd.to_numeric(offices_2025["Total_kWh"], errors="coerce")
atms_2025["Total_kWh"] = pd.to_numeric(atms_2025["Total_kWh"], errors="coerce")

# =========================================================
# 22. CALCULATE TOTAL ELECTRICITY
# =========================================================
offices_total_2024 = offices_2024["Total_kWh"].sum()
atms_total_2024 = atms_2024["Total_kWh"].sum()

offices_total_2025 = offices_2025["Total_kWh"].sum()
atms_total_2025 = atms_2025["Total_kWh"].sum()

# =========================================================
# 23. CALCULATE MONTHLY AVERAGE ELECTRICITY
# =========================================================
offices_avg_month_2024 = offices_total_2024 / 12
atms_avg_month_2024 = atms_total_2024 / 12

offices_avg_month_2025 = offices_total_2025 / 12
atms_avg_month_2025 = atms_total_2025 / 12

# =========================================================
# 24. CALCULATE TWO-YEAR AVERAGE
# =========================================================
offices_two_year_avg_total = (offices_total_2024 + offices_total_2025) / 2
atms_two_year_avg_total = (atms_total_2024 + atms_total_2025) / 2

offices_two_year_avg_month = (offices_avg_month_2024 + offices_avg_month_2025) / 2
atms_two_year_avg_month = (atms_avg_month_2024 + atms_avg_month_2025) / 2

# =========================================================
# 25. CREATE SUMMARY TABLE
# =========================================================
summary_df = pd.DataFrame({
    "Category": ["Offices and Branches", "Off-site ATMs"],
    "2024_Total_kWh": [offices_total_2024, atms_total_2024],
    "2025_Total_kWh": [offices_total_2025, atms_total_2025],
    "Two_Year_Avg_Total_kWh": [offices_two_year_avg_total, atms_two_year_avg_total],
    "2024_Avg_Monthly_kWh": [offices_avg_month_2024, atms_avg_month_2024],
    "2025_Avg_Monthly_kWh": [offices_avg_month_2025, atms_avg_month_2025],
    "Two_Year_Avg_Monthly_kWh": [offices_two_year_avg_month, atms_two_year_avg_month]
})

print("\n=== Electricity Summary ===")
print(summary_df)

summary_df.to_excel("Electricity_Offices_vs_ATMs_Summary.xlsx", index=False)

# =========================================================
# 26. FIND WHICH CATEGORY USES MORE ELECTRICITY
# =========================================================
if offices_two_year_avg_total > atms_two_year_avg_total:
    print("\nOffices and Branches use more electricity on average across the two years.")
else:
    print("\nOff-site ATMs use more electricity on average across the two years.")

# =========================================================
# 27. GRAPH 1 - 2024 vs 2025 TOTAL ELECTRICITY BY CATEGORY
# =========================================================
x = np.arange(len(summary_df["Category"]))
width = 0.35

plt.figure(figsize=(10, 6))
plt.bar(x - width/2, summary_df["2024_Total_kWh"], width, label="2024")
plt.bar(x + width/2, summary_df["2025_Total_kWh"], width, label="2025")

plt.xticks(x, summary_df["Category"])
plt.title("Total Electricity Consumption by Category (2024 vs 2025)")
plt.xlabel("Category")
plt.ylabel("Total Electricity Consumption (kWh)")
plt.legend()
plt.grid(axis="y")
plt.tight_layout()
plt.show()

# =========================================================
# 28. GRAPH 2 - TWO-YEAR AVERAGE TOTAL ELECTRICITY
# =========================================================
plt.figure(figsize=(8, 5))
plt.bar(summary_df["Category"], summary_df["Two_Year_Avg_Total_kWh"])

for x_label, y_value in zip(summary_df["Category"], summary_df["Two_Year_Avg_Total_kWh"]):
    plt.text(x_label, y_value, f"{y_value:,.0f}", ha="center", va="bottom")

plt.title("Two-Year Average Total Electricity Consumption")
plt.xlabel("Category")
plt.ylabel("Average Total Electricity Consumption (kWh)")
plt.grid(axis="y")
plt.tight_layout()
plt.show()

# =========================================================
# 29. GRAPH 3 - TWO-YEAR AVERAGE MONTHLY ELECTRICITY
# =========================================================
plt.figure(figsize=(8, 5))
plt.plot(summary_df["Category"], summary_df["Two_Year_Avg_Monthly_kWh"], marker="o", linewidth=2)

for x_label, y_value in zip(summary_df["Category"], summary_df["Two_Year_Avg_Monthly_kWh"]):
    plt.text(x_label, y_value, f"{y_value:,.0f}", ha="center", va="bottom")

plt.title("Two-Year Average Monthly Electricity Consumption")
plt.xlabel("Category")
plt.ylabel("Average Monthly Electricity Consumption (kWh)")
plt.grid(True)
plt.tight_layout()
plt.show()

# =========================================================
# 30. READ ELECTRICITY SHEETS
# =========================================================
e24 = pd.read_excel(file_2024, sheet_name="Eletricity connsumtion actual ", header=None)
e25 = pd.read_excel(file_2025, sheet_name="Electricity- 2025", header=None)

# =========================================================
# 31. EXTRACT OFFICES / BRANCHES AND OFF-SITE ATMS
#    These row ranges follow your current file structure
# =========================================================

# -------- 2024 --------
offices_2024 = e24.iloc[3:67, [0, 14]].copy()
offices_2024.columns = ["Location", "Total_kWh"]

atms_2024 = e24.iloc[69:109, [0, 14]].copy()
atms_2024.columns = ["Location", "Total_kWh"]

# -------- 2025 --------
offices_2025 = e25.iloc[1:62, [0, 14]].copy()
offices_2025.columns = ["Location", "Total_kWh"]

atms_2025 = e25.iloc[65:105, [0, 14]].copy()
atms_2025.columns = ["Location", "Total_kWh"]

# =========================================================
# 32. CLEAN DATA
# =========================================================
def clean_data(df):
    df = df.copy()
    df["Location"] = df["Location"].astype(str).str.strip()
    df["Total_kWh"] = pd.to_numeric(df["Total_kWh"], errors="coerce")
    df = df.dropna(subset=["Location", "Total_kWh"])
    df = df[df["Location"] != ""]
    return df

offices_2024 = clean_data(offices_2024)
offices_2025 = clean_data(offices_2025)
atms_2024 = clean_data(atms_2024)
atms_2025 = clean_data(atms_2025)

# =========================================================
# 33. TOP ELECTRICITY USERS BY YEAR
# =========================================================
top_offices_2024 = offices_2024.sort_values("Total_kWh", ascending=False).head(10)
top_offices_2025 = offices_2025.sort_values("Total_kWh", ascending=False).head(10)

top_atms_2024 = atms_2024.sort_values("Total_kWh", ascending=False).head(10)
top_atms_2025 = atms_2025.sort_values("Total_kWh", ascending=False).head(10)

print("\n=== Top 10 Offices / Branches - 2024 ===")
print(top_offices_2024)

print("\n=== Top 10 Offices / Branches - 2025 ===")
print(top_offices_2025)

print("\n=== Top 10 Off-site ATMs - 2024 ===")
print(top_atms_2024)

print("\n=== Top 10 Off-site ATMs - 2025 ===")
print(top_atms_2025)

# =========================================================
# 34. SAVE YEARLY TOP USERS
# =========================================================
top_offices_2024.to_excel("Top_10_Offices_2024.xlsx", index=False)
top_offices_2025.to_excel("Top_10_Offices_2025.xlsx", index=False)
top_atms_2024.to_excel("Top_10_ATMs_2024.xlsx", index=False)
top_atms_2025.to_excel("Top_10_ATMs_2025.xlsx", index=False)

# =========================================================
# 35. MERGE 2024 AND 2025 TO FIND TWO-YEAR AVERAGE
# =========================================================
offices_2024_ren = offices_2024.rename(columns={"Total_kWh": "Total_kWh_2024"})
offices_2025_ren = offices_2025.rename(columns={"Total_kWh": "Total_kWh_2025"})

atms_2024_ren = atms_2024.rename(columns={"Total_kWh": "Total_kWh_2024"})
atms_2025_ren = atms_2025.rename(columns={"Total_kWh": "Total_kWh_2025"})

offices_combined = pd.merge(offices_2024_ren, offices_2025_ren, on="Location", how="outer")
atms_combined = pd.merge(atms_2024_ren, atms_2025_ren, on="Location", how="outer")

offices_combined["Total_kWh_2024"] = offices_combined["Total_kWh_2024"].fillna(0)
offices_combined["Total_kWh_2025"] = offices_combined["Total_kWh_2025"].fillna(0)

atms_combined["Total_kWh_2024"] = atms_combined["Total_kWh_2024"].fillna(0)
atms_combined["Total_kWh_2025"] = atms_combined["Total_kWh_2025"].fillna(0)

offices_combined["Two_Year_Average_kWh"] = (
    offices_combined["Total_kWh_2024"] + offices_combined["Total_kWh_2025"]
) / 2

atms_combined["Two_Year_Average_kWh"] = (
    atms_combined["Total_kWh_2024"] + atms_combined["Total_kWh_2025"]
) / 2

# =========================================================
# 36. FIND TOP ELECTRICITY USERS ACROSS BOTH YEARS
# =========================================================
top_offices_avg = offices_combined.sort_values("Two_Year_Average_kWh", ascending=False).head(10)
top_atms_avg = atms_combined.sort_values("Two_Year_Average_kWh", ascending=False).head(10)

print("\n=== Top 10 Offices / Branches by Two-Year Average ===")
print(top_offices_avg)

print("\n=== Top 10 Off-site ATMs by Two-Year Average ===")
print(top_atms_avg)

top_offices_avg.to_excel("Top_10_Offices_Two_Year_Average.xlsx", index=False)
top_atms_avg.to_excel("Top_10_ATMs_Two_Year_Average.xlsx", index=False)

# =========================================================
# 37. IDENTIFY THE HIGHEST ELECTRICITY USER
# =========================================================
highest_office = top_offices_avg.iloc[0]
highest_atm = top_atms_avg.iloc[0]

print("\n=== Highest Electricity Using Office / Branch ===")
print(highest_office)

print("\n=== Highest Electricity Using Off-site ATM ===")
print(highest_atm)

# =========================================================
# 38. GRAPH - TOP 10 OFFICES / BRANCHES BY TWO-YEAR AVERAGE
# =========================================================
plt.figure(figsize=(12, 6))
plt.barh(top_offices_avg["Location"], top_offices_avg["Two_Year_Average_kWh"])
plt.title("Top 10 Offices / Branches by Two-Year Average Electricity Use")
plt.xlabel("Average Electricity Consumption (kWh)")
plt.ylabel("Office / Branch")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# =========================================================
# 37. GRAPH - TOP 10 OFF-SITE ATMS BY TWO-YEAR AVERAGE
# =========================================================
plt.figure(figsize=(12, 6))
plt.barh(top_atms_avg["Location"], top_atms_avg["Two_Year_Average_kWh"])
plt.title("Top 10 Off-site ATMs by Two-Year Average Electricity Use")
plt.xlabel("Average Electricity Consumption (kWh)")
plt.ylabel("Off-site ATM")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# =========================================================
# 40. OPTIONAL - COMBINED TOP USERS (OFFICES + ATMS TOGETHER)
# =========================================================
offices_combined["Category"] = "Office / Branch"
atms_combined["Category"] = "Off-site ATM"

all_locations = pd.concat([
    offices_combined[["Location", "Category", "Two_Year_Average_kWh"]],
    atms_combined[["Location", "Category", "Two_Year_Average_kWh"]]
], axis=0)

top_all_locations = all_locations.sort_values("Two_Year_Average_kWh", ascending=False).head(15)

print("\n=== Top 15 Overall Electricity Users ===")
print(top_all_locations)

top_all_locations.to_excel("Top_15_Overall_Electricity_Users.xlsx", index=False)

plt.figure(figsize=(13, 7))
plt.barh(top_all_locations["Location"], top_all_locations["Two_Year_Average_kWh"])
plt.title("Top 15 Overall Electricity Users (Offices / Branches + Off-site ATMs)")
plt.xlabel("Two-Year Average Electricity Consumption (kWh)")
plt.ylabel("Location")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# =========================================================
# 41. USE PREVIOUS COMBINED DATA (FROM YOUR CODE)
# =========================================================
# Ensure you already created:
# offices_combined, atms_combined

# Add category labels
offices_combined["Category"] = "Office / Branch"
atms_combined["Category"] = "Off-site ATM"

# Combine both
all_locations = pd.concat([
    offices_combined[["Location", "Category", "Two_Year_Average_kWh"]],
    atms_combined[["Location", "Category", "Two_Year_Average_kWh"]]
], axis=0)

# =========================================================
# 42. SORT BY ELECTRICITY USAGE
# =========================================================
pareto_df = all_locations.sort_values(
    "Two_Year_Average_kWh", ascending=False
).reset_index(drop=True)

# =========================================================
# 43. CALCULATE CUMULATIVE %
# =========================================================
total_consumption = pareto_df["Two_Year_Average_kWh"].sum()

pareto_df["Cumulative_kWh"] = pareto_df["Two_Year_Average_kWh"].cumsum()
pareto_df["Cumulative_%"] = (pareto_df["Cumulative_kWh"] / total_consumption) * 100

print("\n=== Pareto Table ===")
print(pareto_df.head(15))

# Save table
pareto_df.to_excel("Pareto_Analysis_Electricity.xlsx", index=False)

# =========================================================
# 44. FIND 80% CONTRIBUTION POINT
# =========================================================
pareto_80 = pareto_df[pareto_df["Cumulative_%"] <= 80]

print("\n=== Locations contributing to ~80% electricity ===")
print(pareto_80)

print("\nNumber of locations contributing to 80% usage:", len(pareto_80))

# =========================================================
# 45. PARETO CHART (BAR + LINE)
# =========================================================
fig, ax1 = plt.subplots(figsize=(14, 6))

# Bar chart (electricity usage)
ax1.bar(
    pareto_df["Location"],
    pareto_df["Two_Year_Average_kWh"]
)
ax1.set_ylabel("Electricity Consumption (kWh)")
ax1.set_xlabel("Location")
ax1.set_xticklabels(pareto_df["Location"], rotation=90)

# Line chart (cumulative %)
ax2 = ax1.twinx()
ax2.plot(
    pareto_df["Location"],
    pareto_df["Cumulative_%"],
    marker="o"
)
ax2.set_ylabel("Cumulative Percentage (%)")

# 80% reference line
ax2.axhline(80, linestyle="--")

plt.title("Pareto Analysis of Electricity Consumption (Locations)")
plt.tight_layout()
plt.show()

# ==========================================
# 46. GIVEN TOTAL CARBON FOOTPRINT VALUES
# ==========================================
cf_2024 = 1503.72
cf_2025 = 1540.46

# yearly increase
annual_change = cf_2025 - cf_2024   # 36.74

# ==========================================
# 47. BASELINE FORECAST 2026-2030
#    same chart, corrected realistic values
# ==========================================
years = [2026, 2027, 2028, 2029, 2030]

baseline = []
start_value = cf_2025

for i in range(1, 6):
    baseline.append(start_value + annual_change * i)

# ==========================================
# 48. 2025 SOURCE SHARES / SCENARIO FACTORS
#    based on your 2025 dataset values
# ==========================================
total_2025 = 1540.46
electricity_2025 = 1258.17
refrigerant_2025 = 217.79
diesel_2025 = 18.44
vehicle_2025 = 45.97

# shares
elec_share = electricity_2025 / total_2025
refrig_share = refrigerant_2025 / total_2025
diesel_share = diesel_2025 / total_2025
vehicle_share = vehicle_2025 / total_2025

# scenario factors
# Scenario 1: Electricity -5%
scenario1_factor = 1 - (0.05 * elec_share)

# Scenario 2: Refrigerant -10%
scenario2_factor = 1 - (0.10 * refrig_share)

# Scenario 3: Diesel -10% and Vehicle -8%
scenario3_factor = 1 - ((0.10 * diesel_share) + (0.08 * vehicle_share))

# Scenario 4: Combined Strategy
scenario4_factor = 1 - (
    (0.08 * elec_share) +
    (0.12 * diesel_share) +
    (0.15 * refrig_share) +
    (0.10 * vehicle_share)
)

# ==========================================
# 49. APPLY SCENARIOS TO BASELINE
# ==========================================
scenario1 = [x * scenario1_factor for x in baseline]
scenario2 = [x * scenario2_factor for x in baseline]
scenario3 = [x * scenario3_factor for x in baseline]
scenario4 = [x * scenario4_factor for x in baseline]

# ==========================================
# 50. CREATE TABLE
# ==========================================
df = pd.DataFrame({
    "Year": years,
    "Baseline": baseline,
    "Scenario 1: Electricity -5%": scenario1,
    "Scenario 2: Refrigerant -10%": scenario2,
    "Scenario 3: Diesel + Vehicle Reduction": scenario3,
    "Scenario 4: Combined Strategy": scenario4
})

print("\nCorrected Baseline vs Scenario Table")
print(df.round(2))

df.to_excel("Corrected_Baseline_vs_Scenarios_2026_2030.xlsx", index=False)

# ==========================================
#51. SAME STYLE LINE CHART
# ==========================================
plt.figure(figsize=(12, 6))

plt.plot(df["Year"], df["Baseline"], marker="o", label="Baseline")
plt.plot(df["Year"], df["Scenario 1: Electricity -5%"], marker="o", label="Scenario 1: Electricity -5%")
plt.plot(df["Year"], df["Scenario 2: Refrigerant -10%"], marker="o", label="Scenario 2: Refrigerant -10%")
plt.plot(df["Year"], df["Scenario 3: Diesel + Vehicle Reduction"], marker="o", label="Scenario 3: Diesel + Vehicle Reduction")
plt.plot(df["Year"], df["Scenario 4: Combined Strategy"], marker="o", label="Scenario 4: Combined Strategy")

plt.title("Projected Total Carbon Footprint: Baseline vs Scenarios")
plt.xlabel("Year")
plt.ylabel("Total Carbon Footprint (tCO2e)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()
#....................................................................
# source-wise emissions
sources_2024 = {
    "Onsite Diesel Generators": 40.079789,
    "Refrigerant Leakage": 183.204100,
    "Fire Extinguishers": 0.099000,
    "Company Vehicles": 43.293438,
    "Electricity": 1237.044020
}

sources_2025 = {
    "Onsite Diesel Generators": 18.437006,
    "Refrigerant Leakage": 217.788562,
    "Fire Extinguishers": 0.086955,
    "Company Vehicles": 45.974425,
    "Electricity": 1258.168052
}

# Scope 1 = all except electricity
scope1_2024 = (
    sources_2024["Onsite Diesel Generators"] +
    sources_2024["Refrigerant Leakage"] +
    sources_2024["Fire Extinguishers"] +
    sources_2024["Company Vehicles"]
)

scope1_2025 = (
    sources_2025["Onsite Diesel Generators"] +
    sources_2025["Refrigerant Leakage"] +
    sources_2025["Fire Extinguishers"] +
    sources_2025["Company Vehicles"]
)

# Scope 2 = electricity
scope2_2024 = sources_2024["Electricity"]
scope2_2025 = sources_2025["Electricity"]

# Total emissions
total_2024 = scope1_2024 + scope2_2024
total_2025 = scope1_2025 + scope2_2025

# Percentage increase
percent_increase = ((total_2025 - total_2024) / total_2024) * 100

# DataFrame
quantification_df = pd.DataFrame({
    "Emission Category": ["Scope 1 Emissions", "Scope 2 Emissions", "Total Emissions"],
    "2024 (tCO2e)": [scope1_2024, scope2_2024, total_2024],
    "2025 (tCO2e)": [scope1_2025, scope2_2025, total_2025]
})

print(quantification_df.round(2))
print(f"\nPercentage increase in total emissions from 2024 to 2025: {percent_increase:.2f}%")

# -----------------------------
# Graph: Grouped Bar Chart
# -----------------------------
x = np.arange(len(quantification_df["Emission Category"]))
width = 0.35

plt.figure(figsize=(10, 6))

bars1 = plt.bar(x - width/2, quantification_df["2024 (tCO2e)"], width, label="2024")
bars2 = plt.bar(x + width/2, quantification_df["2025 (tCO2e)"], width, label="2025")

# Value labels on bars
for bar in bars1:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, height + 10, f"{height:.2f}",
             ha='center', va='bottom', fontsize=9)

for bar in bars2:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, height + 10, f"{height:.2f}",
             ha='center', va='bottom', fontsize=9)

# Show 2.44% increase above Total Emissions bars
total_x_2024 = x[2] - width/2
total_x_2025 = x[2] + width/2
total_y_2024 = total_2024
total_y_2025 = total_2025

# line connecting total bars
plt.plot([total_x_2024, total_x_2025], [total_y_2024, total_y_2025], linestyle="--", linewidth=1.5)

# annotation
mid_x = (total_x_2024 + total_x_2025) / 2
mid_y = (total_y_2024 + total_y_2025) / 2
plt.text(mid_x, mid_y + 60, f"Increase = {percent_increase:.2f}%", ha='center', fontsize=10)

plt.xticks(x, quantification_df["Emission Category"])
plt.ylabel("Emissions (tCO2e)")
plt.xlabel("Emission Category")
plt.title("Quantification of Scope 1 and Scope 2 Emissions (2024 vs 2025)")
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

#pie chart...........................................

# Data
labels = [
    "Diesel Generators",
    "Refrigerant Leakage",
    "Fire Extinguishers",
    "Company Vehicles",
    "Electricity"
]

values_2024 = [40.079789, 183.204100, 0.099000, 43.293438, 1237.044020]
values_2025 = [18.437006, 217.788562, 0.086955, 45.974425, 1258.168052]

# Plot
fig, axes = plt.subplots(1, 2, figsize=(12, 6))

# 2024 Pie
axes[0].pie(
    values_2024,
    labels=labels,
    autopct='%1.1f%%',
    startangle=140
)
axes[0].set_title("Carbon Footprint Distribution - 2024")

# 2025 Pie
axes[1].pie(
    values_2025,
    labels=labels,
    autopct='%1.1f%%',
    startangle=140
)
axes[1].set_title("Carbon Footprint Distribution - 2025")

plt.tight_layout()
plt.show()

# ==========================================
# 51. 2025 EMISSION VALUES
# ==========================================
sources_2025 = {
    "Onsite Diesel Generators": 18.437006,
    "Refrigerant Leakage": 217.788562,
    "Fire Extinguishers": 0.086955,
    "Company Vehicles": 45.974425,
    "Electricity": 1258.168052
}

# ==========================================
# 52. CREATE DATAFRAME
# ==========================================
strategy_df = pd.DataFrame({
    "Emission Source": list(sources_2025.keys()),
    "Emission_tCO2e": list(sources_2025.values())
})

# total emissions
total_emissions = strategy_df["Emission_tCO2e"].sum()

# percentage contribution
strategy_df["Contribution_%"] = (strategy_df["Emission_tCO2e"] / total_emissions) * 100

# sort from highest to lowest
strategy_df = strategy_df.sort_values("Emission_tCO2e", ascending=False).reset_index(drop=True)

# priority rank
strategy_df["Priority_Rank"] = range(1, len(strategy_df) + 1)

# ==========================================
# 53. ADD PRACTICAL STRATEGY SUGGESTIONS
# ==========================================
strategy_actions = {
    "Electricity": "Improve branch energy efficiency, optimize lighting and AC usage",
    "Refrigerant Leakage": "Regular maintenance and leakage monitoring of cooling systems",
    "Company Vehicles": "Reduce unnecessary travel and improve fuel efficiency",
    "Onsite Diesel Generators": "Minimize generator dependence and improve backup efficiency",
    "Fire Extinguishers": "Maintain proper monitoring and periodic checks"
}

strategy_df["Suggested_Strategy"] = strategy_df["Emission Source"].map(strategy_actions)

print("\n=== Practical Sustainability Strategy Analysis ===")
print(strategy_df.round(2))

# save table
strategy_df.to_excel("Practical_Sustainability_Strategies.xlsx", index=False)

# ==========================================
# 54. GRAPH – STRATEGY PRIORITY BY EMISSION SOURCE
# ==========================================
plt.figure(figsize=(10, 6))
bars = plt.bar(strategy_df["Emission Source"], strategy_df["Contribution_%"])

# add percentage labels
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, height + 1, f"{height:.1f}%", ha="center")

plt.title("Priority Areas for Practical Sustainability Strategies")
plt.xlabel("Emission Source")
plt.ylabel("Contribution to Total Emissions (%)")
plt.xticks(rotation=20)
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()