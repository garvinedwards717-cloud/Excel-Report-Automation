# main.py
import pandas as pd
import matplotlib.pyplot as plt
import os

# --------------------------
# File paths
# --------------------------
DATA_FILE = "data/sales_data.xlsx"        # Your input Excel file
OUTPUT_FILE = "output/cleaned_sales_data.xlsx"  # Where cleaned report will go
CHART_FILE = "output/sales_chart.png"     # Where chart will be saved

# --------------------------
# 1️⃣ Read Excel
# --------------------------
try:
    df = pd.read_excel(DATA_FILE, engine="openpyxl")
    print("✅ Raw data loaded successfully from Excel:")
    print(df)  # <-- Print raw data here
except FileNotFoundError:
    print(f"❌ File not found: {DATA_FILE}")
    exit()
except Exception as e:
    print(f"❌ Error loading Excel file: {e}")
    exit()

# --------------------------
# 2️⃣ Clean / Aggregate data
# --------------------------
# Example: total sales per product
df_clean = df.groupby("Product", as_index=False)["Sales"].sum()
print("\n✅ Aggregated / cleaned data:")
print(df_clean)

# --------------------------
# 3️⃣ Save cleaned report
# --------------------------
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
df_clean.to_excel(OUTPUT_FILE, index=False)
print(f"\n✅ Cleaned report saved to: {OUTPUT_FILE}")

# --------------------------
# 4️⃣ Create chart
# --------------------------
plt.figure(figsize=(8, 5))
plt.bar(df_clean["Product"], df_clean["Sales"], color="skyblue")
plt.xlabel("Product")
plt.ylabel("Total Sales")
plt.title("Total Sales per Product")
plt.tight_layout()

os.makedirs(os.path.dirname(CHART_FILE), exist_ok=True)
plt.savefig(CHART_FILE)
plt.close()
print(f"✅ Chart saved to: {CHART_FILE}")