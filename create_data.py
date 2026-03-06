import pandas as pd
import os

# Ensure data folder exists
os.makedirs("data", exist_ok=True)

data = {
    "Date": ["2025-01-01","2025-01-02","2025-01-03","2025-01-04","2025-01-05","2025-01-06"],
    "Product": ["Laptop","Mouse","Keyboard","Laptop","Monitor","Mouse"],
    "Sales": [1200,40,80,1300,300,50]
}

df = pd.DataFrame(data)

df.to_excel("data/sales_data.xlsx", index=False)

print("Excel file created successfully!")