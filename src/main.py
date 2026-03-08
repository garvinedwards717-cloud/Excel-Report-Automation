from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def print_section(title: str) -> None:
    """Print a formatted section header."""
    print("\n" + "=" * 60)
    print(title)
    print("=" * 60)


def main() -> None:
    """Run the Excel report automation workflow."""
    base_dir = Path(__file__).resolve().parents[1]
    data_file = base_dir / "data" / "sales_data.xlsx"
    output_dir = base_dir / "output"
    output_dir.mkdir(exist_ok=True)

    output_file = output_dir / "cleaned_sales_data.xlsx"
    chart_file = output_dir / "sales_chart.png"

    print_section("EXCEL REPORT AUTOMATION")

    try:
        df = pd.read_excel(data_file, engine="openpyxl")
        print("[OK] Raw data loaded from: data/sales_data.xlsx")
    except FileNotFoundError:
        print("[ERROR] Input file not found: data/sales_data.xlsx")
        print("[TIP] Run: python src/create_data.py")
        return
    except Exception as error:
        print(f"[ERROR] Failed to read Excel file: {error}")
        return

    try:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
        df["Product"] = df["Product"].astype(str).str.strip()

        df = df.dropna(subset=["Date", "Product", "Sales"])
        df = df.sort_values("Date").reset_index(drop=True)

        print_section("CLEANED SALES DATA PREVIEW")
        print(df.to_string(index=False))

        df.to_excel(output_file, index=False, engine="openpyxl")
        print("\n[OK] Cleaned Excel file saved: output/cleaned_sales_data.xlsx")
    except Exception as error:
        print(f"[ERROR] Failed during data cleaning or saving: {error}")
        return

    try:
        sales_summary = df.groupby("Product", as_index=False)["Sales"].sum()

        print_section("SALES SUMMARY")
        print(sales_summary.to_string(index=False))

        plt.figure(figsize=(8, 5))
        plt.bar(sales_summary["Product"], sales_summary["Sales"])
        plt.title("Sales by Product")
        plt.xlabel("Product")
        plt.ylabel("Total Sales")
        plt.tight_layout()
        plt.savefig(chart_file)
        plt.close()

        print("\n[OK] Sales chart saved: output/sales_chart.png")
    except Exception as error:
        print(f"[ERROR] Failed to create sales chart: {error}")
        return

    print_section("PROCESS COMPLETE")
    print("[OK] Excel automation completed successfully.")


if __name__ == "__main__":
    main()