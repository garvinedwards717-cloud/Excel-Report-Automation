

from pathlib import Path
import pandas as pd


def main() -> None:
    base_dir = Path(__file__).resolve().parents[1]
    data_dir = base_dir / "data"
    data_dir.mkdir(exist_ok=True)

    data = {
        "Date": [
            "2025-01-01",
            "2025-01-02",
            "2025-01-03",
            "2025-01-04",
            "2025-01-05",
            "2025-01-06",
        ],
        "Product": [
            "Laptop",
            "Mouse",
            "Keyboard",
            "Laptop",
            "Monitor",
            "Mouse",
        ],
        "Sales": [1200, 40, 80, 1300, 300, 50],
    }

    df = pd.DataFrame(data)
    output_file = data_dir / "sales_data.xlsx"
    df.to_excel(output_file, index=False)

    print(f"[OK] Excel file created successfully: {output_file}")


if __name__ == "__main__":
    main()