#!/usr/bin/env python3
import argparse
import csv
from pathlib import Path

from openpyxl import Workbook


def convert_csv_to_xlsx(csv_path: Path, xlsx_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active

    with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            worksheet.append(row)

    workbook.save(xlsx_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert CSV file to XLSX file.")
    parser.add_argument("csv_file", type=Path, help="Input CSV file path.")
    parser.add_argument("xlsx_file", type=Path, nargs="?", help="Output XLSX file path.")
    args = parser.parse_args()

    csv_path = args.csv_file.resolve()
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    xlsx_path = args.xlsx_file.resolve() if args.xlsx_file else csv_path.with_suffix(".xlsx")
    convert_csv_to_xlsx(csv_path, xlsx_path)
    print(f"Converted: {csv_path} -> {xlsx_path}")


if __name__ == "__main__":
    main()
