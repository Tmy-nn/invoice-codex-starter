"""Generate fully synthetic customer rows for a local demonstration."""

import argparse
from pathlib import Path

from openpyxl import Workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a synthetic invoice workbook")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--count", type=int, default=10)
    args = parser.parse_args()
    if not 1 <= args.count <= 1000:
        parser.error("count must be 1 to 1000")
    if args.output.exists():
        parser.error(f"file already exists (no overwrite): {args.output}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    sheet = book.active
    sheet.title = "invoices"
    sheet.append(("customer_id", "customer_name", "email", "description", "amount_yen"))
    for number in range(1, args.count + 1):
        sheet.append((f"S{number:04d}", f"架空顧客{number:04d}", f"customer{number:04d}@example.invalid", "サンプル業務委託料", 10000 + number * 100))
    book.save(args.output)
    print(f"Created {args.count} synthetic rows: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
