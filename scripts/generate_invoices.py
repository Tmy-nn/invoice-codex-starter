"""Local CLI; creates PDFs and .eml drafts only. Never sends mail."""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from src.invoices import InvoiceError, InvoiceRunError, generate  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Create sample invoice PDFs and unsent .eml drafts")
    parser.add_argument("--input", required=True, type=Path, help="customer workbook (.xlsx)")
    parser.add_argument("--month", required=True, help="billing month YYYY-MM")
    parser.add_argument("--expected-count", required=True, type=int, help="stop if workbook count differs")
    parser.add_argument("--output", required=True, type=Path, help="local output folder")
    args = parser.parse_args()
    try:
        run_dir = generate(args.input, args.month, args.expected_count, args.output)
    except InvoiceError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except InvoiceRunError as exc:
        print(f"ERROR: batch incomplete at {exc.run_dir}. Review INCOMPLETE.txt; do not use partial files.", file=sys.stderr)
        return 3
    except OSError:
        print("ERROR: input or output path could not be accessed. No completion is confirmed.", file=sys.stderr)
        return 3
    print(f"Created {args.expected_count} PDF and .eml pairs plus manifest: {run_dir}")
    print("No emails were sent. Review every recipient, amount and attachment before manual use.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
