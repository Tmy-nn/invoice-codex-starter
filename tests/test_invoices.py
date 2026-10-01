import csv
import json
import unittest
import uuid
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from src.invoices import InvoiceError, InvoiceRunError, generate, read_rows


def workbook(path: Path, records: list[tuple]) -> None:
    book = Workbook()
    sheet = book.active
    sheet.append(("customer_id", "customer_name", "email", "description", "amount_yen"))
    for record in records:
        sheet.append(record)
    book.save(path)


class InvoiceTests(unittest.TestCase):
    def setUp(self):
        # Keep test artifacts in ignored local output for inspection; never remove files.
        self.root = Path("output") / "test-runs" / uuid.uuid4().hex
        self.root.mkdir(parents=True, exist_ok=False)
        self.input = self.root / "input.xlsx"
        self.records = [
            ("A001", "架空顧客一", "a@example.invalid", "サンプル作業", 11000),
            ("A002", "架空顧客二", "b@example.invalid", "サンプル作業", 22000),
        ]

    def test_two_outputs_and_attachment(self):
        workbook(self.input, self.records)
        run = generate(self.input, "2026-10", 2, self.root / "output")
        self.assertEqual(len(list(run.glob("*.pdf"))), 2)
        self.assertEqual(len(list(run.glob("*.eml"))), 2)
        with (run / "manifest.csv").open(encoding="utf-8-sig", newline="") as handle:
            records = list(csv.DictReader(handle))
        self.assertEqual([r["amount_yen"] for r in records], ["11000", "22000"])
        summary = json.loads((run / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["expected_count"], 2)
        self.assertEqual(summary["pdf_count"], 2)
        self.assertEqual(summary["eml_count"], 2)
        self.assertEqual(summary["total_amount_yen"], 33000)
        self.assertTrue((run / "COMPLETE.txt").exists())
        self.assertFalse((run / "INCOMPLETE.txt").exists())
        message = BytesParser(policy=policy.default).parsebytes((run / "draft_2026-10_A001.eml").read_bytes())
        self.assertEqual(message["To"], "a@example.invalid")
        attachments = list(message.iter_attachments())
        self.assertEqual(len(attachments), 1)
        self.assertTrue(attachments[0].get_payload(decode=True).startswith(b"%PDF"))
        self.assertIsNone(message["From"])
        self.assertIsNone(message["Date"])

    def test_validation_prevents_any_output(self):
        invalid = self.records + [("A001", "重複", "bad address", "", -1)]
        workbook(self.input, invalid)
        with self.assertRaisesRegex(InvoiceError, "duplicate customer_id"):
            generate(self.input, "2026-10", 3, self.root / "output")
        self.assertFalse((self.root / "output").exists())

    def test_expected_count_prevents_output(self):
        workbook(self.input, self.records)
        with self.assertRaisesRegex(InvoiceError, "count mismatch"):
            generate(self.input, "2026-10", 200, self.root / "output")
        self.assertFalse((self.root / "output").exists())

    def test_new_run_does_not_overwrite_prior_run(self):
        workbook(self.input, self.records)
        first = generate(self.input, "2026-10", 2, self.root / "output")
        second = generate(self.input, "2026-10", 2, self.root / "output")
        self.assertNotEqual(first, second)
        self.assertTrue((first / "manifest.csv").exists())

    def test_non_integer_amount_rejected(self):
        workbook(self.input, [("A001", "架空", "a@example.invalid", "業務", 100.5)])
        with self.assertRaisesRegex(InvoiceError, "positive integer"):
            read_rows(self.input)

    def test_formula_like_name_rejected_before_manifest(self):
        workbook(self.input, [("A001", "=1+1", "a@example.invalid", "業務", 100)])
        with self.assertRaisesRegex(InvoiceError, "invalid customer_name"):
            generate(self.input, "2026-10", 1, self.root / "output")
        self.assertFalse((self.root / "output").exists())

    def test_partial_failure_retains_incomplete_marker(self):
        workbook(self.input, self.records)
        with patch("src.invoices.render_pdf", side_effect=OSError("disk failure")):
            with self.assertRaises(InvoiceRunError) as caught:
                generate(self.input, "2026-10", 2, self.root / "output")
        run = caught.exception.run_dir
        self.assertTrue((run / "INCOMPLETE.txt").exists())
        self.assertFalse((run / "COMPLETE.txt").exists())
        self.assertFalse((run / "summary.json").exists())

    def test_pdf_width_rejected_before_output(self):
        workbook(self.input, [("A001", "請" * 60, "a@example.invalid", "業務", 100)])
        with self.assertRaisesRegex(InvoiceError, "does not fit PDF"):
            generate(self.input, "2026-10", 1, self.root / "output")
        self.assertFalse((self.root / "output").exists())


if __name__ == "__main__":
    unittest.main()
