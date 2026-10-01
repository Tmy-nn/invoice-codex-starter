---
name: invoice-generate
description: 架空データまたは承認済みの請求データから請求書PDFと送信しないメール下書きファイルを作るときに使う。
---

# 請求書生成

1. `README.md` と `docs/invoice-workflow.md` を読む。
2. 入力ファイル、請求月、期待件数、出力先を確認する。情報が足りない場合は生成前に止める。
3. サンプル検証では `python scripts/make_sample_data.py --output data/sample_customers.xlsx --count 10` を使う。
4. `python scripts/generate_invoices.py --input data/sample_customers.xlsx --month 2026-10 --expected-count 10 --output output` の形式で実行する。請求月と件数は依頼に合わせて変える。
5. `COMPLETE.txt` を確認し、`summary.json` の件数・合計額・PDFとEMLの数、`manifest.csv` の各行を担当者に示す。`INCOMPLETE.txt` が残る結果は使わない。
6. EMLはローカルファイルであり、メールソフトの下書きフォルダに自動登録されたことを意味しない。送信は行わない。
7. 実データのGit追加・外部共有は行わない。
