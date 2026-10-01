# 企業向け Codex 請求書スターター

Excelの架空顧客一覧から、請求書PDFと**未送信のメール下書きファイル（`.eml`）**を一括生成する独立した作業環境です。各行は `manifest.csv`、件数と合計額は `summary.json` で照合できます。Codexはリポジトリ直下の `AGENTS.md` とローカルの請求書スキルを参照します。

**現在の版は機能検証用です。** 発行者、登録番号、税区分、振込先、実際の請求様式は未設定です。PDFには `SAMPLE / 架空データ` を明記しています。実際の請求やメール送信には使用しないでください。

## 架空データで試す

Python 3.11以上を用意し、このリポジトリのフォルダで実行します。Windowsでは `python` の代わりに `py`、macOSでは `python3` が必要な環境もあります。

```bash
python -m venv .venv
```

仮想環境を有効にします。

| OS | コマンド |
|---|---|
| Windows PowerShell | `.venv\Scripts\Activate.ps1` |
| macOS / Linux | `source .venv/bin/activate` |

```bash
python -m pip install -r requirements.txt
python scripts/make_sample_data.py --output data/sample_customers.xlsx --count 10
python scripts/generate_invoices.py --input data/sample_customers.xlsx --month 2026-10 --expected-count 10 --output output
```

出力先は `output/2026-10/run-.../` です。1回の実行でPDF 10件、EML 10件、`manifest.csv`、件数・合計額の `summary.json`、`COMPLETE.txt` を作ります。途中で失敗すると `INCOMPLETE.txt` が残り、そのフォルダのPDF・EMLは使用しません。再実行では新しい `run-...` フォルダを作り、既存結果を上書きしません。メールソフトへの下書き登録や送信は行いません。

200件の処理量を試すときは、別のサンプルExcelを作ります。

```bash
python scripts/make_sample_data.py --output data/sample_200.xlsx --count 200
python scripts/generate_invoices.py --input data/sample_200.xlsx --month 2026-10 --expected-count 200 --output output
```

生成スクリプトは既存の入力Excelを上書きしません。入力列は `customer_id`、`customer_name`、`email`、`description`、`amount_yen` の5列です。**`amount_yen` は税込の確定金額としてそのまま表示**し、税額の計算や端数処理はしません。列の詳細と担当者の確認手順は [請求書ワークフロー](docs/invoice-workflow.md) にあります。

## Codexで使う

Codexでこのフォルダを開き、「架空データ10件から2026-10の請求書を作って」と依頼できます。操作ルールは [AGENTS.md](AGENTS.md)、再利用する手順は [.agents/skills/invoice-generate/SKILL.md](.agents/skills/invoice-generate/SKILL.md) にあります。Codexがコマンドを提案・実行しても、結果の件数、宛先、金額とPDFは担当者が確認してください。

## GitHubに入れるもの

このリポジトリに入れるのはコード、説明、テスト、再配布可能な日本語フォントだけです。`data/`、`output/`、Excel、PDF、EML、認証情報は `.gitignore` で除外します。**非公開リポジトリであっても、取引先や従業員の実データをGitへ追加しない**運用です。Codexをローカルで使う場合も、推論のため入力が外部に送られ得るため、実データ利用前に自社の契約とデータ取扱条件を確認してください。

## 動作確認と導入順

```bash
python -m unittest discover -s tests -v
```

この公開版で動くのは架空データの請求書フローのみです。ほかの業務への展開例は [拡張例](docs/roadmap.md) に整理しています。本番化の差し替え項目とリポジトリの引き渡し方は [導入・引き渡し手順](docs/handoff.md) にあります。

## 由来

実行コードはWindowsのExcel COMに依存しない形で作成しました。リポジトリ構成は [OpenAI CookbookのCodexワークフロー例](https://github.com/openai/openai-cookbook/blob/main/examples/codex/iterating-development-workflows-with-codex.md) を参考にしています。公開リポジトリのコードを複写していません。
