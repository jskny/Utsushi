# Utsushi

自社の帳票(Excel)を、レイアウトを崩さずPDF化するための専用変換プロダクト。汎用的なExcel→PDF変換は目指さない。
詳細は `.kiro/steering/product.md` を参照。

## 現在の状態

中核機能(`.kiro/specs/excel-report-pdf-conversion/`)の実装は完了している。`dotnet build` / `dotnet test` / `dotnet format` はいずれもグリーン。

- `src/` — Core / Parsing / ReportDefinition / Substitution / Layout / Rendering の5レイヤー+共通基盤、ファサード `Utsushi`、CLI
- `tests/` — 各レイヤーのユニットテストとゴールデン(回帰)テスト
- `samples/reports/<帳票コード>/` — 帳票定義(`definition.json`)とテンプレート(`template.xlsx`)
- `tools/Utsushi.SampleGenerator/` — サンプル帳票 `.xlsx` の生成ツール(製品コードからは参照されない)
- `.kiro/steering/` — 常時適用される方針(製品概要・技術方針・プロジェクト構成)
- `.kiro/specs/excel-report-pdf-conversion/` — 中核機能の要件定義書・設計書・実装タスクリスト
- `docs/帳票定義スキーマ.md` — `definition.json` のスキーマ
- `docs/ライブラリの使い方.md` — ファサード `Utsushi`(`ReportPdfConverter`)のAPI・例外の使い方
- `docs/開発環境メモ.md` — Claude Code on the web実行環境で裏取りした環境固有の注意点(SDKセットアップ、日本語フォント、`pkill -f`の自己マッチ問題など)
- `docs/実装設計失敗事例集.md` — 過去の不具合・設計ミス・検出漏れとその原因・対応・教訓の記録。実装・レビュー着手前に関連する節を確認する

> **サンプル帳票について**: 実運用の帳票テンプレートは社内データのためリポジトリに含められない。
> `samples/reports/` にあるのは、実帳票と同等の構造(結合セル・罫線・数値書式・印刷範囲・印刷タイトル・複数ページ)を持つ**合成サンプル**である。
> 実帳票を追加する際は `.kiro/steering/structure.md`「帳票の追加手順」に従う。

変更に着手する前に、該当レイヤーの既存実装を `code-investigator` で確認すること。実装中に環境起因と思われるエラーに遭遇したら、まず `docs/開発環境メモ.md` を確認する。

## 開発の進め方(spec駆動)

このプロジェクトはAWS Kiroスタイルのspec駆動開発に従う。

1. 新機能・大きな変更に着手する前に、`.kiro/specs/<機能名>/` に `requirements.md`(EARS形式の受け入れ基準)→ `design.md`(アーキテクチャ・データモデル)→ `tasks.md`(要件番号を参照した実装タスク)の順で仕様を作成・更新する。
2. 既存の中核機能については `.kiro/specs/excel-report-pdf-conversion/` を必ず参照する。ここに書かれた要件・設計から逸脱する実装をする場合は、先に仕様側を更新する。
3. `.kiro/steering/` の3ファイル(product.md / tech.md / structure.md)はプロジェクト全体に常時適用される方針であり、個別のspecより優先する。

## 必ず守る制約

- **商用ライブラリ禁止・Office Interop禁止**(`.kiro/steering/tech.md`)。新規依存ライブラリを追加する前にライセンスを確認する。
- 対象は登録済みの自社帳票のみ。任意のExcelファイルを汎用的に扱う機能は追加しない。
- レイヤー間の依存は一方向(`Parsing → ReportDefinition → Substitution → Layout → Rendering`)。上位レイヤーが下位レイヤーの実装詳細を直接参照しない。
- 帳票固有の分岐は帳票定義データ側に置き、共通レイヤーのコードにハードコードしない。

## サブエージェント

以下は `.claude/agents/` に定義済み。該当する状況でプロアクティブに使用する。

| サブエージェント | 使うタイミング |
|---|---|
| `code-investigator` | 実装に着手する前の既存コード調査、バグ調査時の関連コード特定、影響範囲の洗い出し。メインの会話コンテキストを消費したくない大量のファイル読み込みが必要な調査全般 |
| `code-reviewer` | 実装変更が完了した直後、コミット前 |
| `layout-fidelity-reviewer` | Layout/Renderingレイヤー(改ページ計算・結合セル・フォントメトリクス・単位換算など)への変更直後。本プロダクトで最もリスクが高い領域専門のレビュー |
| `test-writer` | 新規/変更実装にテストが不足している場合 |
| `security-reviewer` | Excelファイル入力・帳票定義・出力パスを扱う変更、新規依存ライブラリ追加時。差分に閉じず`src/Utsushi.Parsing/OpenXml/`全体の未対応ループも横断確認する(過去に上限漏れが後追いで複数回見つかった経緯があるため) |
| `spec-compliance-reviewer` | requirements/design/tasksの作成・更新時、実装完了後の仕様整合性確認(仕様↔実装のトレーサビリティ) |
| `doc-reviewer` | `.kiro/steering/`・`.kiro/specs/`・`docs/`・`CLAUDE.md`等の追加・更新直後。ドキュメント間の矛盾、技術的記載の裏取り、参照切れ、体裁を確認 |

## 開発コマンド

```bash
dotnet build
dotnet test
dotnet format

# 帳票サンプル(.xlsx)を再生成する
dotnet run --project tools/Utsushi.SampleGenerator -- samples/reports

# ゴールデンファイルの期待値を更新する(差分は必ずレビューしてからコミット)
UTSUSHI_UPDATE_GOLDEN=1 dotnet test

# CLIで変換を試す
dotnet run --project src/Utsushi.Cli -- \
  --report invoice --input samples/reports/invoice/template.xlsx \
  --output /tmp/invoice.pdf --definitions samples/reports \
  --set CustomerName="株式会社サンプル 御中" --set InvoiceNo="INV-0001" --set TotalAmount="¥1,000"
```

> 開発環境に対象フォント(MS PGothic等)が無い場合、既定の厳格モードでは `FontNotAvailableException` になる。
> 動作確認だけなら `--allow-font-fallback IPAGothic` のように代替フォントを指定する(見た目は崩れる)。
