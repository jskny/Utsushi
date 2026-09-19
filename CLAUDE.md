# Utsushi

自社の帳票(Excel)を、レイアウトを崩さずPDF化するための専用変換プロダクト。汎用的なExcel→PDF変換は目指さない。
詳細は `.kiro/steering/product.md` を参照。

## 現在の状態

コードベースは未着手。現時点で整備済みなのは以下のみ:

- `.kiro/steering/` — 常時適用される方針(製品概要・技術方針・プロジェクト構成)
- `.kiro/specs/excel-report-pdf-conversion/` — 中核機能の要件定義書・設計書・実装タスクリスト
- `.claude/agents/` — レビュー・テスト作成用サブエージェント

実装に着手する際は、`.kiro/specs/excel-report-pdf-conversion/tasks.md` のタスクを順に消化すること。

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
| `code-reviewer` | 実装変更が完了した直後、コミット前 |
| `test-writer` | 新規/変更実装にテストが不足している場合 |
| `security-reviewer` | Excelファイル入力・帳票定義・出力パスを扱う変更、新規依存ライブラリ追加時 |
| `spec-compliance-reviewer` | requirements/design/tasksの作成・更新時、実装完了後の仕様整合性確認 |

## 開発コマンド(プロジェクト作成後)

```bash
dotnet build
dotnet test
dotnet format
```
