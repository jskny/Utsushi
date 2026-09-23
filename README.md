# Utsushi

自社の帳票(Excel)を、レイアウトを崩さずPDF化するための専用変換ライブラリ。

見積書・請求書・納品書など、あらかじめ登録された自社の既知の帳票テンプレートを対象に、特定セルの文字列を動的な値に置換した上でPDF化する。**汎用的なExcel→PDF変換は目指さない**(任意のExcelファイルを扱う機能は対象外)。詳細な目的・スコープは [`.kiro/steering/product.md`](.kiro/steering/product.md) を参照。

## 使い方

呼び出し元プロダクトからの利用方法(API・例外の扱い)は [`docs/ライブラリの使い方.md`](docs/ライブラリの使い方.md) を参照。NuGetパッケージとしての配布は行っておらず、ソースツリー内から `src/Utsushi/Utsushi.csproj` への `ProjectReference` を前提とする。

帳票定義(`definition.json`)の書き方は [`docs/帳票定義スキーマ.md`](docs/帳票定義スキーマ.md) を参照。

## 開発

このリポジトリはAWS Kiroスタイルのspec駆動開発に従う。開発時の方針・制約・サブエージェントの使い方は [`CLAUDE.md`](CLAUDE.md) にまとめている(このリポジトリで作業するAIエージェント向けだが、人間の開発者にも同様に適用される)。

```bash
dotnet build
dotnet test
dotnet format
```

- `src/` — Core / Parsing / ReportDefinition / Substitution / Layout / Rendering の5レイヤー+共通基盤、ファサード `Utsushi`、CLI
- `tests/` — 各レイヤーのユニットテストとゴールデン(回帰)テスト
- `samples/reports/` — 帳票定義とテンプレートの合成サンプル(実運用の帳票テンプレートは社内データのため含まれない)
- `docs/` — スキーマ・利用方法・開発環境固有の注意点・過去の失敗事例集

対象開発環境はVisual Studio 2019(.NET 5、C# 9.0)。詳細は [`.kiro/steering/tech.md`](.kiro/steering/tech.md) を参照。

## ライセンス

[MIT License](LICENSE)。使用しているOSSライブラリのライセンス方針は [`.kiro/steering/tech.md`](.kiro/steering/tech.md) を参照(商用ライブラリ・Office Interopは不使用)。
