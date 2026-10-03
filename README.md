# Utsushi

自社の帳票(Excel)を、レイアウトを崩さずPDF化するための専用変換ライブラリ。

見積書・請求書・納品書など、あらかじめ登録された自社の既知の帳票テンプレートを対象に、特定セルの文字列を動的な値に置換した上でPDF化する。**汎用的なExcel→PDF変換は目指さない**。補助的に、社内で作ったExcelファイルを帳票定義なしでそのままPDF化することもできる(置換キーは使えずセル番地での書き換えのみ・見た目の一致は保証しないベストエフォート)。詳細な目的・スコープは [`.kiro/steering/product.md`](.kiro/steering/product.md) を参照。

## 使い方

```csharp
using var pdf = new Utsushi.Excel2Pdf("input.xlsx");
pdf.SetText("C2", "Hello World");
pdf.SetValue("D5", 123.5);   // セルの表示形式(通貨・日付など)で表示される
pdf.Save("output.pdf");
```

上の例は帳票定義なしの変換(ベストエフォート。未対応の要素は黙って出力されない)。登録済みの帳票は
`new Utsushi.Excel2Pdf("template.xlsx", "invoice", "samples/reports")` のように帳票コードと帳票定義のルートを渡して使う。

呼び出し元プロダクトからの利用方法(API・例外の扱い)は [`docs/ライブラリの使い方.md`](docs/ライブラリの使い方.md) を参照。NuGetパッケージとしての配布は行っておらず、ソースツリー内から `src/Utsushi/Utsushi.csproj` への `ProjectReference` を前提とする。

帳票定義(`definition.json`)の書き方は [`docs/帳票定義スキーマ.md`](docs/帳票定義スキーマ.md) を参照。

Excelでテンプレートを作る担当者向けの注意点(フォント、住所など差し込みセルの設定、再現できるExcelの機能・できない機能)は
[`docs/テンプレート作成ガイド.md`](docs/テンプレート作成ガイド.md) を参照。

### 主な特長

- Excelの印刷設定(印刷範囲・改ページ・印刷タイトル・余白・拡大縮小・ヘッダー/フッター)、結合セル・罫線・画像・図形を再現する
- エラーにならないまま差し込んだ値が欠けたり化けたりすることを防ぐ(制御文字・空の必須値・行の高さに収まらない住所・印刷範囲外のセル・字形の無い文字はエラーとして知らせる)
- 日本語フォント BIZ UDPゴシックを同梱する。外字の補完や代替フォントの許容時に使われ、日本語フォントの無いサーバーでも日本語を描画できる(テンプレートで指定したフォントそのものは、既定ではサーバーに必要)
- 人名・住所の外字・異体字に対応する(IPAmj明朝をインストールすれば「𠮷」などJIS第4水準外の文字も描ける)
- PDFには使った字形だけを埋め込むため小さい(請求書サンプルで約54KB)。PDF内の文字列検索・コピーもできる
- Office・Excel・商用ライブラリは不要(.NET 5 / SkiaSharp / Open XML SDK)

### 本番環境に必要なもの

- テンプレートで使うフォント(ＭＳ Ｐゴシック等)を、PDFを作るサーバーにインストールする
- JIS第4水準外の人名用漢字を扱う場合は IPAmj明朝 をインストールする

詳細は [`docs/ライブラリの使い方.md`](docs/ライブラリの使い方.md)「本番環境に導入する前に」を参照。

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
- `docs/` — テンプレート作成ガイド・スキーマ・利用方法・開発環境固有の注意点・過去の失敗事例集

呼び出し元プロダクトの開発環境がVisual Studio 2019(.NET 5、C# 9.0まで)であるため、その制約に合わせている(このプロジェクト自体の開発にVisual Studioは必須ではない)。
Visual Studio 2019ではルートの `Utsushi.sln` を開く。.NET 6以降のSDKも入っているPCでは、`global.json` で .NET 5 SDKに固定する必要がある(手順は [`docs/開発環境メモ.md`](docs/開発環境メモ.md)「1. .NET SDK」の「Visual Studio 2019 でのビルド」)。
詳細は [`.kiro/steering/tech.md`](.kiro/steering/tech.md) を参照。

## ライセンス

[MIT License](LICENSE)。

同梱の日本語フォント BIZ UDPゴシック(`src/Utsushi.Rendering/Fonts/BIZUDPGothic-Regular.ttf`)は
[SIL Open Font License 1.1](src/Utsushi.Rendering/Fonts/BIZUDPGothic-OFL.txt) に従う(MIT License の対象外)。

使用しているOSSライブラリのライセンス方針は [`.kiro/steering/tech.md`](.kiro/steering/tech.md) を参照(商用ライブラリ・Office Interopは不使用)。
