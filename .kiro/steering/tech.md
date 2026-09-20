---
inclusion: always
---

# Utsushi 技術方針

## 実行環境

- **言語/ランタイム**: C# / .NET 5
- **対象OS**: サーバーサイドでの実行を想定(Windows依存・COM依存を作らない。将来的なLinuxコンテナでの実行を妨げない設計とする)

> **.NET 5固定の理由**: 本ライブラリを利用する予定の呼び出し元プロダクトが .NET 5 上で動作しており、そのプロセスに読み込まれる(同一プロセス内でアセンブリとして参照される)ことを前提とするため、`net5.0` をターゲットフレームワークとして固定する。呼び出し元の .NET バージョンが上がらない限り、Utsushi側だけを新しいTFMに上げることはできない。

> **Visual Studio 2019 対応**: 呼び出し元プロダクトの開発環境が Visual Studio 2019 であるため、Utsushi側も VS2019 でビルドできる必要がある。VS2019(最終版16.11)がバンドルするC#コンパイラは **C# 9.0 までしかサポートしない**(C# 10はVS2022以降が必要。これは一般に知られている事実であり、この開発環境〔Claude Code on the web〕にはVisual Studio自体が無いため実機検証はできていない)。TFMが`net5.0`であることとC#の言語バージョンは別軸のため、`net5.0`自体はVS2019で問題なくビルドできても、`LangVersion`を10以上にするとVS2019ではコンパイルエラーになる。このためLangVersionは`9.0`に固定し、file-scoped namespaceやrecord struct等C# 10以降の構文は使用しない。加えてVS2019は新しいXML形式のソリューションファイル(`.slnx`)を認識できないため、`vs2019/Utsushi.sln`(classic形式)を`Utsushi.slnx`と並行して維持する(ルート直下に置くと`.slnx`と衝突し`dotnet build`等の引数なし実行が壊れるため`vs2019/`配下に置く。詳細は`structure.md`「ソリューション構成」参照)。

> **注記(既知のリスク)**: .NET 5 は Microsoft のサポートが終了(EOL)しており、セキュリティパッチは提供されない。この点は呼び出し元プロダクトの制約に起因する既知のリスクとして許容し、呼び出し元が .NET 8 以降へ移行した際にはUtsushi側のTFMも追随できるよう、特定バージョンのランタイムAPIに過度に依存しない実装を心掛ける。

> **開発環境での .NET 5 SDKの扱い(重要)**: Claude Code on the webの実行環境(Ubuntu 24.04)のaptリポジトリには `dotnet-sdk-5.0` パッケージ自体が存在しない(`dotnet-sdk-8.0` / `dotnet-sdk-10.0` のみ)。ただし、**.NET 5 SDK/ランタイムを別途インストールする必要はない**。新しいSDK(検証時は `dotnet-sdk-10.0`)だけで `net5.0` をターゲットにしたビルド・テスト実行が問題なく行えることを実機検証済み。手順・根拠は `docs/開発環境メモ.md` の「.NET SDK」節を参照。要点のみ以下に記す。
>
> - `TargetFramework` を `net5.0` にしたクラスライブラリ/テストプロジェクトは、新しいSDKでも `net5.0` 用の参照アセンブリがNuGetから自動取得されビルドが通る(`dotnet build` 実測確認済み)。
> - `dotnet test` 等で実際に **実行** するプロジェクト(テストプロジェクト、CLI/APIのエントリーポイント等。参照専用のクラスライブラリには不要)には `<RollForward>LatestMajor</RollForward>` を追加すること。これによりインストール済みの新しいランタイム(例: .NET 10)にロールフォワードして実行できる(`dotnet test` が実際にグリーンになることを確認済み)。
> - 既定の `LangVersion` は `net5.0` の場合 C# 9 相当。Utsushiは上記「Visual Studio 2019 対応」の理由により `<LangVersion>9.0</LangVersion>` を明示しており、file-scoped namespace・record struct等C# 10以降の構文は使わない(この制約はVS2019固有のもので、この開発環境〔Claude Code on the web〕自体には起因しない)。

## ライセンス制約

- **商用ライブラリ禁止**: 有償ライセンス、または商用利用時に課金が発生するライセンス(例: Polyform Noncommercial、Community版に収益上限があるものなど)は採用しない。
- **Office Interop禁止**: `Microsoft.Office.Interop.Excel` 等のCOM相互運用機構は使用しない。
- 採用する依存ライブラリは **MIT / Apache-2.0 / BSD 等の永続的な無償OSSライセンス**に限定する。
- 新規に依存ライブラリを追加する際は、ライセンス種別を確認し `security-reviewer` サブエージェントでのレビュー対象とする。

## 採用技術(候補と選定理由)

対象帳票の解析・レイアウト計算・PDF描画をそれぞれ別レイヤーの責務とし、各レイヤーで以下の採用を基本方針とする(詳細は `.kiro/specs/excel-report-pdf-conversion/design.md` を参照)。

| レイヤー | 候補ライブラリ | ライセンス | 採用理由 |
|---|---|---|---|
| Excel構造解析 | `DocumentFormat.OpenXml`(Open XML SDK) | MIT | Microsoft公式のOOXML SDK。印刷範囲・改ページ・列幅/行高・結合セル・スタイル・共有文字列など、レイアウト再現に必要な構造情報を欠落なく取得できる。読み取り専用用途であればInteropより高速・安定。 |
| (補助) Excel構造解析 | `ClosedXML` / `NPOI` | MIT / Apache-2.0 | OpenXml SDKは低レベルAPIのため、値の取得補助や動作検証用の代替実装として利用を検討可(必須ではない)。 |
| 描画・PDF出力 | `SkiaSharp` | MIT | クロスプラットフォームな2D描画ライブラリ。フォントメトリクスに基づく正確なテキスト配置と、`SKDocument` によるPDF直接出力を1ライブラリで完結できるため、自作レイアウトエンジンの「計算した座標に正確に描画する」という要件に最も適合する。 |

採用済みのバージョン(net5.0 互換を確認済み):

| パッケージ | バージョン | 備考 |
|---|---|---|
| `DocumentFormat.OpenXml` | 2.20.0 | netstandard2.0 対応。3.x は net6.0 以降が必要 |
| `SkiaSharp` | 2.88.8 | netstandard2.0 対応 |
| `SkiaSharp.NativeAssets.Linux` | 2.88.8 | Linuxコンテナでの実行に必要 |

> **既知の制約(SkiaSharp / PDFのフォント)**: NuGetで配布される SkiaSharp のネイティブビルドは、
> PDF出力時に**フォントのサブセット化を行わず、使用フォントを丸ごと埋め込む**(2.88系・3.x系とも実測確認)。
> 日本語フォントは数MBあるため、1ページの帳票でも出力PDFが4MB前後になる。
> **PDF内の文字列検索を維持するため、フォント埋め込みを既定とし、このサイズは許容する方針**。
> サイズを優先する場合は `PdfTextRendering.Outline`(文字のアウトライン化。同帳票で91KB、
> 見た目は同一だがPDF内検索は不可)を選べる。
>
> **太字/斜体の注意**: SkiaSharp に書体レベルで太字/斜体を合成させると、PDFが Type 3 フォントになり
> 文字列検索ができなくなる。Utsushi は通常字形を埋め込んで描画時に装飾を合成することで回避している。
> 実字形の太字フォントファイルがある場合は `FontResolverOptions.FontFiles` に
> `ファミリ名:bold` のキーで登録すれば、合成せずにそれを使う。
>
> 詳細は `.kiro/specs/excel-report-pdf-conversion/design.md`「Rendering レイヤー」を参照。

商用配布物(Interop、EPPlus 5以降の商用ライセンス、Aspose、Spire.PDF等)は候補から除外する。

## 開発コマンド

以下を標準コマンドとする(ソリューション構成は `structure.md` を参照)。いずれもリポジトリのルートで実行する。

```bash
dotnet build
dotnet test
dotnet format        # コードスタイル整形
```

## コーディング規約

- 各プロジェクト(`.csproj`)共通で以下を設定する。
  - `<TargetFramework>net5.0</TargetFramework>`
  - `<Nullable>enable</Nullable>`
  - `<LangVersion>9.0</LangVersion>`(呼び出し元プロダクトの開発環境がVisual Studio 2019であるため。上記「Visual Studio 2019 対応」を参照)
  - テストプロジェクト・実行可能プロジェクト(参照専用のクラスライブラリを除く)には `<RollForward>LatestMajor</RollForward>` を追加する(この開発環境に.NET 5ランタイムが無くても実行できるようにするため。詳細は上記「開発環境での .NET 5 SDKの扱い」を参照)。
- 名前空間は従来のブロック形式(`namespace X { ... }`)を使用する。file-scoped namespace(`namespace X;`)はC# 10以降の構文でありVS2019では使えないため使用しない。
- `record` / `record class`(positional record・`with`式含む)はC# 9の機能でありVS2019でも使用可。ただし **`record struct` はC# 10の機能のため使用しない**。値の等価性を持つ構造体が必要な場合は `readonly struct` + `IEquatable<T>` を手書きする(`src/Utsushi.Core/ArgbColor.cs` を参照)。
- レイヤー間の依存方向を一方向に保つ(下記 `structure.md` の依存ルールを参照)。境界を越えた参照(例: PDF描画レイヤーが直接OpenXmlの型を参照する等)を作らない。
- 帳票ごとの個別分岐(`if (帳票名 == "請求書") { ... }` のようなコード)は「帳票定義データ」側に寄せ、レイアウト計算・描画ロジックには帳票固有のハードコードを持ち込まない。
- 座標・寸法計算はミリ単位/ポイント単位の混在を避け、内部表現の単位を design.md に定義し統一する。

## テスト戦略

- ユニットテスト: 改ページ計算、列幅/行高からのピクセル(ポイント)換算、文字列置換ロジックなど、純粋なロジックを対象にする。
- ゴールデンファイル(回帰)テスト: 登録済み帳票サンプルを変換し、レビュー済みの期待結果と比較する。
  比較対象はPDFバイナリではなく `PagedLayout`(ページごとの描画命令)のスナップショットとする。
  PDFバイナリは生成日時・圧縮・ライブラリのバージョンで変わり、回帰検出に使えないため。
  詳細は `.kiro/specs/excel-report-pdf-conversion/design.md`「テスト戦略」を参照。
- テストフレームワークは xUnit を基本とする。パッケージバージョンは net5.0 世代のものに合わせる(`Microsoft.NET.Test.Sdk` 17.1.0 / `xunit` 2.4.1 / `xunit.runner.visualstudio` 2.4.3 の組み合わせで `dotnet test` が正常に動作することを確認済み。新しすぎる組み合わせ〔`Microsoft.NET.Test.Sdk` 17.11.1 等〕はnet5.0でテストが発見されない事象を確認しているため避ける)。

## 依存ライブラリ追加時のルール

1. ライセンスが「無償・商用利用可」であることを確認する。
2. `security-reviewer` サブエージェントで既知の脆弱性・メンテナンス状況を確認する。
3. `.kiro/steering/tech.md` の対象表を更新する。
