---
inclusion: always
---

# Utsushi プロジェクト構成方針

> 中核機能(`.kiro/specs/excel-report-pdf-conversion/`)の実装は完了している。
> 以降の変更も本方針に従うこと。実際の構成は下記「ソリューション構成」を参照。

## レイヤー構成

自社帳票のPDF化パイプラインは以下の一方向の依存関係を持つレイヤーに分割する。

```
共通基盤 (Core)  ← 全レイヤーが参照してよい唯一の下位プロジェクト
        ↓
Excel解析レイヤー (Parsing)
        ↓
帳票定義レイヤー (ReportDefinition)
        ↓
値差し込み(置換)レイヤー (Substitution)
        ↓
レイアウト計算レイヤー (Layout)
        ↓
PDF描画レイヤー (Rendering)
        ↓
ファサード (Utsushi)  ← 呼び出し元プロダクトが参照する唯一のアセンブリ
```

`Utsushi.Core` は例外階層(`UtsushiException` とその派生)と、レイヤーをまたぐ値型
(`CellAddress` / `CellRange` / `ArgbColor` / `RectPt` / `PointPt` / `Units`)のみを持つ。
ロジックは置かず、どのレイヤーからも参照してよい。逆に `Utsushi.Core` は他のどのプロジェクトも参照しない。

- 上位レイヤー(Rendering側)は下位レイヤー(Parsing側)の実装詳細(例: OpenXml SDKの型)を直接知ってはならない。レイヤー間の受け渡しは各レイヤーが定義する内部モデル(POCO)を介する。
- 帳票ごとの差異(セル位置、置換対象キー、既知の特殊処理)は「帳票定義レイヤー」のデータ(JSON等)に閉じ込め、他レイヤーのコードに帳票固有の分岐を持ち込まない。

## ソリューション構成

```
Utsushi/
├── src/
│   ├── Utsushi.Core/               # 例外階層とレイヤー共通の値型のみ。他プロジェクトを参照しない
│   ├── Utsushi.Parsing/            # Excel(OOXML)構造の読み取り。DocumentFormat.OpenXml に依存してよい唯一のプロジェクト
│   ├── Utsushi.ReportDefinition/   # 帳票定義モデルとローダー(JSON定義ファイル → 内部モデル)
│   ├── Utsushi.Substitution/       # 指定セルへの文字列置換ロジック
│   ├── Utsushi.Layout/             # 改ページ・印刷範囲・列幅行高・フォントメトリクスに基づくレイアウト計算
│   ├── Utsushi.Rendering/          # SkiaSharp を用いたPDF描画。Layoutレイヤーの計算結果のみを入力とする
│   ├── Utsushi/                    # ファサード(ReportPdfConverter)。呼び出し元プロダクトはここだけを参照する
│   └── Utsushi.Cli/                # コマンドライン入口。ファサードを呼ぶだけの薄い層
├── tests/
│   ├── Utsushi.TestSupport/        # テストプロジェクト間で共有するヘルパー(製品コードからは参照されない)
│   ├── Utsushi.Parsing.Tests/
│   ├── Utsushi.ReportDefinition.Tests/
│   ├── Utsushi.Substitution.Tests/
│   ├── Utsushi.Layout.Tests/
│   ├── Utsushi.Rendering.Tests/
│   └── Utsushi.Golden.Tests/       # 登録済み帳票サンプルによるゴールデン(回帰)テストと、エッジケース検証用サンプルの変換テスト
│       └── Fixtures/<帳票コード>/  # レビュー済みの期待出力
├── tools/
│   ├── Utsushi.SampleGenerator/    # 帳票サンプル(.xlsx)の生成ツール。製品コードからは参照されない
│   └── edge-case-samples/          # エッジケース検証用サンプルの生成スクリプト(Python + openpyxl。開発用のみ)
├── samples/
│   ├── reports/<帳票コード>/       # definition.json(手書き)と template.xlsx(生成物)
│   └── edge-cases/<帳票コード>/    # エッジケース検証用(利用者が作りがちなテンプレート)。登録済み帳票ではない
├── docs/                           # 人が読む補足ドキュメント
├── .kiro/
│   ├── steering/                   # 本ファイル群。プロジェクト全体に常時適用される方針
│   └── specs/                      # 機能ごとの要件定義書・設計書・タスクリスト
├── .claude/
│   └── agents/                     # コードレビュー・テスト作成・セキュリティレビュー等のサブエージェント定義
├── Directory.Build.props           # 全プロジェクト共通のビルド設定(TFM/LangVersion/Nullable、VS2019コンパイラでの検証スイッチ)
├── Directory.Build.targets         # テストプロジェクト共通設定(RollForward/テストパッケージ)
└── Utsushi.sln                     # classic形式。Visual Studio 2019でもそのまま開ける
```

> ソリューションファイルはルート直下の `Utsushi.sln`(classic形式)1つだけとする。
> 呼び出し元プロダクトの開発環境である Visual Studio 2019 は新しいXML形式の `.slnx` を認識できないため、`.slnx` は置かない(`.kiro/steering/tech.md`「Visual Studio 2019 対応」参照)。
> ヘッダーの `# Visual Studio Version 16` はVS2019を示す値であり、VS2022の値(17)に書き換えない。.NET 10 SDKの `dotnet sln add/remove` はこのヘッダーを保持することを確認済み。
> ルートに `.sln` と `.slnx` が共存すると `dotnet build`/`dotnet test`/`dotnet format` の引数なし実行が「複数のプロジェクト/ソリューションファイルがある」エラーになるため、`.slnx` を追加しないこと。
> プロジェクトを追加・削除した場合は `dotnet sln Utsushi.sln add/remove src/<Project>/<Project>.csproj` で更新する(`<PATH>` はslnファイルの場所ではなく実行時のカレントディレクトリからの相対パス)。
> `dotnet build` / `dotnet test` / `dotnet format` はいずれもルートで引数なしに実行できる(既定で `Utsushi.sln` が使われる)。
> テストはリポジトリのルートを `Utsushi.sln` の有無で探す(`tests/Utsushi.TestSupport/TestPaths.cs`)ため、ソリューションファイルの名前・場所を変える場合はそちらも合わせて直す。

## 命名規則

- 名前空間・プロジェクト名は `Utsushi.<レイヤー名>` とする。
  - 例外: `Utsushi.ReportDefinition` プロジェクトの名前空間は `Utsushi.ReportDefinitions`(複数形)とする。
    帳票定義の型名が `ReportDefinition` であり、同名の名前空間と衝突して参照が曖昧になるため。
- 帳票定義ファイルは `samples/reports/<帳票コード>/definition.json` のように帳票コード単位でディレクトリを分ける。
  テンプレートの Excel は同じディレクトリに `template.xlsx` として置く。
- テストのゴールデンファイルは `tests/Utsushi.Golden.Tests/Fixtures/<帳票コード>/` に配置する。

## 帳票の追加手順(標準ワークフロー)

共通レイヤーのコードは変更しない。以下の3つのデータを足すだけで完結する(要件8.1)。

1. `samples/reports/<帳票コード>/template.xlsx` に対象Excelファイル(個人情報等はマスキング)を追加する。
   テンプレートの作り方・登録前のチェックリストは `docs/テンプレート作成ガイド.md` を参照。
2. 同じディレクトリに `definition.json`(帳票定義)を追加する。スキーマは `docs/帳票定義スキーマ.md` を参照。
3. `tests/Utsushi.Golden.Tests/ReportConversionGoldenTests.RegisteredReports` に帳票コードと置換値を追加し、
   テストを1度実行してゴールデンファイルを生成する。**生成された内容を必ずレビューしてからコミットする**
   (期待値の更新は `UTSUSHI_UPDATE_GOLDEN=1 dotnet test`)。
4. 新規要件が既存のレイアウト計算で表現できない場合のみ、`.kiro/specs/` に新しい要件を追記し、`Utsushi.Layout` を拡張する。

## ドキュメントとサブエージェントの関係

- 新機能・大きな変更に着手する前に、`.kiro/specs/<機能名>/` に `requirements.md` → `design.md` → `tasks.md` の順で仕様を作成・更新する(spec駆動開発)。
- 実装完了後は `code-reviewer`、テスト未整備の場合は `test-writer`、外部入力を扱う変更(Excel読み込み・ファイルパス処理等)には `security-reviewer`、仕様との整合性確認には `spec-compliance-reviewer` を用いる。各サブエージェントの詳細は `.claude/agents/` を参照。
