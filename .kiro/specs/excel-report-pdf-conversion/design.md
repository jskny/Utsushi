# 設計書: 自社帳票のExcel→PDF変換

対象要件: `.kiro/specs/excel-report-pdf-conversion/requirements.md`

## 概要

自社帳票専用のExcel→PDF変換パイプラインを、責務ごとに5つのレイヤーへ分割して実装する。
各レイヤーは下位レイヤーの実装詳細を知らず、内部モデル(POCO)を介してのみやり取りする。
これにより「Excelの解析方法」や「PDFの描画方法」を後から差し替えても、帳票定義や置換ロジックに影響が及ばないようにする。

```
[.xlsx ファイル]
      │  (1) Parsing
      ▼
WorkbookModel (内部モデル: セル値/スタイル/ページ設定)
      │  (2) ReportDefinition ロード + 突合
      ▼
ReportModel (帳票定義と結び付いた内部モデル)
      │  (3) Substitution
      ▼
ReportModel (置換キー→値を反映済み)
      │  (4) Layout
      ▼
PagedLayout (ページ分割・座標計算済みの描画命令列)
      │  (5) Rendering
      ▼
[.pdf ファイル]
```

## コンポーネントとインターフェース

### 1. Parsing レイヤー (`Utsushi.Parsing`)

- **責務**: `DocumentFormat.OpenXml` を用いて `.xlsx` を読み込み、`WorkbookModel` を構築する。
- **入力**: ファイルパス or Stream
- **出力**: `WorkbookModel`(シート・セル・スタイル・ページ設定を保持するPOCO)
- **主なインターフェース**:
  ```csharp
  public interface IWorkbookReader
  {
      WorkbookModel Read(Stream xlsxStream);
  }
  ```
- OpenXml SDK以外の型(`SpreadsheetDocument` 等)を `WorkbookModel` の外に漏らさない。上位レイヤーは `WorkbookModel` のみを参照する。
- 取得するページ設定: 印刷範囲(`definedNames` の `_xlnm.Print_Area`)、手動改ページ(`rowBreaks`/`colBreaks`)、用紙サイズ・余白・拡大縮小(`pageSetup`)、印刷タイトル(`_xlnm.Print_Titles`)、印刷順序。

### 2. ReportDefinition レイヤー (`Utsushi.ReportDefinition`)

- **責務**: 帳票定義(JSON)をロードし、`WorkbookModel` と突合して `ReportModel` を構築する。帳票ごとの固有情報(置換キーとセル番地のマッピング、はみ出し時の挙動、許容誤差等)はすべてここに閉じ込める。
- **帳票定義スキーマ(例)**:
  ```json
  {
    "reportCode": "invoice",
    "sheetName": "請求書",
    "substitutionFields": [
      { "key": "InvoiceNo", "cell": "C3", "required": true, "overflow": "shrink" },
      { "key": "IssueDate", "cell": "C4", "required": true, "overflow": "clip" }
    ],
    "toleranceMm": 0.5,
    "unsupportedElements": "error"
  }
  ```
- **主なインターフェース**:
  ```csharp
  public interface IReportDefinitionRepository
  {
      ReportDefinition Load(string reportCode);
  }

  public interface IReportModelBuilder
  {
      ReportModel Build(WorkbookModel workbook, ReportDefinition definition);
  }
  ```
- `Build` 時に、定義が参照するシート名・セル番地が `WorkbookModel` 上に実在するかを検証する(要件1.4)。

### 3. Substitution レイヤー (`Utsushi.Substitution`)

- **責務**: 置換キーと値の辞書を受け取り、`ReportModel` 上の対象セルの値のみを書き換える。書式(スタイル)には一切触れない。
- **主なインターフェース**:
  ```csharp
  public interface ICellSubstitutor
  {
      ReportModel Apply(ReportModel report, IReadOnlyDictionary<string, string> values);
  }
  ```
- 必須キー未指定・未知キー指定はここで例外(要件2.3, 2.4)。
- はみ出し時の挙動(`overflow: shrink|clip|wrap`)は帳票定義の値をそのままLayoutレイヤーに引き渡すためのフラグとして `ReportModel` に保持する(実際の折り返し/縮小計算はLayoutレイヤーの責務)。

### 4. Layout レイヤー (`Utsushi.Layout`)

- **責務**: `ReportModel` から、ページ分割済み・座標確定済みの描画命令列 `PagedLayout` を計算する。Excelのレイアウトルールのうち、対象帳票に必要な範囲を再現する。
- **主な計算**:
  - 列幅(文字単位)→ ポイント換算(既定フォントの文字幅メトリクスを使用)
  - 行高(ポイント)の反映
  - 印刷範囲のクリッピング
  - 手動改ページの適用
  - 自動改ページ計算(1ページの印字可能領域 = 用紙サイズ - 余白、を拡大縮小率で除した論理サイズに対し、行高/列幅の累積が収まる位置で分割)
  - 印刷タイトル行/列の複製
  - 結合セルの矩形統合
  - セル内テキストのフォントメトリクスに基づく配置(左右/上下揃え、インデント、縮小表示)
- **主なインターフェース**:
  ```csharp
  public interface IReportLayoutEngine
  {
      PagedLayout Compute(ReportModel report);
  }
  ```
- `PagedLayout` は「ページのリスト」であり、各ページは「描画すべき矩形(セル背景・罫線・テキストラン)のリスト」を持つ、Renderingレイヤーに依存しない中間表現とする。

### 5. Rendering レイヤー (`Utsushi.Rendering`)

- **責務**: `PagedLayout` を `SkiaSharp` の `SKDocument`(PDFバックエンド)で描画し、PDFファイルを生成する。
- **主なインターフェース**:
  ```csharp
  public interface IPdfRenderer
  {
      void Render(PagedLayout layout, Stream output);
  }
  ```
- 1ページ = 1 `SKCanvas` への描画。矩形塗りつぶし(背景)→罫線→テキストの順で描画し、Excelの重なり順を再現する。
- フォントは埋め込み(サブセット化)を基本とし、対象帳票が使用するフォントが実行環境に存在しない場合はビルド/デプロイ時にエラーとする(実行時のフォールバックによる見た目崩れを避ける)。

## データモデル(概要)

```csharp
public sealed record WorkbookModel(IReadOnlyList<SheetModel> Sheets);

public sealed record SheetModel(
    string Name,
    IReadOnlyDictionary<CellAddress, CellModel> Cells,
    IReadOnlyList<MergedRange> MergedRanges,
    IReadOnlyList<double> ColumnWidths,
    IReadOnlyList<double> RowHeights,
    PageSetup PageSetup);

public sealed record CellModel(string? Value, CellStyle Style);

public sealed record CellStyle(
    FontStyle Font,
    BorderSet Borders,
    HorizontalAlignment HAlign,
    VerticalAlignment VAlign,
    string? NumberFormat,
    string? BackgroundColor);

public sealed record ReportModel(
    ReportDefinition Definition,
    SheetModel Sheet); // 置換適用後もこの型のまま(値のみ更新)

public sealed record PagedLayout(IReadOnlyList<PageLayout> Pages);

public sealed record PageLayout(
    PaperSize Paper,
    Orientation Orientation,
    IReadOnlyList<DrawCommand> Commands);
```

## エラーハンドリング方針

- 例外階層は `UtsushiException` を基底とし、以下を派生させる。
  - `ReportDefinitionNotFoundException`(要件1.4)
  - `UnsupportedWorkbookElementException`(要件1.5)
  - `SubstitutionKeyNotFoundException` / `RequiredSubstitutionValueMissingException`(要件2.3, 2.4)
  - `InvalidExcelFileException`(要件6.1, 6.2)
  - `ReportDefinitionSchemaException`(要件6.3)
- すべての例外は、帳票コード・シート名・セル番地・処理段階(Parsing/Substitution/Layout/Rendering)を構造化プロパティとして保持し、ログ出力時に特定できるようにする(要件6.4)。
- Renderingレイヤーは一時ファイル/一時ストリームに書き込み、正常終了時のみ最終出力先へ確定させる(要件5.4: 不完全PDFを残さない)。

## テスト戦略

- **ユニットテスト**: 各レイヤーのインターフェース単位(特にLayoutレイヤーの改ページ計算・フォントメトリクス換算、Substitutionレイヤーの必須/未知キー判定)。
- **ゴールデンテスト**: `samples/reports/` の帳票サンプルを実際に変換し、レビュー済みの期待出力(ページ数・各ページの主要な描画命令のスナップショット)と比較する。PDFバイナリの完全一致ではなく、意味のある差分(座標・テキスト・罫線)を検出できる比較方法を採用する。
- **契約テスト**: `ReportDefinition` のJSONスキーマ検証(不正な定義をロード時に検出できることを確認)。

## 未決事項 / 今後の検討

- フォントメトリクス取得の具体的な方法(SkiaSharpのフォントメトリクスAPIと、Excelの列幅(文字単位)換算式の対応関係)は、最初の帳票定義を実装する際にプロトタイプで検証し、本設計書を更新する。
- 数式を含むセルの扱い(値のみ読み取るか、限定的に再計算するか)は要件化されていないため、最初の対象帳票の内容を確認した上で `requirements.md` に追記する。
