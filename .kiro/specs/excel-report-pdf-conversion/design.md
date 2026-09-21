# 設計書: 自社帳票のExcel→PDF変換

対象要件: `.kiro/specs/excel-report-pdf-conversion/requirements.md`

## 概要

自社帳票専用のExcel→PDF変換パイプラインを、責務ごとに5つのレイヤーへ分割して実装する。
各レイヤーは下位レイヤーの実装詳細を知らず、内部モデル(POCO)を介してのみやり取りする。
これにより「Excelの解析方法」や「PDFの描画方法」を後から差し替えても、帳票定義や置換ロジックに影響が及ばないようにする。

すべてのレイヤーは共通基盤 `Utsushi.Core`(例外階層とレイヤーをまたぐ値型のみ)を参照してよい。
呼び出し元プロダクトが参照するのは、5レイヤーを組み立てるファサード `Utsushi`(`ReportPdfConverter`)だけである。

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

## 単位と座標系

内部表現の単位は **ポイント(pt、1pt = 1/72インチ)** に統一する(`.kiro/steering/tech.md`「コーディング規約」)。
帳票定義や外部仕様がミリ単位を使う場合も、レイヤーの境界を越える前にポイントへ変換する。
換算は `Utsushi.Core` の `Units` に集約する。

- **座標系**: ページ左上が原点。X軸は右方向、Y軸は**下方向**(SkiaSharp / PDF の描画APIと一致させるため)。
- **行高**: OOXML の `row/@ht` は元からポイント単位。換算不要。
- **余白**: OOXML の `pageMargins` はインチ。`Units.InchesToPoints` で換算する。
- **列幅**: OOXML の `col/@width` は「標準フォントで数字を何文字ぶん表示できるか」の単位で、
  96dpi のピクセルを経由して換算する。詳細は次節。

### 列幅のポイント換算(要件4.5)

設計当初の未決事項だったフォントメトリクスとの対応関係は、以下のとおり決定した。

ECMA-376 Part 1, 18.3.1.13 が定める換算式をそのまま用いる。

```
pixels = Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)
points = pixels * 72 / 96
```

- `MDW`(Maximum Digit Width)は標準フォントの最大数字幅(96dpiのピクセル)。
- `Truncate` が2回入るため、素朴な `width * MDW` とは最大1px程度ずれる。この誤差は列数ぶん累積して
  改ページ位置に影響しうるため、式を近似せずそのまま実装する(`ExcelUnitConverter`)。
- `MDW` はフォントとサイズに依存する。SkiaSharp のフォントメトリクスから自動導出するのではなく、
  **帳票定義の `maxDigitWidthPx`(既定7)で明示的に与える**方式とした。理由は2つある。
  - 実行環境にインストールされたフォントの版差で `MDW` が変わると、改ページ位置が環境依存になる。
  - 帳票ごとに標準フォントが異なりうるため、帳票固有の値は帳票定義側へ寄せる方針と整合する。

### フォントメトリクスの参照方法

Layout レイヤーはテキスト配置(左右/上下揃え、縮小表示、折り返し)に実フォントのメトリクスを必要とするが、
SkiaSharp に直接依存してはならない。そこで `IFontMetricsProvider` を **Layout 側に定義し、実装を Rendering 側に置く**
(依存性逆転)。実装は2つある。

| 実装 | 置き場所 | 用途 |
|---|---|---|
| `SkiaFontMetricsProvider` | Rendering | 実運用。実フォントのメトリクスを使う |
| `ApproximateFontMetricsProvider` | Layout | テスト・ゴールデン比較。全角1em/半角0.5emの決定的な近似 |

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
- **数式セル**: 数式は評価せず、OOXMLにキャッシュされている計算結果の値のみを読み取る(要件1.6)。
- **数値書式**: `numFmt` を適用した表示文字列を `CellModel.FormattedValue` に持たせる。
  汎用の数値書式エンジンではなく、自社帳票が使う範囲(金額・数量・日付・パーセント)のサブセット実装とする。
  解釈できない書式(指数表記・分数表記など)は例外にせず General 相当へフォールバックし、
  表示の崩れはゴールデンテストで検出する。

### 2. ReportDefinition レイヤー (`Utsushi.ReportDefinition`)

- **責務**: 帳票定義(JSON)をロードし、`WorkbookModel` と突合して `ReportModel` を構築する。帳票ごとの固有情報(置換キーとセル番地のマッピング、はみ出し時の挙動、許容誤差等)はすべてここに閉じ込める。
- **帳票定義スキーマ(例)**: 全プロパティの一覧と制約は `docs/帳票定義スキーマ.md` を参照。
  ```json
  {
    "schemaVersion": 1,
    "reportCode": "invoice",
    "sheetName": "請求書",
    "substitutionFields": [
      { "key": "InvoiceNo", "cell": "C3", "required": true, "overflow": "shrink" },
      { "key": "IssueDate", "cell": "C4", "required": true, "overflow": "clip" }
    ],
    "toleranceMm": 0.5,
    "unsupportedElements": "error",
    "maxDigitWidthPx": 7,
    "printArea": "A1:F34"
  }
  ```
- スキーマ検証には JSON Schema ライブラリを使わず、必須項目・型・値域・キー重複を明示的に検証する。
  スキーマが小さく固定的であり、依存を1つ減らせるため(`.kiro/steering/tech.md`「依存ライブラリ追加時のルール」)。
  検証エラーには問題のあったプロパティのパス(例: `substitutionFields[0].cell`)を含める(要件6.3)。
- 帳票コードは呼び出し元から渡されうるため、ディレクトリ名として解決する際に
  パス区切り文字や `..` によるルート外への脱出を拒否する。
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

      ReportModel ApplyCellOverrides(ReportModel report, IReadOnlyDictionary<string, string> cellOverrides);
  }
  ```
- 必須キー未指定・未知キー指定はここで例外(要件2.3, 2.4)。
- はみ出し時の挙動(`overflow: shrink|clip|wrap`)は帳票定義の値をそのままLayoutレイヤーに引き渡すためのフラグとして `ReportModel` に保持する(実際の折り返し/縮小計算はLayoutレイヤーの責務)。
- `ApplyCellOverrides` は、帳票定義の置換キー(`SubstitutionFields`)を経由せず、セル番地(A1形式の文字列。キーは `CellAddress.TryParse` で解釈する)を直接指定して値を書き換える第二の経路(要件2.7, 2.8)。
  - `Apply` と同じく `CellModel.WithText` で値のみを差し替え、書式には触れない。対象セルが未存在(空セル)の場合は既定スタイルの新規セルを作る点も `Apply` と同一。
  - `Apply` と同様、全入力を検証してから一括で書き換える二段階構成(`ParseAndValidateAddresses` → 書き換え)を取る。
  - `Apply` の未知キー検証(要件2.3)・必須キー検証(要件2.4)の対象外。帳票定義に登録の無いセルも指定できる。
  - セル番地がA1形式として解釈できない場合は `InvalidCellOverrideAddressException`(Stage=Substitution)を送出する(要件2.8)。
  - 指定したセルが結合セル範囲内にあり、かつ先頭(アンカー)セルでない場合は `NonAnchorMergedCellOverrideException` を送出する(要件2.9)。Layoutレイヤー(`PageCommandBuilder`)は結合範囲のアンカーセルの値のみを描画するため、アンカー以外への上書きは値がモデルには反映されてもPDFに一切出力されない静かなデータ欠落になる。これを防ぐため、Substitutionレイヤーの時点で検出する。
  - `OverflowByCell` には書き込まず、かつ上書き対象セルの既存エントリ(名前付きキー方式`Apply`が同じセルに残した`overflow`指定)があれば削除する。これにより、同じセルを`Apply`と`ApplyCellOverrides`の両方で扱った場合でも、はみ出し時は常にExcel側のセル書式(`wrapText`/`shrinkToFit`)に従う(帳票定義でのはみ出し挙動指定は、この経路では行えない)。
  - `Utsushi.ReportPdfConverter` は `Convert` / `ConvertToFile` / `ComputeLayout` に任意パラメータ `IReadOnlyDictionary<string, string>? cellOverrides = null` を追加し、`Apply` の後に `ApplyCellOverrides` を適用する(名前付きキーでの必須値検証を経てから、セル直接指定で上書きできるようにするため)。CLIは `--override <セル番地>=<値>` オプションでこれを渡す。

### 4. Layout レイヤー (`Utsushi.Layout`)

- **責務**: `ReportModel` から、ページ分割済み・座標確定済みの描画命令列 `PagedLayout` を計算する。Excelのレイアウトルールのうち、対象帳票に必要な範囲を再現する。
- **主な計算**:
  - 列幅(文字単位)→ ポイント換算(既定フォントの文字幅メトリクスを使用)
  - 行高(ポイント)の反映
  - 印刷範囲のクリッピング
  - 手動改ページの適用
  - 自動改ページ計算(1ページの印字可能領域 = 用紙サイズ - 余白、を拡大縮小率で除した論理サイズに対し、行高/列幅の累積が収まる位置で分割)
  - 印刷タイトル行/列の複製(印刷タイトルは印刷範囲と独立に指定でき、印刷範囲の外の行/列でも各ページに繰り返す。要件3.4)
  - 複数の印刷範囲をそれぞれ独立したページ群として計算(要件3.6)
  - 必須の置換フィールドの対象セルが、印刷範囲・印刷タイトルのいずれにも含まれない場合はエラーとする(要件2.6)
  - ページヘッダー/フッターの書式コード展開と配置(要件3.7〜3.9)
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
- 出力するPDFのバージョンは SkiaSharp の PDF バックエンドが生成する **PDF 1.4** とする(要件5.3)。
- 対象帳票が使用するフォントが実行環境に存在しない場合は `FontNotAvailableException` で失敗させる
  (`FontResolver` の既定は厳格モード)。実行時の暗黙フォールバックによる見た目崩れを避けるため。
  検証だけを先に行いたい場合は `FontResolver.EnsureAvailable` を使う。
- **フォント埋め込み方針(決定済み)**: PDF内の文字列検索・コピーを維持するため、
  **フォントを埋め込む `PdfTextRendering.EmbedFont` を既定とする**。
  SkiaSharp の PDF バックエンド(NuGetで配布されるネイティブビルド)は
  フォントのサブセット化を行わず使用フォントを丸ごと埋め込むため(2.88 系・3.x 系の双方で実測確認)、
  日本語帳票では出力PDFが4MB前後になるが、この配布サイズは許容する。
  サイズを優先したい場合のために、文字をベクタのアウトラインとして出力する
  `PdfTextRendering.Outline` も選べるようにしてある(同一帳票で 4.3MB → 91KB、見た目は同一)。
  ただしPDF内の文字列検索・コピー・テキスト抽出ができなくなる。
- **太字・斜体の再現**: 実フォントが太字/斜体の字形を持つ場合はその実字形を使う。
  持たない場合(MS PGothic や IPAGothic など、Excel 側も合成で表示しているもの)は、
  **通常字形を埋め込んだうえで描画時に装飾を合成する**。
  - 太字: 輪郭を塗りと一緒に描く(em の3%)。文字送り幅は変わらないためレイアウトに影響しない。
  - 斜体: `SKFont.SkewX` で傾ける。
  - SkiaSharp に書体レベルで合成させると PDF が **Type 3 フォント**になり、
    文字列検索・コピー・テキスト抽出ができなくなるうえファイルサイズも増えるため、この方式を採る。
  - 実字形の有無は、通常字形の書体とフォントデータ(バイト長と先頭4KB)が一致するかで判定する。
    `SKTypeface.FromFamilyName` は字形が無くても合成書体を返すため、書体の属性では判別できない。
  - 実字形の太字フォントファイルがある場合は、`FontResolverOptions.FontFiles` に
    `ファミリ名:bold` のキーで登録すれば合成せずにそれを使う。
- 失敗時に不完全なPDFを残さないため、いったんメモリ上に完全なPDFを作ってから出力先へ転送する。
  ファイル出力では同一ディレクトリ上の一時ファイルへ書いてから置換する(要件5.4)。

## データモデル(概要)

```csharp
// --- Parsing レイヤー ---
public sealed record WorkbookModel(IReadOnlyList<SheetModel> Sheets, FontStyle DefaultFont);

public sealed record SheetModel(
    string Name,
    IReadOnlyDictionary<CellAddress, CellModel> Cells,
    IReadOnlyList<MergedRange> MergedRanges,
    IReadOnlyList<double> ColumnWidths,   // 索引0が列A。単位はExcelの「文字数」
    IReadOnlyList<double> RowHeights,     // 索引0が行1。単位はポイント
    double DefaultColumnWidth,            // 上記リストの範囲外の列に適用する既定値
    double DefaultRowHeight,
    IReadOnlySet<int> HiddenColumns,      // 非表示行/列は印刷されないため保持する
    IReadOnlySet<int> HiddenRows,
    PageSetupModel PageSetup);

public sealed record CellModel(
    string? Value,                        // Excelが保持している生の値
    CellValueKind ValueKind,              // Blank/Text/Number/Boolean/Error
    CellStyle Style,
    string? FormattedValue = null,        // 数値書式を適用した表示文字列
    bool HasFormula = false);

public sealed record CellStyle(
    FontStyle Font,
    BorderSet Borders,
    HorizontalAlignment HAlign,
    VerticalAlignment VAlign,
    string? NumberFormat,
    ArgbColor BackgroundColor,
    bool WrapText,
    bool ShrinkToFit,
    int Indent);

// --- ReportDefinition / Substitution レイヤー ---
public sealed record ReportModel(
    ReportDefinition Definition,
    SheetModel Sheet,                     // 置換適用後もこの型のまま(値のみ更新)
    IReadOnlyDictionary<CellAddress, OverflowBehavior> OverflowByCell,
    FontStyle DefaultFont);               // ヘッダー/フッターの既定フォント

// --- Layout レイヤー ---
public sealed record PagedLayout(
    IReadOnlyList<PageLayout> Pages, string ReportCode, string SheetName);

public sealed record PageLayout(
    PaperSize Paper,
    PageOrientation Orientation,
    double WidthPt,                       // 向きを適用した実寸
    double HeightPt,
    IReadOnlyList<DrawCommand> Commands,  // 背景 → 罫線 → テキストの順
    int PageNumber,
    (int First, int Last) RowRange,       // 診断・テスト用
    (int First, int Last) ColumnRange,
    double ScaleFactor);                  // 座標には適用済み。診断用に保持する

// 描画命令。座標はページ左上原点のポイントで、余白・拡大縮小を適用済み。
public abstract record DrawCommand;
public sealed record FillRectCommand(RectPt Rect, ArgbColor Color) : DrawCommand;
public sealed record LineCommand(
    PointPt From, PointPt To, ArgbColor Color, double WidthPt, LineDashStyle Dash) : DrawCommand;
public sealed record TextCommand(
    PointPt Origin, string Text, FontStyle Font, TextAnchor Anchor, RectPt? ClipRect) : DrawCommand;
```

`TextCommand.Origin` の X は `Anchor`(Left/Center/Right)の基準点、Y はベースライン位置を表す。
文字列の実際の幅は描画時のフォントで決まるため、Layout は基準点だけを確定させ、
左右揃えの最終的な字送りは Rendering が行う。

## エラーハンドリング方針

- 例外階層は `UtsushiException`(`Utsushi.Core.Exceptions`)を基底とし、以下を派生させる。
  - `ReportDefinitionNotFoundException`(要件1.4)
  - `ReportStructureMismatchException`(要件1.4: シート名・セル番地の不一致)
  - `UnsupportedWorkbookElementException`(要件1.5)
  - `SubstitutionKeyNotFoundException` / `RequiredSubstitutionValueMissingException`(要件2.3, 2.4)
  - `InvalidCellOverrideAddressException`(要件2.8。セル番地直接指定がA1形式として解釈できない場合)
  - `NonAnchorMergedCellOverrideException`(要件2.9。セル番地直接指定の対象が結合セル範囲の非アンカー位置の場合)
  - `InvalidExcelFileException`(要件6.1, 6.2, 6.5。`Reason` で非xlsx/破損/パスワード保護/ファイルを開けない(存在しない・アクセス不可)を区別する)
  - `ReportDefinitionSchemaException`(要件6.3。問題のあったプロパティパスを保持する)
  - `LayoutComputationException` / `PdfRenderingException` / `FontNotAvailableException`
- すべての例外は、帳票コード・シート名・セル番地・処理段階(Parsing/Substitution/Layout/Rendering)を構造化プロパティとして保持し、ログ出力時に特定できるようにする(要件6.4)。
- Renderingレイヤーは一時ファイル/一時ストリームに書き込み、正常終了時のみ最終出力先へ確定させる(要件5.4: 不完全PDFを残さない)。

## テスト戦略

- **ユニットテスト**: 各レイヤーのインターフェース単位(特にLayoutレイヤーの改ページ計算・フォントメトリクス換算、Substitutionレイヤーの必須/未知キー判定)。
- **ゴールデンテスト**: `samples/reports/` の帳票サンプルを実際に変換し、レビュー済みの期待出力と比較する。
  比較対象は **PDFバイナリではなく `PagedLayout` を行指向テキストにしたスナップショット** とする。
  PDFバイナリは生成日時・圧縮・SkiaSharpのバージョンで変わり、意味のない差分が出るため。
  スナップショットは1行1描画命令で、座標は0.01pt(約0.0035mm)に丸める。
  帳票定義の許容誤差(既定0.5mm)より十分細かく、浮動小数の最下位ビットの揺れは吸収できる粒度である。
  - 期待値の置き場所: `tests/Utsushi.Golden.Tests/Fixtures/<帳票コード>/layout.snapshot.txt`
  - 更新方法: `UTSUSHI_UPDATE_GOLDEN=1 dotnet test`。**差分は必ずレビューしてからコミットする**。
  - レイアウト計算には決定的な `ApproximateFontMetricsProvider` を使う。実フォントのメトリクスは
    実行環境のフォント構成に依存し、CIとローカルでゴールデンが一致しなくなるため。
  - 実フォントを使った経路は「PDFが生成できること」までを自動テストの範囲とし、
    最終的な見た目のレビューは人が行う。
- **契約テスト**: `ReportDefinition` のJSONスキーマ検証(不正な定義をロード時に検出できることを確認)。
  `samples/reports/` 配下の実際の定義がすべてロードできることも確認する。

## 決定済みの旧未決事項

- ~~フォントメトリクス取得の具体的な方法~~ → 「単位と座標系」の節に記載。
  列幅換算は ECMA-376 の式を用い、MDW は帳票定義の `maxDigitWidthPx` で与える。
  テキスト配置用のフォントメトリクスは `IFontMetricsProvider` 越しに参照する。
- ~~数式を含むセルの扱い~~ → 数式は評価せず、キャッシュ済みの計算結果の値のみを読み取る。
  `requirements.md` の要件1.6 として要件化済み。
- ~~PDFにフォントを埋め込むかアウトライン化するか~~ → PDF内検索を維持するため埋め込みを既定とする。
  サイズ(日本語帳票で4MB前後)は許容する。
- ~~太字の合成(fake bold)による Type 3 フォント化~~ → 通常字形を埋め込み、描画時に輪郭を
  太らせて再現する方式に変更。Type 3 にならず、文字列検索も維持される。
- ~~複数の印刷範囲~~ → 範囲ごとに独立したページ群として出力する。要件3.6 として要件化済み。
- ~~ヘッダー/フッター~~ → 読み取りと描画に対応。要件3.7〜3.9 として要件化済み。

## 未決事項 / 今後の検討

- **改ページをまたぐ結合セルの文字位置**: 結合範囲がページ境界をまたぐ場合、現状は
  「そのページに見えている部分」を1つの矩形として扱い、その中にテキストを配置する。
  Excel は結合範囲全体を基準に文字を配置したうえでページ境界で切り取るため、
  両ページに同じ文字列が現れる点が異なる。印刷タイトル行にある結合セル
  (各ページに繰り返すことが正しいもの)とは区別が必要なため、対象帳票で
  「本文中の結合セルが改ページをまたぐ」ケースが出てから実装する。
  なお罫線については、範囲の本当の上端/下端/左端/右端がそのページに含まれる場合のみ
  対応する辺を描画するよう対応済み(改ページの切れ目には罫線を出さない)。文字位置のみ未対応。
- **PDFのフォントサブセット化**: 「フォントを埋め込む」方針を採ったため、
  出力PDFは使用フォント全体を含む(日本語帳票で4MB前後)。
  配布サイズを下げつつ検索性も保つにはサブセット化が必要だが、SkiaSharp のネイティブビルドが
  対応していないため、PDF生成後の後処理か別ライブラリの検討が要る。
  現時点ではサイズを許容する判断のため着手しない。
- **ヘッダー/フッターの未対応コード**: 画像(`&G`)と文字色(`&K`)は読み飛ばしている。
  対象帳票で必要になった時点で要件化する。
  `&F`(ファイル名)は、Utsushi が Stream を入力に取り元のファイル名を持たないため、
  帳票コードを代わりに展開している。
