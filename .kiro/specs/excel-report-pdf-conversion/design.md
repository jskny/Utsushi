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
- **画像(要件9)**: `WorksheetPart.DrawingsPart.WorksheetDrawing` 配下の `xdr:twoCellAnchor` /
  `xdr:oneCellAnchor` のうち `xdr:pic`(画像)のみを対象とする。`a:blip` の `r:embed` から
  `ImagePart` を解決し、バイナリ(`GetStream()`)と `ContentType` を読み取る。
  `ContentType` が `image/png` / `image/jpeg` / `image/gif` / `image/bmp` のいずれでもない場合
  (EMF/WMF等のベクタ形式を含む)は「サポート外要素」として扱い、既存の
  `DetectUnsupportedElements`(`OpenXmlWorkbookReader`)と同じ `unsupportedElements` ポリシーに
  従う(`ElementKind = "UnsupportedImageFormat"`)。対応形式かどうかは
  `ContentType` の静的な許可リストのみで判定し、SkiaSharpによる実デコード確認はしない
  (Rendering層のSkiaSharpにParsing層が依存しないため。`.kiro/steering/structure.md`の
  レイヤー依存方向)。実際にSkiaSharpでデコードできない不正なバイナリだった場合は
  Renderingレイヤーで `PdfRenderingException` になる。
  - **既存の`DetectUnsupportedElements`との整合(要修正点)**: 現状の実装
    (`OpenXmlWorkbookReader.cs:534`)は `DrawingsPart is not null` の時点で無条件に
    `UnsupportedWorkbookElementException("Drawing")` を送出しており、中身が画像だけでも
    即座に止まってしまう。画像対応の実装では、この判定を「`DrawingsPart` 内のアンカーを
    列挙し、`xdr:pic` 以外(`xdr:sp`/`xdr:grpSp`/`xdr:cxnSp` 等)が1つでもあれば
    `ElementKind = "Drawing"` として例外化、`xdr:pic` のみで構成される場合は例外化せず
    画像読み取りへ進む」という分岐に置き換える必要がある(`xdr:graphicFrame` = グラフは
    既存どおり525行目で先に個別検出される)。この置き換えを行わない限り、画像を1つでも
    含むシートは `unsupportedElements: "error"` の帳票定義で常に失敗し続ける。
  - **信頼できない入力に対する安全弁(要件9.6)**: `ReportPdfConverter.Convert` は呼び出しごとに
    外部から供給される `xlsxStream` をそのまま処理するため、埋め込み画像のバイナリも
    「一度だけ人が確認して固定化された資産」ではなく信頼できない入力として扱う
    (security-reviewer指摘)。以下の上限・検証を設ける。
    - 1シートあたりの画像アンカー数の上限(既定50枚)。超過分は `unsupportedElements` の
      設定に従う(`ElementKind = "TooManyImages"`)。
    - 画像1枚あたりの読み取りバイト数の上限(既定10MB)。超過時も同様
      (`ElementKind = "ImageTooLarge"`)。
    - `ContentType` は宣言に過ぎず実バイト列と一致する保証がないため、ファイル先頭の
      シグネチャ(PNG/JPEG/GIF/BMPのマジックバイト)が一致することを確認する。
      不一致の場合は `ElementKind = "UnsupportedImageFormat"` として扱う。
    - Renderingレイヤーでは、`SKBitmap.Decode` で実際にデコードする前に
      `SKBitmap.DecodeBounds` で宣言上の幅・高さを確認し、上限(既定4096px)を超える場合は
      `PdfRenderingException` とする(数百バイトのファイルが巨大な展開後サイズを要求する
      「ピクセル爆弾」対策)。
    - 2セルアンカー(対角セル指定)の幅・高さ計算は、対角セルにセル番地の上限
      (最大1,048,576行×16,384列)近くを指定された場合の計算量を抑えるため、
      合算する列/行数に上限(既定4096)を設ける。
- **図形(要件10)**: 画像(`xdr:pic`)と同じ `xdr:twoCellAnchor` / `xdr:oneCellAnchor` の下に
  現れる `xdr:sp`(シェイプ)のうち、`xdr:spPr/a:prstGeom/@prst` が対応済みプリセット一覧
  (要件10.1補足)に含まれるものだけを読み取る。アンカー(左上セル・オフセット・
  固定/セル追従の範囲)は画像と全く同じ形式のため、`ImageModel`/`ShapeModel` 共通の
  基底として抽出した `AnchorExtent`(旧`ImageExtent`。`FixedAnchorExtent`/
  `CellSpanAnchorExtent`)をそのまま流用する。
  - **z-order(要件10.3)の扱い**: Excelは画像・図形を区別せず、`drawing.xml` 内の出現順で
    重なりを決める。この順序を保つため、`SheetModel.Images` を廃止し、
    `SheetModel.DrawingObjects: IReadOnlyList<DrawingObjectModel>`
    (`ImageModel`/`ShapeModel` はいずれも `DrawingObjectModel` を継承)に置き換える。
    `ReadImages`/`ReadShapes` は個別に呼ばず、アンカーを1回だけ列挙しながら
    `xdr:pic`/`xdr:sp` を判別して1つの順序付きリストを構築する
    (`ReadDrawingObjects` に統合)。
  - **プリセットの判定と非対応プリセットの扱い**: `@prst` の値を `ShapePresetType` へ
    マッピングする静的な辞書と突き合わせる。一致しない場合(星形・フローチャート記号・
    自由曲線 `custGeom` 等)は要件10.7により「サポート外要素」として扱う
    (`ElementKind = "UnsupportedShapePreset"`)。
  - **既存の`DetectUnsupportedElements`/`HasNonPictureDrawingObject`との整合**: 画像対応時に
    「`xdr:pic` 以外が1つでもあれば `Drawing` として例外化」としていた判定を、
    「`xdr:pic` および `xdr:sp`(シェイプ)以外(接続線`xdr:cxnSp`、グループ`xdr:grpSp`、
    図表枠は別途検出済み)が1つでもあれば `Drawing` として例外化」に拡張する。
    ここでの`xdr:sp`の判定は**構造的**(要素の種類がシェイプかどうか)であり、
    プリセットが対応済み一覧に含まれるかどうかは問わない。プリセットの対応可否は
    画像の`ContentType`許可リスト判定(要件9.4)と同じ位置付けで、後段の図形読み取り
    (`ReadShape`)が個別に検証し、非対応プリセットは`ElementKind = "UnsupportedShapePreset"`
    として`unsupportedElements`ポリシーに従う(下記「プリセットの判定と非対応プリセットの扱い」)。
    この2段構えにより、`UnsupportedShapePreset`が「画像の`UnsupportedImageFormat`」と
    同じ経路(`DetectUnsupportedElements`を通過した後の個別検証)で意味を持つ。
  - **幾何情報**: プリセット種別に加え、`a:avLst/a:gd`(調整ガイド)の `name`/`fmla="val N"`
    を `name → N/100000.0` の辞書として読み取り、`ShapePresetType` ごとに定義した
    ガイド名の並び順(例: `rightArrow` なら `["adj1", "adj2"]`)で `IReadOnlyList<double>`
    に整形する(該当ガイドが無ければそのプリセットのECMA-376既定値を使う。既定値表は
    Renderingレイヤーに置く。詳細は後述)。パス生成そのものはRenderingレイヤーの責務であり、
    Parsingレイヤーは数値の抽出のみを行う。
  - **塗りつぶし・枠線**: `xdr:spPr/a:solidFill` は単色、`a:gradFill/a:gsLst` は
    先頭と末尾の `a:gs` の色のみを開始色・終了色として採用し(要件10.6)、
    `a:lin/@ang`(60,000分の1度)があれば角度として読み取る(無ければ既定角度0度=左から右)。
    `a:noFill` は塗りなし(`Fill = null`)。枠線は `a:ln` の `a:solidFill` の色と
    `@w`(EMU)から変換した太さを読み取り、`a:ln` 自体が無い/`a:noFill` の場合は
    `Outline = null` とする。
  - **回転**: `xdr:spPr/a:xfrm/@rot`(60,000分の1度、時計回り)を `RotationDegrees`
    (度)に変換する。座標系が「Y軸下方向」(単位と座標系の節)のため、そのままの符号で
    時計回りの回転として扱える。
  - **テキスト**: `xdr:txBody` があれば `a:bodyPr/@anchor`(t/ctr/b)を垂直配置、
    各 `a:p/a:pPr/@algn`(l/ctr/r)を段落ごとの水平配置として読み取り、
    `a:r/a:rPr`(サイズ・太字・斜体・色・書体)と `a:t` からセルの `FontStyle` と
    同じ型で `ShapeTextRun` を構築する。折り返し(禁則処理を含まない単純な幅基準の折り返し)
    はLayoutレイヤーが `IFontMetricsProvider` を使って行うため、Parsingレイヤーでは
    段落・ランをそのまま保持するだけで折り返しは行わない。
  - **信頼できない入力に対する安全弁(要件10.8)**: 画像と同様、図形も無条件に信頼できる
    入力ではない。以下の上限を設ける。
    - 1シートあたりの図形アンカー数の上限(既定50個。画像の上限とは独立にカウントする)。
      超過分は `unsupportedElements` の設定に従う(`ElementKind = "TooManyShapes"`)。
    - 図形1つに含まれる全テキスト(段落・ランを連結した文字数)の上限(既定2000文字)。
      超過時も同様(`ElementKind = "ShapeTextTooLong"`)。文字数を実際に折り返し計算へ
      渡す前に拒否することで、極端に長い文字列に対する折り返し計算量を避ける。
    - 幅・高さの計算に使う列/行数の上限(既定4096)は画像と共通の
      `SpanWidthPt`/`SpanHeightPt` をそのまま流用するため、別途の上限追加は不要。

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
  - **画像・図形の配置(要件9, 10)**: `SheetModel.DrawingObjects`(画像・図形が`drawing.xml`の
    出現順で混在するリスト)を先頭から順に処理し、アンカーセルの位置(結合セルの
    矩形計算と同じ`SheetGrid`の列幅/行高累積・拡大縮小の適用)にセル内オフセットを加えて
    ページ左上原点のポイント座標へ変換する。
    2セルアンカー(対角セル指定)の幅・高さは、`SheetGrid`(印刷範囲にクリップされた格子)
    ではなく`SheetModel`から直接取得した列幅/行高で計算する。`SheetGrid`は印刷範囲外の列/行を
    「幅0」として保持しないため、これを流用すると対角セルが印刷範囲のすぐ外にあるだけで
    画像・図形が実際より小さく計算されてしまう(非表示列/行は0として扱う点は結合セル等の
    既存ロジックと同様。Excel自体は印刷範囲の設定に関わらず実際の列幅でサイズを決めるため、
    これに合わせる)。改ページ位置をまたぐ画像・図形の扱いは「未決事項」を参照。
    `DrawingObjects`の出現順をそのまま`DrawCommand`の出現順として`Commands`リストに
    追加する(セル内容の描画コマンドより後ろにまとめて追加する点は画像単独の場合と同じ。
    「エラーハンドリング方針」節の直前の描画順序の説明を参照)。
  - **図形内テキストの折り返し・配置(要件10.4)**: `ShapeModel.Text`(段落・ランの木構造)を、
    セル内テキストの折り返しと同じ`IFontMetricsProvider`を使い、図形の矩形幅を基準に
    単純な幅基準の折り返し(禁則処理なし。セル内テキストの折り返しと同水準)で複数行に
    分割する。各行の水平位置は段落の`HAlign`、行全体の垂直位置は`VAlign`と行数から
    (セル内テキストの上下揃えと同じ考え方で)算出し、`ShapeCommand.TextLines`の
    各`ShapeTextLine`として矩形内のポイント座標(回転前、シェイプ自身のローカル座標)を
    確定させる。回転の適用はRenderingレイヤーの責務とする(座標変換をLayoutに持ち込むと
    `PagedLayout`が回転行列という新しい概念を持つことになり、既存の「軸に平行な矩形の
    集まり」という単純なモデルから外れるため)。
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
- 1ページ = 1 `SKCanvas` への描画。矩形塗りつぶし(背景)→罫線→テキスト→**画像・図形**
  (`drawing.xml`の出現順)の順で描画する。Excelはシート上に浮かぶ描画オブジェクト
  (画像・図形等)をセルの内容より上のレイヤーとして描画するため、画像・図形は他のセル内容と
  重なる場合に最前面へ来るようにする(要件9.3, 10.3)。
- **画像の描画(要件9)**: `ImageCommand` は `SKBitmap.Decode(byte[])`でデコードし、
  `SKCanvas.DrawBitmap(bitmap, destRect)` で `ImageCommand.Rect` へ描画する
  (SkiaSharp 2.88.8で利用可能な標準API)。既存の `ToSkRect(RectPt)` をそのまま使う。
  ページ境界外にはみ出す部分は `SKCanvas` が自然にクリップするため、追加のクリップ処理は不要。
  `SKBitmap.Decode` で実際に展開する前に `SKBitmap.DecodeBounds` で宣言上のピクセル寸法を確認し、
  上限(既定4096px)を超える場合はデコードせず `PdfRenderingException` とする(要件9.6。
  ピクセル爆弾対策。詳細はParsingレイヤー節「信頼できない入力に対する安全弁」を参照)。
- **図形の描画(要件10)**: `ShapeCommand` ごとに、回転がある場合は
  `canvas.Save()` → `canvas.RotateDegrees(RotationDegrees, centerX, centerY)`
  (中心は`Rect`の中心)→ 描画 → `canvas.Restore()` で図形本体とテキストの両方を
  まとめて回転させる(要件10.5)。
  - **パス生成**: `ShapePresetType`ごとに、`Rect`のローカル座標(左上原点、幅・高さ)と
    `AdjustmentValues`(Parsingレイヤーが`a:avLst`から抽出済み。要素が無ければ
    ECMA-376の既定値を`ShapeGeometryBuilder`内の定数表から補う)から`SKPath`を組み立てる
    `ShapeGeometryBuilder`をRenderingレイヤー内に新設する。
    - `rect`: そのままの矩形。
    - `roundRect`: 角丸半径 = `min(幅, 高さ) * adj1`(既定`adj1 = 0.16667`)。
    - `ellipse`: `Rect`に内接する楕円。
    - `triangle`: 上辺中央の頂点+下辺の二等辺三角形(既定は二等辺。`adj1`は頂点の水平位置、
      既定0.5=中央)。
    - `rightArrow`/`leftArrow`/`upArrow`/`downArrow`: 軸方向の矢尻(幅比`adj2`、既定0.5)と
      軸に垂直な向きの軸の太さ比(`adj1`、既定0.5)から7点の矢印多角形を組む
      (ECMA-376 `ST_ShapeType`の`rightArrow`定義に準拠、他3方向は90度単位の回転で導出)。
    - `leftRightArrow`/`upDownArrow`: 両端に矢尻を持つ形状として、上記の片方向矢印の
      パス生成を両端に適用する。
    - `wedgeRectCallout`/`wedgeRoundRectCallout`/`wedgeEllipseCallout`: 本体(矩形/角丸矩形/楕円)
      に加え、`adj1`,`adj2`(既定 -0.25, 0.75。本体に対する引き出し先端の相対位置)から
      吹き出しの引き出し三角形を1つ追加する(既定値は実装時にPDFをラスタライズして
      目視確認し、左下方向に自然な引き出しになる値として選んだ。ECMA-376の一次資料への
      当たり直しはできていない)。
    - 上記いずれのプリセットも「対応済み一覧に限定する」設計(要件10.1補足)のため、
      `custGeom`(自由曲線)や一覧外の`prst`値はParsingレイヤーの時点で
      サポート外要素として弾かれ、ここには到達しない。
    - 上記の既定調整値・`a:xfrm`/`a:lin`の角度単位(60,000分の1度・時計回り)はECMA-376の
      定義に基づく想定値であり、この設計時点では一次資料への当たり直しをしていない。
      実装(タスク14.13, 14.14)時にPDFをラスタライズして目視確認したところ、
      `leftRightArrow`/`upDownArrow`(双方向矢印)に単方向矢印と同じ矢尻長さ比の既定値(0.5)を
      適用すると、両端の矢尻だけで幅を使い切り軸(シャフト)が消えて菱形に潰れることが判明した。
      双方向矢印専用の既定値(0.25)に変更して解消した(`ShapeGeometryBuilder`)。
      その他の既定値は目視確認の範囲で明らかな破綻は見られなかったが、実際の帳票で使う段になって
      Excel生成XMLとの厳密な突き合わせが必要になった場合は改めて検証する。
  - **塗りつぶし**: `Fill`が`SolidShapeFill`なら`SKPaint.Color`、
    `LinearGradientShapeFill`なら`SKShader.CreateLinearGradient`で`Rect`の対角線相当の
    2点(開始色→終了色、`AngleDegrees`をもとに`Rect`の中心から角度方向に伸ばした2点)を
    グラデーションの始点・終点とするシェーダーを`SKPaint.Shader`に設定して`SKCanvas.DrawPath`
    (`SKPaintStyle.Fill`)する。`Fill`が`null`(`noFill`)なら塗りつぶしを描画しない。
  - **枠線**: `Outline`があれば同じ`SKPath`を`SKPaintStyle.Stroke`・`StrokeWidth = WidthPt`で
    描画する。`null`なら描画しない。
  - **テキスト**: `TextLines`の各`ShapeTextLine`を、セル内テキスト描画と同じフォント解決・
    太字/斜体合成のロジック(既存の`DrawText`相当の処理を再利用)で描画する。
    座標はLayoutレイヤーが算出済みの(回転前の)ローカル座標であり、
    シェイプ本体と同じ`Save`/`RotateDegrees`/`Restore`のブロック内で描画することで
    回転が正しく反映される。
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
    PageSetupModel PageSetup,
    IReadOnlyList<DrawingObjectModel> DrawingObjects); // シート上の画像・図形(要件9, 10)。drawing.xmlの出現順(=重なり順)。

// シートに浮かぶ描画オブジェクト(画像・図形)の共通の位置決め情報。
public abstract record DrawingObjectModel(
    CellAddress AnchorCell,               // アンカー左上セル
    PointPt AnchorOffset,                 // アンカーセル左上からのオフセット(pt)
    AnchorExtent Extent);

// 画像(要件9)。ContentTypeがラスター形式の許可リスト外の場合はサポート外要素として扱う。
public sealed record ImageModel(
    byte[] Data,
    string ContentType,                   // 例: "image/png"
    CellAddress AnchorCell,
    PointPt AnchorOffset,
    AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

// 図形(要件10)。対応済みプリセット一覧に含まれないprstGeomはサポート外要素として扱う。
public sealed record ShapeModel(
    ShapePresetType Preset,
    IReadOnlyList<double> AdjustmentValues, // a:avLstのガイド値。プリセットごとに定めた順序で並ぶ。空なら既定値を使う
    double RotationDegrees,               // a:xfrm/@rotから変換。時計回り
    ShapeFill? Fill,                      // nullはnoFill(塗りつぶし無し)
    ShapeOutline? Outline,                // nullは枠線無し
    ShapeTextBody? Text,                  // nullはxdr:txBody無し
    CellAddress AnchorCell,
    PointPt AnchorOffset,
    AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

public enum ShapePresetType
{
    Rect, RoundRect, Ellipse, Triangle,
    RightArrow, LeftArrow, UpArrow, DownArrow, LeftRightArrow, UpDownArrow,
    WedgeRectCallout, WedgeRoundRectCallout, WedgeEllipseCallout,
}

public abstract record ShapeFill;
public sealed record SolidShapeFill(ArgbColor Color) : ShapeFill;
public sealed record LinearGradientShapeFill(ArgbColor StartColor, ArgbColor EndColor, double AngleDegrees) : ShapeFill;

public sealed record ShapeOutline(ArgbColor Color, double WidthPt);

// 段落・ランの構造はOOXMLのa:pPr/a:rPrにあわせる。折り返しはLayoutレイヤーが行う(未折り返しの原文)。
public sealed record ShapeTextBody(
    IReadOnlyList<ShapeTextParagraph> Paragraphs,
    VerticalAlignment VAlign);            // a:bodyPr/@anchor
public sealed record ShapeTextParagraph(
    IReadOnlyList<ShapeTextRun> Runs,
    HorizontalAlignment HAlign);          // a:pPr/@algn
public sealed record ShapeTextRun(string Text, FontStyle Font);

public abstract record AnchorExtent;

// oneCellAnchor相当: セルに対して固定サイズ(セルの拡大縮小に連動しない)。
public sealed record FixedAnchorExtent(double WidthPt, double HeightPt) : AnchorExtent;

// twoCellAnchor相当: 対角のセル+オフセットで範囲が決まる(セルの拡大縮小に連動)。
public sealed record CellSpanAnchorExtent(CellAddress ToCell, PointPt ToOffset) : AnchorExtent;

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
    IReadOnlyList<DrawCommand> Commands,  // 背景 → 罫線 → テキスト → 画像の順
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
public sealed record ImageCommand(RectPt Rect, byte[] Data, string ContentType) : DrawCommand;

// 図形(要件10)。Rect/TextLinesの座標は回転前のローカル座標。回転はRenderingレイヤーが適用する。
public sealed record ShapeCommand(
    RectPt Rect,
    ShapePresetType Preset,
    IReadOnlyList<double> AdjustmentValues,
    double RotationDegrees,
    ShapeFill? Fill,
    ShapeOutline? Outline,
    IReadOnlyList<ShapeTextLine> TextLines) : DrawCommand;

public sealed record ShapeTextLine(PointPt Origin, string Text, FontStyle Font, TextAnchor Anchor);
```

`TextCommand.Origin` の X は `Anchor`(Left/Center/Right)の基準点、Y はベースライン位置を表す。
文字列の実際の幅は描画時のフォントで決まるため、Layout は基準点だけを確定させ、
左右揃えの最終的な字送りは Rendering が行う。

## エラーハンドリング方針

- 例外階層は `UtsushiException`(`Utsushi.Core.Exceptions`)を基底とし、以下を派生させる。
  - `ReportDefinitionNotFoundException`(要件1.4)
  - `ReportStructureMismatchException`(要件1.4: シート名・セル番地の不一致)
  - `UnsupportedWorkbookElementException`(要件1.5, 9.4, 9.6, 10.7, 10.8。`ElementKind`は
    `"Drawing"`/`"Chart"`/`"LegacyDrawing"`/`"ExternalReference"`に加え、デコード不能または
    申告と実バイト列が一致しない画像形式を示す`"UnsupportedImageFormat"`、画像枚数の上限超過を
    示す`"TooManyImages"`、画像サイズの上限超過を示す`"ImageTooLarge"`、対応済み一覧に無い
    プリセットジオメトリを示す`"UnsupportedShapePreset"`、図形個数の上限超過を示す
    `"TooManyShapes"`、図形内テキストの文字数上限超過を示す`"ShapeTextTooLong"`を持つ)
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
- **改ページをまたぐ画像・図形**(要件9, 10): 画像・図形のアンカー左上セルが属するページにのみ
  全体を配置し、他のページには何も描画しない(結合セルのような「見えている部分だけ切り出す」対応は
  行わない)。ページ全体からはみ出す部分は `SKCanvas` が自然にクリップするため見た目が崩れる
  ことはないが、Excel側で2ページ目に一部かかるレイアウトを組んでいる場合、
  そのページには何も表示されない点でExcelの見た目と異なる。会社ロゴや注記の吹き出しのような
  「常に1ページの決まった位置に収まる」用途を主眼に置いた割り切りであり、
  対象帳票で実際に問題になった場合に改めて対応する。
- **画像形式の拡張**(要件9.4): 現時点でサポートするのはPNG/JPEG/GIF/BMPのみ。
  EMF/WMF(Excelがベクタ図形やクリップボード貼り付け画像を保存する際によく使う形式)は
  SkiaSharpが直接デコードできず、対応するには追加の変換ライブラリ(ライセンス確認が必要)か
  自前のパーサが要る。対象帳票で実際に必要になった時点で改めて検討する。
- **図形プリセットの拡張・接続線・グループ化**(要件10.1, 10.3, 10.7): 対応済みプリセットは
  自社帳票での実用上の必要性から選んだ13種にとどめており、星形・フローチャート記号・
  自由曲線(`custGeom`)は「サポート外要素」のままである。接続線(`xdr:cxnSp`)と
  グループ化された図形(`xdr:grpSp`)も対象外(補足10.3)。いずれも対象帳票で
  実際に必要になった時点で要件を追記して拡張する。
- **グラデーションの多段階・角度の完全再現**(要件10.6): 現時点では開始色・終了色の2点のみの
  線形グラデーションで近似しており、3点以上のグラデーションストップやExcel特有の
  グラデーション角度の細かい仕様は再現しない。対象帳票で見た目の差異が問題になった場合に
  改めて検討する。
- **OOXMLパーツ全体の非圧縮サイズに対する上限が無い**(security-reviewer指摘、要件6.1系の
  信頼できない入力に対する安全弁の一部として今後検討): `SpreadsheetDocument.Open` は
  `OpenSettings`(`MaxCharactersInPart`等)を指定せずに呼んでいるため、`drawing.xml`を含む
  各パーツ全体のDOM展開自体には上限が無い。要件9.6・10.8で設けた画像枚数・図形個数・
  図形内テキスト文字数の上限は、あくまで「DOM展開後、実際のデコード・折り返し計算等の
  重い処理へ進む前」の安全弁であり、DOM展開そのものを止める仕組みではない。画像対応時から
  存在する既存のギャップだが、図形内テキスト(`xdr:txBody`)という「XML中に際限なく
  埋め込める」経路が増えたことで実害が生じやすくなったため、対象帳票で問題になった場合は
  `OpenSettings.MaxCharactersInPart`の設定を検討する。
- **画像・図形の上限がシート単位でありワークブック単位の合算上限が無い**(security-reviewer指摘):
  `MaxImagesPerSheet`/`MaxShapesPerSheet`はシートごとにリセットされるカウンタであり、
  ワークブック全体でシートをまたいだ合算上限は無い。`ReportPdfConverter.Convert`は
  常に帳票定義の`sheetName`1枚に処理対象を絞る(`WorkbookReadOptions.SheetNameFilter`)ため
  現状の呼び出し経路では実害は無いが、`IWorkbookReader`/`OpenXmlWorkbookReader`は
  `public`であり、`SheetNameFilter`を指定しない(全シート読み取り)呼び出し方をする
  将来のコードが現れた場合はシート数倍の積み上げに対する上限が無い。対象帳票で
  実際に必要になった時点でワークブック単位の合算上限を検討する。
