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
- **入力ファイル全体の規模に対する安全弁(要件6.6)**: 画像・図形・結合セルのような個別の
  描画オブジェクトの上限とは別に、ファイル全体の規模に対しても上限を設ける。
  - `SpreadsheetDocument.Open` の前に `GuardPackageSize` で、ZIPエントリの宣言サイズ
    (中央ディレクトリの展開後サイズ)の合計を `MaxXlsxPackageBytes`(既定1GiB)と比較する
    (いわゆる「ZIP爆弾」対策)。不正なZIP構造の判定はこの時点では行わず、
    `SpreadsheetDocument.Open` 失敗時の既存の分類(`MapOpenFailure`)に委ねる。
  - `Read(Stream, ...)` がシーク不可ストリームをメモリへ複製する箇所(`EnsureSeekable`)にも
    同じ `MaxXlsxPackageBytes` を上限として設け、複製中に超過した時点で打ち切る
    (ファイルからの `FileStream` は常にシーク可能なため通常はこの経路を通らないが、
    任意の `Stream` を受け付ける `IWorkbookReader.Read` の公開APIとしての安全弁)。
  - 共有文字列(`sharedStrings.xml`)の件数に `MaxSharedStringCount`(既定20万件)、
    1シートのセル総数に `MaxCellsPerSheet`(既定50万個)の上限を設ける。
    行番号(`row/@r`)自体が上限近くを指す不正な入力は、実際のセル数によらず
    行高リストの構築(`EnsureSize`)を巨大化させるため、セルを読む前にこの時点で拒否する。
  - 上記いずれも超過時は `unsupportedElements` の設定によらず常に
    `InvalidExcelFileException`(`Reason = TooLarge`)を送出する(要件6.6補足のとおり、
    特定の要素だけをスキップして続行できる性質のものではないため)。
  - フォント・塗りつぶし・罫線・数値書式・`cellXfs`(`styles.xml`)の件数には
    `StyleTable.MaxStyleTableEntries`(既定1万件)・`ColorResolver.MaxIndexedColorCount`
    (既定1万件)で上限を設けるが、こちらは例外化せず読み取りを打ち切るのみとする。
    範囲外の`styleIndex`・色索引は既存の「既定書式/fallbackへフォールバック」という
    挙動(`StyleTable.GetCellStyle`・`ColorResolver.ResolveIndexed`)にそのまま従うため、
    値の欠落ではなく見た目の劣化にとどまり、画像・図形の個数上限をIgnoreモードで
    超過した場合と同水準の扱いになる。
- **結合セル範囲(要件2.9)**: `mergeCell` 要素を `MergedRange` として読み取る。1シートあたりの
  結合範囲の個数に上限(`MaxMergedRangesPerSheet`、既定1000)を設ける(`ElementKind =
  "TooManyMergedRanges"`。`unsupportedElements` ポリシーに従う)。Layoutレイヤーの結合セル矩形統合・
  罫線合成(`SheetModel.FindMergedRange`)は結合範囲の個数に比例する線形走査をセルごとに行うため、
  個数を無制限に許すと処理量がページ内セル数×結合範囲数で増大する(画像・図形の個数上限
  (`MaxImagesPerSheet`/`MaxShapesPerSheet`)と同じ理由によるDoS対策。security-reviewer指摘)。
  個々の結合範囲の大きさ(行数・列数)自体の上限は、Layoutレイヤー節の`MaxSpanCells`を参照。
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
  - **回転(要件9.7)**: `xdr:pic/xdr:spPr/a:xfrm/@rot` を図形と全く同じ変換
    (`ReadImage`/`ReadGroupChildImage`双方で図形の「回転」節と同じ60,000分の1度→度の
    変換式を使う)で読み取り、`ImageModel.RotationDegrees`/`GroupChildImage.RotationDegrees`
    に保持する。ロゴ画像は通常回転しないが、捺印画像(角度をつけた印影)のように
    回転させて配置する運用があるため、画像対応(要件9)の当初実装から後付けで対応した。
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
  - **既存の`DetectUnsupportedElements`/`HasUnsupportedDrawingObject`との整合(接続線・グループ対応で再拡張)**:
    シェイプ対応時に「`xdr:pic` および `xdr:sp`(シェイプ)以外が1つでもあれば `Drawing` として
    例外化」としていた構造判定(`HasUnsupportedDrawingObject`。旧名`HasNonPictureDrawingObject`)を、
    「`xdr:pic`/`xdr:sp`/`xdr:cxnSp`(接続線)/`xdr:grpSp`(グループ)のいずれでもない描画
    オブジェクト(図表枠は別途検出済み)が1つでもあれば `Drawing` として例外化」にさらに拡張する。
    `xdr:grpSp`はグループ内部を再帰的に列挙する必要はなく、グループ要素自身の種類のみで
    構造的に許容する(内部要素の再帰検証は本節ではなく`ReadGroupShape`が担う。後述)。
    この構造判定は要素の種類のみを見ており、プリセットが対応済み一覧に含まれるかどうかは
    問わない。プリセットの対応可否(シェイプ・接続線それぞれ)は画像の`ContentType`許可リスト
    判定(要件9.4)と同じ位置付けで、後段の個別読み取り(`ReadShape`/`ReadConnector`)が
    検証し、非対応プリセットは`ElementKind = "UnsupportedShapePreset"`として
    `unsupportedElements`ポリシーに従う(下記「プリセットの判定と非対応プリセットの扱い」)。
    この2段構えにより、`UnsupportedShapePreset`が「画像の`UnsupportedImageFormat`」と
    同じ経路(`DetectUnsupportedElements`を通過した後の個別検証)で意味を持つ。
    グループ内部に非対応プリセット・非対応の描画オブジェクトが1つでもある場合は、
    `ReadGroupShape`が子要素を再帰的に検証したうえでグループ全体を`UnsupportedShapePreset`と
    して扱う(要件10.7・10.10、後述「グループ(要件10.10)」節)。
  - **幾何情報**: プリセット種別に加え、`a:avLst/a:gd`(調整ガイド)の `name`/`fmla="val N"`
    を `name → N/100000.0` の辞書として読み取り、`ShapePresetType` ごとに定義した
    ガイド名の並び順(例: `rightArrow` なら `["adj1", "adj2"]`)で `IReadOnlyList<double>`
    に整形する(該当ガイドが無ければそのプリセットのECMA-376既定値を使う。既定値表は
    Renderingレイヤーに置く。詳細は後述)。パス生成そのものはRenderingレイヤーの責務であり、
    Parsingレイヤーは数値の抽出のみを行う。
  - **塗りつぶし・枠線**: `xdr:spPr/a:solidFill` は単色。`a:gradFill/a:gsLst` は
    全ての `a:gs`(位置`@pos`・色)を `GradientStop` のリストとして読み取る(要件10.6。
    3点以上のグラデーションストップに対応)。角度は子要素が `a:lin` なら線形
    (`LinearGradientShapeFill`、`@ang` を60,000分の1度から度に変換)、`a:path`
    (`@path="circle"`)なら放射状(`RadialGradientShapeFill`。`a:fillToRect` があれば
    その中心を放射の中心として読み取り、無ければ矩形中心)として`ShapeFill`の派生型を
    分岐させる。`a:path` の `@path="rect"`/`"shape"` は放射状として近似する(要件10.6補足)。
    `a:noFill` は塗りなし(`Fill = null`)。枠線は `a:ln` の `a:solidFill` の色と
    `@w`(EMU)から変換した太さを読み取り、`a:ln` 自体が無い/`a:noFill` の場合は
    `Outline = null` とする(接続線も同じ`ReadShapeOutline`を流用する)。
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
    - 1シートあたりの図形・接続線・グループ(グループ内部の子孫要素も含む)の合計個数の
      上限(既定50個。画像の上限とは独立にカウントする)。超過分は `unsupportedElements`
      の設定に従う(`ElementKind = "TooManyShapes"`。要件10.8の「合計個数」に対応するため
      画像対応時からの名称をそのまま流用し、対象範囲を図形・接続線・グループへ拡張する)。
    - 図形1つに含まれる全テキスト(段落・ランを連結した文字数)の上限(既定2000文字)。
      超過時も同様(`ElementKind = "ShapeTextTooLong"`)。文字数を実際に折り返し計算へ
      渡す前に拒否することで、極端に長い文字列に対する折り返し計算量を避ける。
    - グループのネスト段数の上限(既定5段)。超過した時点でそのグループ全体を拒否する
      (`ElementKind = "GroupNestingTooDeep"`)。再帰的な子グループの展開によるスタック消費・
      計算量の増大を防ぐ。
    - 幅・高さの計算に使う列/行数の上限(既定4096)は画像と共通の
      `SpanWidthPt`/`SpanHeightPt` をそのまま流用するため、別途の上限追加は不要。

- **接続線(要件10.9)**: `xdr:cxnSp`(`xdr:sp`とは別の要素)のうち、
  `xdr:spPr/a:prstGeom/@prst` が `straightConnector1`/`bentConnector2`/`bentConnector3`/
  `curvedConnector2`/`curvedConnector3` のいずれかであるものを `ConnectorModel` として
  読み取る。アンカー(セル位置・範囲)・回転・枠線は図形と同じ仕組み(`ReadAnchorExtent`/
  `ReadShapeOutline`)を再利用する。接続線は塗りつぶし・テキストを持たないため
  `ShapeModel` を流用せず専用の型とする。反転(`a:xfrm/@flipH`, `@flipV`)は接続線の
  経路(どちら向きに折れる/曲がるか)を決めるため新たに読み取る(`FlipHorizontal`/
  `FlipVertical`)。パス生成(実際にどう折れ線・曲線を引くか)はRenderingレイヤーの責務。
  非対応プリセットの接続線(`bentConnector4`/`5`, `curvedConnector4`/`5` 等)は
  `ElementKind = "UnsupportedShapePreset"` として扱う(図形と同じ判定ロジックを流用)。
  接続線の個数は図形と合算した`MaxShapesPerSheet`でカウントする。
  - **接続点の参照(要件10.11)**: `xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvCxnSpPr`配下の
    `a:stCxn`/`a:endCxn`(いずれも`@id`+`@idx`の属性を持つ空要素。無くてもよい)を、
    `ConnectorModel.StartConnection`/`EndConnection`
    (`ConnectionRef(uint ShapeId, uint SiteIndex)?`。要素が無ければ`null`)として
    読み取る。接続点の実際の解決(`ShapeId`から参照先の矩形を引き、`SiteIndex`から
    矩形上の座標を求める)はLayoutレイヤーの責務とする(Parsingの時点では他の描画
    オブジェクトの最終ページ座標がまだ確定していない。ページ分割・グループ展開は
    Layoutが行うため)。グループ内の`GroupChildConnector`も同じ形で
    `StartConnection`/`EndConnection`を持つ。

- **接続点解決のための図形ID読み取り(要件10.11)**: `stCxn`/`endCxn`が参照する`id`は、
  参照先の`xdr:sp`/`xdr:pic`/`xdr:grpSp`(トップレベル・グループ内問わず)が持つ
  `NonVisualDrawingProperties/@id`と同じ値である。このため`ShapeModel`/`ImageModel`/
  `GroupShapeModel`、およびグループ内の`GroupChildShape`/`GroupChildImage`/
  `GroupChildGroup`(`GroupChildConnector`は接続先として参照される対象ではないため
  不要)に`Id: uint`を追加し、読み取り時にそのまま保持する。IDの妥当性(参照先の存在確認、
  同一ページ上にあるか)はLayoutレイヤーでの解決時に判定する。接続先が見つからない場合は
  例外にはせず、要件10.11の既定動作にフォールバックする(接続点解決は見た目向上のための
  機能であり、変換の可否を左右しないため)。IDが重複していた場合は例外にはせず、
  `BuildConnectionTargetTable`(id→矩形テーブル)が出現順で後から見つかった方の図形の
  座標を採用する(`@id`はOOXMLスキーマ上必須かつExcelが重複させないため通常は起こらない。
  code-reviewer指摘によりこの割り切りを明記)。

- **グループ化された図形(要件10.10)**: `xdr:grpSp` を `GroupShapeModel` として読み取る。
  グループの`grpSpPr/a:xfrm`(`TransformGroup`)から、グループ自身の回転
  (`RotationDegrees`)と、子座標空間の原点・大きさ(`a:chOff`→`ChildOffset`、
  `a:chExt`→`ChildExtent`。EMUからポイントへ変換)を読み取る。
  グループの直接の子要素(`xdr:sp`/`xdr:pic`/`xdr:cxnSp`/入れ子の`xdr:grpSp`)を
  出現順に列挙し、`GroupChildModel`(`GroupChildShape`/`GroupChildImage`/
  `GroupChildConnector`/`GroupChildGroup`)へ変換する。子要素はセルアンカー
  (`xdr:from`/`xdr:to`)を持たず、`a:xfrm/a:off`・`a:ext`(グループの子座標空間上の
  絶対位置)で位置が決まる点が、シート直下の描画オブジェクトと異なる
  (子要素の位置・サイズを最終的なページ座標へ変換する計算はLayoutレイヤーの責務。
  「Layout レイヤー」の節を参照)。
  - グループ内の要素のいずれか1つでも対応済みプリセット一覧に含まれない場合
    (非対応プリセットの図形・接続線・`xdr:graphicFrame`等)は、要件10.10により
    グループ全体を`ElementKind = "UnsupportedShapePreset"`のサポート外要素として扱う
    (グループの一部だけを描画すると、Excel上の見た目・要素間の位置関係が意図せず崩れるため)。
  - 入れ子のグループは同じ規則を再帰的に適用し、`MaxShapeNestingDepth`(既定5)を
    超えた時点で全体を`ElementKind = "GroupNestingTooDeep"`として拒否する。
  - グループ・接続線を含めたシート全体の描画オブジェクト総数(トップレベル+グループ内の
    全子孫)が`MaxShapesPerSheet`を超える場合も同様に拒否する(要件10.8)。

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
    画像・図形1つの表示サイズ(`MaxDrawingObjectDimensionPt`、既定5000pt)の上限は、印刷拡大率
    (`_scale`)適用後の最終的なページ座標上の矩形サイズに対して適用する(トップレベル
    (`TryComputeDrawingObjectRect`)・グループ内子要素(`ToGroupChildRect`)のいずれも同じ適用点
    に揃えている。以前はトップレベルのみ`_scale`適用前の論理サイズに適用しており、印刷拡大率が
    100%を超える帳票で上限の実効値がグループ内子要素と食い違っていた〈layout-fidelity-reviewer指摘〉)。
    2セルアンカーの合算列/行数の上限(`MaxSpanCells`)と同じ値・同じ考え方を、結合セルの外周罫線
    合成(`ResolveColumnEdge`/`ResolveRowEdge`。範囲内の各セルを走査して可視な罫線を探す)、および
    結合範囲のうちページ上に見えている部分の判定(`FindVisibleSpan`)にも適用する。`mergeCell`要素の
    範囲サイズはParsingレイヤー(`OpenXmlWorkbookReader.ReadMergedRanges`)で上限を設けていないため、
    極端に大きい結合範囲(対角セルにセル番地の上限近くを指定するなど)を持つ入力に対する走査量の
    増大を、2セルアンカー画像と同じ理由で防ぐ〈security-reviewer/code-reviewer指摘〉。
  - **図形内テキストの折り返し・配置(要件10.4)**: `ShapeModel.Text`(段落・ランの木構造)を、
    セル内テキストの折り返しと同じ`IFontMetricsProvider`を使い、図形の矩形幅を基準に
    単純な幅基準の折り返し(禁則処理なし。セル内テキストの折り返しと同水準)で複数行に
    分割する。各行の水平位置は段落の`HAlign`、行全体の垂直位置は`VAlign`と行数から
    (セル内テキストの上下揃えと同じ考え方で)算出し、`ShapeCommand.TextLines`の
    各`ShapeTextLine`として矩形内のポイント座標(回転前、シェイプ自身のローカル座標)を
    確定させる。回転の適用はRenderingレイヤーの責務とする(座標変換をLayoutに持ち込むと
    `PagedLayout`が回転行列という新しい概念を持つことになり、既存の「軸に平行な矩形の
    集まり」という単純なモデルから外れるため)。
  - **接続線の配置(要件10.9)**: `ConnectorModel`は`DrawingObjectModel`のため、画像・図形と
    全く同じ`TryComputeDrawingObjectRect`でページ矩形を求め、`ConnectorCommand`を生成する。
    経路(実際にどう折れ線・曲線を引くか)はRenderingレイヤーの責務。
  - **グループの展開(要件10.10)**: `GroupShapeModel`もまず`TryComputeDrawingObjectRect`で
    グループ自身のページ矩形(`groupRect`)を求める。そのうえで、グループの子座標空間
    (`ChildOffset`/`ChildExtent`)から`groupRect`への比例変換
    (`scaleX = groupRect.Width / ChildExtent.X`、`scaleY = groupRect.Height / ChildExtent.Y`。
    非一様倍率を許容する。Excel自体もグループのリサイズで子要素が縦横別倍率で伸縮しうるため)
    を使って、各子要素の`LocalRect`(子座標空間上の位置・サイズ)を
    `pageRect = groupRect.TopLeft + (LocalRect.TopLeft - ChildOffset) * (scaleX, scaleY)`
    でページ座標へ変換し、`ShapeCommand`/`ImageCommand`/`ConnectorCommand`を生成する
    (`GroupChildShape`のテキスト折り返しも通常の図形と同じロジックを流用するが、
    矩形自体が`scaleX`/`scaleY`で縮小/拡大されているのに合わせて、内側余白
    (`ShapeTextPaddingPt`)・フォントサイズにも`scaleX`と`scaleY`の幾何平均を追加の係数として
    掛ける。トップレベルの図形はこの係数が1.0になるため、印刷拡大率(`_scale`)のみが
    効く従来どおりの挙動のままである。グループが大きく縮小されている場合に余白が
    シェイプ本体ほど縮まらずテキストが矩形からはみ出す/消えることを防ぐための対応
    〈layout-fidelity-reviewer指摘〉)。
    入れ子の`GroupChildGroup`は、自身の`pageRect`を新たな`groupRect`として同じ変換を
    再帰的に適用する。
    こうして生成した子要素の`DrawCommand`列を、`GroupCommand(Center, RotationDegrees, Children)`
    (`Center = groupRect`の中心、`RotationDegrees = `グループ自身の回転角)でまとめて包み、
    1つの`DrawCommand`として`Commands`リストに追加する。グループの回転を子要素の座標
    そのものに焼き込まず`GroupCommand`という薄いラッパーに持たせるのは、グループの回転が
    「グループ全体を剛体として1回だけ回す」座標変換であり、子要素自身の個別の回転
    (`ShapeModel.RotationDegrees`等)とは独立に合成される(Renderingレイヤーで
    `canvas`の回転変換を入れ子にすることで自然に合成できる)ためである。
    詳細はRenderingレイヤーの節を参照。
  - **接続点(コネクションサイト)の解決(要件10.11)**: `EmitDrawingObjects`の出現順を保った
    単一の`foreach`(既存構造)は変更せず、その**前**に軽量な予備パスを1回追加する。
    この予備パス(`BuildConnectionTargetTable`)は`_sheet.DrawingObjects`を走査し、
    接続線(`ConnectorModel`)を除く各要素について既存の`TryComputeDrawingObjectRect`/
    `BuildGroupChildren`と同じ計算を行い、`Id`
    (トップレベルは`ImageModel`/`ShapeModel`/`GroupShapeModel`自身、グループ内は
    `GroupChildShape`/`GroupChildImage`/`GroupChildGroup`)をキーとする
    `Dictionary<uint, (RectPt Rect, ShapePresetType? Preset)>`(このページに限定した
    解決テーブル。画像・グループ自身は`Preset = null`)を組み立てるだけで、
    `DrawCommand`は生成しない(コマンド生成は本来の`foreach`が担当する)。
    グループ内要素は`ToGroupChildRect`変換後の最終ページ矩形を記録する。
    矩形計算をこの予備パスと本来の`foreach`の2回行うことになるが、単純な算術のみで
    コストは無視できる。この設計により、出現順を保ったコマンド生成ロジック自体には
    一切手を入れず(z-orderの保持は既存のとおり自明)、接続線が自分より後に出現する
    図形を参照していても解決できる(予備パスの時点で全IDが判明済みのため)。
    本来の`foreach`が接続線(トップレベルの`ConnectorModel`、グループ内の
    `GroupChildConnector`)に到達した際、`StartConnection`/`EndConnection`の`ShapeId`が
    このテーブルに存在すれば、`ConnectionSiteResolver.Resolve(rect, preset, siteIndex)`で
    実際の座標(`PointPt`)を求め、`ConnectorCommand.ResolvedStart`/`ResolvedEnd`
    (`PointPt?`)に格納する。存在しない場合(参照先がこのページに無い、IDが実在しない、
    `StartConnection`/`EndConnection`が`null`)は`null`のままとし、Renderingレイヤーが
    要件10.9の既定動作(アンカー矩形+反転)にフォールバックする。
    `ConnectionSiteResolver.Resolve`は既定では矩形の上下左右の中点
    (`siteIndex % 4`で0〜3の範囲に丸める。0=上,1=左,2=下,3=右。ECMA-376で`cxnLst`を
    持たない図形の既定の接続点と同じ考え方)を返す。`preset`が`FlowChartInputOutput`
    (平行四辺形)の場合のみ、左右の接続点(idx 1, 3)を実際の傾いた辺の中点
    (`InputOutputSkewRatio`と同じ比率で`x`座標を、辺の傾きに応じて内側に補正)に
    調整する(上下の接続点は上下の辺がもともと水平なため補正不要)。`FlowChartDocument`
    (波形)は、波形の谷の最も深い点が設計上ちょうど配置矩形の下辺中点と一致するため、
    補正は不要で既定の4方向をそのまま使う。それ以外のプリセット・画像・グループ
    (`preset = null`)も既定の4方向をそのまま使う。
    この解決テーブルはページごとに作り直す(接続線と参照先が異なるページに分かれる場合は
    解決できない。改ページをまたぐ画像・図形の既存の割り切りと同じ理由)。
    - **レイヤー依存の一方向ルールに関する注意(重要)**: `InputOutputSkewRatio`は、
      これまで`Utsushi.Rendering`の`ShapeGeometryBuilder`にのみ存在する`private`定数
      だった(`flowChartInputOutput`の実際の描画パス生成に使う)。`ConnectionSiteResolver`
      はLayoutレイヤーに置くため、Rendering内部の定数をそのまま参照することはできない
      (`Utsushi.Rendering`は`Utsushi.Layout`に依存する向きであり、逆方向の参照は
      循環参照になりビルドできない。`.kiro/steering/structure.md`のレイヤー依存の
      一方向ルールにも反する。code-reviewer相当の指摘により設計時に発見)。そのため、
      この定数を`Utsushi.Parsing.Model`(`ShapePresetType`と同じ場所。LayoutもRendering
      も既にこのレイヤーに依存しているため、双方から参照できる)の
      `ShapeGeometryConstants`という新しい`public static class`へ移動し、
      `ShapeGeometryBuilder`側は移動後の定数をそのまま使うよう参照を書き換える
      (値・意味は変えない。定数の置き場所を変えるだけ)。`DocumentWaveDepthRatio`
      (`flowChartDocument`の波形の深さ比率)はLayoutレイヤーから参照する必要が無い
      ため、`ShapeGeometryBuilder`の`private`定数のまま変更しない。
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
- 1ページ = 1 `SKCanvas` への描画。矩形塗りつぶし(背景)→罫線→テキスト→
  **画像・図形・接続線・グループ**(`drawing.xml`の出現順)の順で描画する。Excelはシート上に
  浮かぶ描画オブジェクト(画像・図形等)をセルの内容より上のレイヤーとして描画するため、
  これらは他のセル内容と重なる場合に最前面へ来るようにする(要件9.3, 10.3)。
- **画像の描画(要件9)**: `ImageCommand` は `SKBitmap.Decode(byte[])`でデコードし、
  `SKCanvas.DrawBitmap(bitmap, destRect)` で `ImageCommand.Rect` へ描画する
  (SkiaSharp 2.88.8で利用可能な標準API)。既存の `ToSkRect(RectPt)` をそのまま使う。
  ページ境界外にはみ出す部分は `SKCanvas` が自然にクリップするため、追加のクリップ処理は不要。
  `SKBitmap.Decode` で実際に展開する前に `SKBitmap.DecodeBounds` で宣言上のピクセル寸法を確認し、
  上限(既定4096px)を超える場合はデコードせず `PdfRenderingException` とする(要件9.6。
  ピクセル爆弾対策。詳細はParsingレイヤー節「信頼できない入力に対する安全弁」を参照)。
  回転がある場合は、図形(要件10.5)と全く同じ`canvas.Save()` →
  `canvas.RotateDegrees(RotationDegrees, centerX, centerY)`(中心は`Rect`の中心)→
  `DrawBitmap` → `canvas.Restore()` のパターンで`RotationDegrees`を反映する(要件9.7。
  捺印画像のように回転させて配置する運用があるため対応する)。
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
    - `cloudCallout`(雲形吹き出し、要件10.13): 楕円本体の輪郭を、中心から一定間隔で
      並べた円弧(バンプ)の和集合(`SKPath.Op(SKPathOp.Union)`)で近似した「雲」の
      シルエットに、`wedgeEllipseCallout`と同じ引き出し三角形を追加する。バンプの
      個数・半径は固定値(見た目のバランスを目視確認して決定)のままとし調整ガイドには
      対応しないが、引き出し三角形の先端位置は`wedgeRectCallout`等と同様に`adj1`/`adj2`
      (ファイルに無ければ既定値 -0.25, 0.75)から読み取る(`ShapeAdjustmentGuideNames`に
      `["adj1", "adj2"]`を追加し、`CloudCalloutPath`のシグネチャに`adjustmentValues`を
      追加する)。ECMA-376上も雲形の輪郭自体(バンプ)は固定のパスであり、
      調整ガイドは引き出し位置のみに影響するため、この変更は輪郭の近似精度には影響しない。
    - `callout1`/`callout2`/`callout3`(引き出し線付き吹き出し): 本体は`rect`、
      そこから矩形の外側の1点(既定は左下方向)へ向けて1〜3本の線分からなる
      折れ線(引き出し線)を追加する汎用の「N本引き出し線」生成ロジックとして実装する
      (N=1,2,3をパラメータ化。三角形の塗りつぶしではなく線のみの点でwedge系と異なる)。
    - `star4`/`star5`/`star6`/`star8`(星形、要件10.12): 外接円の半径`R`(=`min(幅,高さ)/2`)と、
      内側の頂点の半径比(ECMA-376既定の調整ガイド名は単一の`adj`)から、外側の頂点と
      内側の頂点を交互に結ぶ`2 * N`角形を組む(`N`=4/5/6/8)。`adj`の既定値(ファイルに
      `a:avLst`の指定が無い場合)はプリセットごとに異なり、`既定値 ÷ 50000`が実際の
      半径比になる: `star4`=12500(0.25)、`star5`=19098(0.382)、`star6`=28868(0.577)、
      `star8`=37500(0.75)(`DefaultStarInnerRadiusRatio`を`ShapePresetType`ごとの
      定数に分割する)。この既定値は二次資料(ECMA-376の実装を参照する複数のOSS
      プロジェクトの記述)を突き合わせて確認したものであり、ECMA-376一次資料そのものへの
      当たり直しはできていない。頂点の回転オフセットは`star4`/`star5`/`star6`/`star8`
      いずれも最初の外側の頂点を真上(-90度)に置く同一の規則を使う(単一の調整ガイドで
      頂点を交互に結ぶ一般的な星形の描画方式は全プリセット共通であるため、既存の実装
      (`StarPath`)のとおりで変更不要。プリセットごとに異なる既定角度を使う根拠は
      ECMA-376上に見当たらない)。
    - `flowChartProcess`(処理): `rect`と同じ矩形。
    - `flowChartDecision`(判断): 矩形の上下左右の中点を結んだ菱形。
    - `flowChartTerminator`(端子): 左右端を半円にした「スタジアム」形状
      (`roundRect`の角丸半径を`高さ/2`に固定した特殊形として実装できる)。
    - `flowChartInputOutput`(入出力): 上下の辺を左右にずらした平行四辺形。ずらし幅の比率
      (`InputOutputSkewRatio`)は要件10.11の接続点解決(Layoutレイヤー、左右の接続点の
      補正)とも共有するため、`Utsushi.Parsing.Model.ShapeGeometryConstants`に定義する
      (下記「レイヤー依存の一方向ルールに関する注意」参照)。
    - `flowChartDocument`(書類): 矩形の下辺を波形(1つの緩やかな凹み)にした形状。波形の
      深さの比率(`DocumentWaveDepthRatio`)はLayoutレイヤーと共有する必要が無いため、
      `Utsushi.Rendering`内部の`private`定数のまま変更しない。
    - `flowChartPredefinedProcess`(定義済み処理): `rect`に加え、左右の辺の内側に
      それぞれ縦線を1本ずつ追加する。
    - `flowChartConnector`(結合子): `ellipse`と同じ楕円(正円になるようExcel側で
      正方形のバウンディングボックスにするのが通常)。
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
  - **塗りつぶし**: `Fill`が`SolidShapeFill`なら`SKPaint.Color`。
    `LinearGradientShapeFill`(`Stops: IReadOnlyList<GradientStop>`、2点以上)なら
    `SKShader.CreateLinearGradient`で`Rect`の中心から`AngleDegrees`方向に対角線の半分
    ぶん伸ばした2点を始点・終点とし、各`GradientStop.Position`(0.0〜1.0)と色を
    `colors`/`colorPos`配列にそのまま渡す(3点以上のストップを正確に反映)。
    `RadialGradientShapeFill`(`Stops`、`CenterFraction: PointPt`)なら
    `SKShader.CreateRadialGradient`で、中心を`Rect.Left + Rect.Width * CenterFraction.X`,
    `Rect.Top + Rect.Height * CenterFraction.Y`、半径を`Rect`の対角線の半分とする
    (`a:fillToRect`が矩形中心以外を指す場合、中心がずれた放射状グラデーションになる)。
    いずれもシェーダーを`SKPaint.Shader`に設定して`SKCanvas.DrawPath`(`SKPaintStyle.Fill`)
    する。`Fill`が`null`(`noFill`)なら塗りつぶしを描画しない。
  - **枠線**: `Outline`があれば同じ`SKPath`を`SKPaintStyle.Stroke`・`StrokeWidth = WidthPt`で
    描画する。`null`なら描画しない。
  - **テキスト**: `TextLines`の各`ShapeTextLine`を、セル内テキスト描画と同じフォント解決・
    太字/斜体合成のロジック(既存の`DrawText`相当の処理を再利用)で描画する。
    座標はLayoutレイヤーが算出済みの(回転前の)ローカル座標であり、
    シェイプ本体と同じ`Save`/`RotateDegrees`/`Restore`のブロック内で描画することで
    回転が正しく反映される。
- **接続線の描画(要件10.9, 10.11)**: `ConnectorCommand`ごとに、`FlipHorizontal`/`FlipVertical`を
  反映した向きで経路(`SKPath`、塗りつぶし無しの開いたパス)を`ConnectorGeometryBuilder`
  (`ShapeGeometryBuilder`とは別に新設。接続線は塗りつぶし・調整ガイドを持たず責務が
  異なるため)で組み立てる。`ConnectorCommand.ResolvedStart`/`ResolvedEnd`(`PointPt?`。
  Layoutレイヤーが要件10.11の接続点解決に成功した場合のみ値を持つ)が両方とも非`null`の
  場合、`ConnectorGeometryBuilder.Build`はこの2点をそのまま始点・終点として使う
  (`Rect`と`FlipHorizontal`/`FlipVertical`は無視する)。片方または両方が`null`の場合は
  従来どおり`Rect`の対角(反転に応じた2頂点)を始点・終点とする。始点・終点が決まった後の
  折れ線・曲線の組み立て方(`bentConnector2`等)はどちらの経路でも共通のロジックを使う。
  - `straightConnector1`: 矩形の対角(反転に応じた2頂点)を結ぶ直線。
  - `bentConnector2`: 中間点1つで直角に折れる2辺(水平→垂直、または反転により
    垂直→水平)。
  - `bentConnector3`: 中間点2つで直角に2回折れる3辺(既定は中央で折り返す)。
  - `curvedConnector2`: 2頂点を結ぶ1本の2次ベジェ曲線(制御点は`bentConnector2`と
    同じ折れ点)。
  - `curvedConnector3`: `bentConnector3`の折れ点を通る2本のベジェ曲線によるS字カーブ。
  - `Outline`があれば`SKPaintStyle.Stroke`で描画する(`null`の場合、Excel上は既定の
    黒い実線1ptで表示されるため、`Outline`が無い接続線にも既定の線色・太さを補う)。
    `ResolvedStart`/`ResolvedEnd`が両方とも`null`(未解決、`Rect`基準の経路)の場合のみ、
    回転(`RotationDegrees`)があれば図形と同じ`Save`/`RotateDegrees`/`Restore`を使う。
    解決済みの場合は絶対座標の両端点をそのまま結んだ経路が最終的な見た目であり、
    `Rect`中心を軸にした追加の回転はかえって位置をずらすため適用しない
    (下記「未決事項」参照)。
- **グループの描画(要件10.10)**: `GroupCommand`ごとに、`canvas.Save()` →
  `canvas.RotateDegrees(RotationDegrees, Center.X, Center.Y)` → `Children`の各
  `DrawCommand`を(`FillRectCommand`等を除く、`ShapeCommand`/`ImageCommand`/
  `ConnectorCommand`/入れ子の`GroupCommand`を対象に)既存のコマンド振り分けロジックを
  再帰的に呼び出して描画 → `canvas.Restore()`。`DrawPage`内のコマンド振り分け
  (`switch`文)を`DrawSingleCommand(SKCanvas, DrawCommand, reportCode, sheetName)`という
  1コマンド分の描画ヘルパーへ切り出す(抽象レコード型`DrawCommand`と紛らわしくなるため、
  メソッド名は型名とは別の`DrawSingleCommand`とする)。`DrawPage`の`foreach`と`GroupCommand`の内部
  展開の両方から呼べるようにする。子要素自身の回転(`ShapeCommand.RotationDegrees`等)は、
  この`canvas`変換がすでに適用された座標系の内側でさらに`Save`/`RotateDegrees`/`Restore`
  するため、グループの回転と子要素個別の回転が正しく合成される(`canvas`の変換行列の
  スタックに任せることで、Renderingレイヤー側で回転の合成を数式的に計算する必要が無い)。
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
    IReadOnlyList<DrawingObjectModel> DrawingObjects); // シート上の画像・図形・接続線・グループ(要件9, 10)。drawing.xmlの出現順(=重なり順)。

// シートに浮かぶ描画オブジェクト(画像・図形)の共通の位置決め情報。
public abstract record DrawingObjectModel(
    CellAddress AnchorCell,               // アンカー左上セル
    PointPt AnchorOffset,                 // アンカーセル左上からのオフセット(pt)
    AnchorExtent Extent);

// 画像(要件9)。ContentTypeがラスター形式の許可リスト外の場合はサポート外要素として扱う。
// Id(NonVisualDrawingProperties/@id)は接続線の接続先解決(要件10.11)のために保持する。
public sealed record ImageModel(
    uint Id,
    byte[] Data,
    string ContentType,                   // 例: "image/png"
    double RotationDegrees,               // a:xfrm/@rot(60,000分の1度)を度に変換(要件9.7)
    CellAddress AnchorCell,
    PointPt AnchorOffset,
    AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

// 図形(要件10)。対応済みプリセット一覧に含まれないprstGeomはサポート外要素として扱う。
// Idはimageと同じ理由(要件10.11)で保持する。
public sealed record ShapeModel(
    uint Id,
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
    CloudCallout, Callout1, Callout2, Callout3,             // 追加(拡張フェーズ)
    Star4, Star5, Star6, Star8,                             // 追加(拡張フェーズ)
    FlowChartProcess, FlowChartDecision, FlowChartTerminator, // 追加(拡張フェーズ)
    FlowChartInputOutput, FlowChartDocument,                  // 追加(拡張フェーズ)
    FlowChartPredefinedProcess, FlowChartConnector,           // 追加(拡張フェーズ)
}

// Layout(接続点解決、左右の接続点の補正)・Rendering(実際の描画)の両方が使う比率定数。
// Utsushi.RenderingはUtsushi.Layoutに依存する向きのため、逆方向の参照を避けるべく
// 双方が依存するUtsushi.Parsing.Modelに置く(要件10.11、レイヤー依存の一方向ルール)。
// DocumentWaveDepthRatio(flowChartDocumentの波形の深さ比率)はLayoutから参照する
// 必要が無いため、ここへは移動せずUtsushi.Rendering内部のprivate定数のまま残す。
public static class ShapeGeometryConstants
{
    public const double InputOutputSkewRatio = 0.2; // flowChartInputOutputの上下辺のずらし幅比率
}

public abstract record ShapeFill;
public sealed record SolidShapeFill(ArgbColor Color) : ShapeFill;
// 3点以上のグラデーションストップに対応(拡張フェーズ)。Positionは0.0〜1.0。
public sealed record GradientStop(double Position, ArgbColor Color);
public sealed record LinearGradientShapeFill(IReadOnlyList<GradientStop> Stops, double AngleDegrees) : ShapeFill;
// 放射状グラデーション(拡張フェーズ)。CenterFractionはRectに対する中心位置の割合(既定0.5,0.5)。
public sealed record RadialGradientShapeFill(IReadOnlyList<GradientStop> Stops, PointPt CenterFraction) : ShapeFill;

public sealed record ShapeOutline(ArgbColor Color, double WidthPt);

// 段落・ランの構造はOOXMLのa:pPr/a:rPrにあわせる。折り返しはLayoutレイヤーが行う(未折り返しの原文)。
public sealed record ShapeTextBody(
    IReadOnlyList<ShapeTextParagraph> Paragraphs,
    VerticalAlignment VAlign);            // a:bodyPr/@anchor
public sealed record ShapeTextParagraph(
    IReadOnlyList<ShapeTextRun> Runs,
    HorizontalAlignment HAlign);          // a:pPr/@algn
public sealed record ShapeTextRun(string Text, FontStyle Font);

// 接続線(要件10.9)。塗りつぶし・テキストを持たない。
// StartConnection/EndConnection(要件10.11)は、a:stCxn/a:endCxnがあれば読み取る。
// 解決(参照先の矩形取得・接続点計算)はLayoutレイヤーの責務。
public sealed record ConnectorModel(
    ConnectorPresetType Preset,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical,
    ShapeOutline? Outline,
    ConnectionRef? StartConnection,
    ConnectionRef? EndConnection,
    CellAddress AnchorCell,
    PointPt AnchorOffset,
    AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

public enum ConnectorPresetType { Straight, Bent2Segment, Bent3Segment, Curved2Segment, Curved3Segment }

// a:stCxn/a:endCxnの@id(参照先のNonVisualDrawingProperties/@id)と@idx(接続点番号)。
public sealed record ConnectionRef(uint ShapeId, uint SiteIndex);

// グループ化された図形(要件10.10)。トップレベルの描画オブジェクトとしてはセルアンカーを持つが、
// 内部の子要素(Children)は独自の子座標空間(ChildOffset/ChildExtent)上の位置で決まる。
// Idはimage/shapeと同じ理由(要件10.11)で保持する。
public sealed record GroupShapeModel(
    uint Id,
    PointPt ChildOffset,                  // a:chOff(pt換算)。子要素の座標系の原点
    PointPt ChildExtent,                  // a:chExt(pt換算)。X=幅, Y=高さ
    IReadOnlyList<GroupChildModel> Children,
    double RotationDegrees,
    CellAddress AnchorCell,
    PointPt AnchorOffset,
    AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

// グループ内の子要素の位置(LocalRect)は、グループのChildOffset/ChildExtent上の座標(pt)であり、
// ページ座標への変換(平行移動+拡大縮小)はLayoutレイヤーが行う。
public abstract record GroupChildModel(RectPt LocalRect);
public sealed record GroupChildShape(
    uint Id, RectPt LocalRect, ShapePresetType Preset, IReadOnlyList<double> AdjustmentValues,
    double RotationDegrees, ShapeFill? Fill, ShapeOutline? Outline, ShapeTextBody? Text) : GroupChildModel(LocalRect);
public sealed record GroupChildImage(uint Id, RectPt LocalRect, byte[] Data, string ContentType) : GroupChildModel(LocalRect);
public sealed record GroupChildConnector(
    RectPt LocalRect, ConnectorPresetType Preset, double RotationDegrees,
    bool FlipHorizontal, bool FlipVertical, ShapeOutline? Outline,
    ConnectionRef? StartConnection, ConnectionRef? EndConnection) : GroupChildModel(LocalRect);
public sealed record GroupChildGroup(
    uint Id, RectPt LocalRect, double RotationDegrees, PointPt ChildOffset, PointPt ChildExtent,
    IReadOnlyList<GroupChildModel> Children) : GroupChildModel(LocalRect);

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
    IReadOnlyList<DrawCommand> Commands,  // 背景 → 罫線 → テキスト → 画像/図形/接続線/グループの順
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
// RotationDegreesはRectの中心を軸とした回転角(度、時計回り。要件9.7)。
public sealed record ImageCommand(RectPt Rect, byte[] Data, string ContentType, double RotationDegrees) : DrawCommand;

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

// 接続線(要件10.9, 10.11)。塗りつぶし・テキストを持たない。
// ResolvedStart/ResolvedEndは要件10.11の接続点解決に成功した場合のみ非null。
// 両方とも非nullならRect/FlipHorizontal/FlipVerticalの代わりにこの2点を始点・終点とする。
public sealed record ConnectorCommand(
    RectPt Rect,
    ConnectorPresetType Preset,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical,
    ShapeOutline? Outline,
    PointPt? ResolvedStart,
    PointPt? ResolvedEnd) : DrawCommand;

// グループ(要件10.10)。Childrenはすでにページ座標へ変換済み(グループ自身の回転は未適用)。
// Renderingレイヤーがcanvasの回転変換でChildrenをまとめて囲むことで、グループの回転と
// 子要素個別の回転を合成する。
public sealed record GroupCommand(
    PointPt Center,
    double RotationDegrees,
    IReadOnlyList<DrawCommand> Children) : DrawCommand;
```

`TextCommand.Origin` の X は `Anchor`(Left/Center/Right)の基準点、Y はベースライン位置を表す。
文字列の実際の幅は描画時のフォントで決まるため、Layout は基準点だけを確定させ、
左右揃えの最終的な字送りは Rendering が行う。

## エラーハンドリング方針

- 例外階層は `UtsushiException`(`Utsushi.Core.Exceptions`)を基底とし、以下を派生させる。
  - `ReportDefinitionNotFoundException`(要件1.4)
  - `ReportStructureMismatchException`(要件1.4: シート名・セル番地の不一致)
  - `UnsupportedWorkbookElementException`(要件1.5, 9.4, 9.6, 10.7, 10.8, 10.9, 10.10。
    `ElementKind`は`"Drawing"`/`"Chart"`/`"LegacyDrawing"`/`"ExternalReference"`に加え、
    デコード不能または申告と実バイト列が一致しない画像形式を示す`"UnsupportedImageFormat"`、
    画像枚数の上限超過を示す`"TooManyImages"`、画像サイズの上限超過を示す`"ImageTooLarge"`、
    対応済み一覧に無いプリセットジオメトリ(図形・接続線どちらの場合も共通)を示す
    `"UnsupportedShapePreset"`、図形・接続線・グループの合計個数の上限超過を示す
    `"TooManyShapes"`、図形内テキストの文字数上限超過を示す`"ShapeTextTooLong"`、
    グループのネスト段数の上限超過を示す`"GroupNestingTooDeep"`、
    結合セル範囲の個数上限超過を示す`"TooManyMergedRanges"`(要件2.9)を持つ)
  - `SubstitutionKeyNotFoundException` / `RequiredSubstitutionValueMissingException`(要件2.3, 2.4)
  - `InvalidCellOverrideAddressException`(要件2.8。セル番地直接指定がA1形式として解釈できない場合)
  - `NonAnchorMergedCellOverrideException`(要件2.9。セル番地直接指定の対象が結合セル範囲の非アンカー位置の場合)
  - `InvalidExcelFileException`(要件6.1, 6.2, 6.5, 6.6。`Reason` で非xlsx/破損/パスワード保護/
    ファイルを開けない(存在しない・アクセス不可)/ワークシートが無い/ファイル・共有文字列・
    セル数が上限超過(`TooLarge`)を区別する)
  - `ReportDefinitionSchemaException`(要件6.3。問題のあったプロパティパスを保持する)
  - `LayoutComputationException` / `PdfRenderingException` / `FontNotAvailableException`
- すべての例外は、帳票コード・シート名・セル番地・処理段階(Parsing/Substitution/Layout/Rendering)を構造化プロパティとして保持し、ログ出力時に特定できるようにする(要件6.4)。
- Renderingレイヤーは一時ファイル/一時ストリームに書き込み、正常終了時のみ最終出力先へ確定させる(要件5.4: 不完全PDFを残さない)。

### 信頼できない入力に対する安全弁 一覧(要件6.6)

`security-reviewer` が横断確認を行う際の起点として、既知の安全弁(上限定数)を一覧にする。
新しく上限を追加した場合はここに追記し、`src/Utsushi.Parsing/OpenXml/` 配下の
`foreach`/`Elements<...>()` ループを新規に追加した場合は、対応する上限がこの表に
載っているかを確認する。

| レイヤー / クラス | 定数 | 既定値 | 超過時の挙動 |
|---|---|---|---|
| Parsing / `OpenXmlWorkbookReader` | `MaxXlsxPackageBytes` | 1 GiB | `InvalidExcelFileException(TooLarge)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxSharedStringCount` | 200,000件 | `InvalidExcelFileException(TooLarge)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxCellsPerSheet` | 500,000個/シート | `InvalidExcelFileException(TooLarge)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxImagesPerSheet` | 50枚/シート | `UnsupportedWorkbookElementException(TooManyImages)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxImageDataBytes` | 10 MiB/画像 | `UnsupportedWorkbookElementException(ImageTooLarge)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapesPerSheet` | 50個/シート(図形・接続線・グループ合計) | `UnsupportedWorkbookElementException(TooManyShapes)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapeTextLength` | 2,000文字 | `UnsupportedWorkbookElementException(ShapeTextTooLong)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapeNestingDepth` | 5段 | `UnsupportedWorkbookElementException(GroupNestingTooDeep)` |
| Parsing / `OpenXmlWorkbookReader` | `MaxGradientStopsPerFill` | 64個/塗りつぶし | 打ち切り(超過分は無視、例外化なし) |
| Parsing / `OpenXmlWorkbookReader` | `MaxMergedRangesPerSheet` | 1,000件/シート | `UnsupportedWorkbookElementException(TooManyMergedRanges)` |
| Parsing / `StyleTable` | `MaxStyleTableEntries` | 10,000件(フォント/塗りつぶし/罫線/数値書式/cellXfs各々) | 打ち切り(超過分は既定書式にフォールバック、例外化なし) |
| Parsing / `ColorResolver` | `MaxIndexedColorCount` | 10,000件 | 打ち切り(超過分は呼び出し元のfallback色、例外化なし) |
| Layout / `PageCommandBuilder` | `MaxSpanCells` | 4,096セル | 結合セル・可視範囲の走査を打ち切り(例外化なし) |
| Layout / `PageCommandBuilder` | `MaxDrawingObjectDimensionPt` | 5,000pt | 描画オブジェクトの寸法をクランプ(例外化なし) |
| Rendering / `SkiaPdfRenderer` | `MaxDecodedImageDimensionPx` | 4,096px | `PdfRenderingException`(デコード前に宣言サイズを検査) |

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
- ~~OOXMLパーツ全体の非圧縮サイズに対する上限が無い~~ → `SpreadsheetDocument.Open`前に
  `GuardPackageSize`でZIPエントリの宣言サイズ合計に上限(`MaxXlsxPackageBytes`、既定1GiB)を
  設け、あわせて共有文字列数(`MaxSharedStringCount`)・シート内セル数(`MaxCellsPerSheet`)にも
  個別に上限を設けた。要件6.6として要件化済み。フォント・罫線等のスタイル要素の件数は
  範囲外索引が既定書式にフォールバックする既存の挙動で吸収できるため、例外化はせず
  読み取り数を`MaxStyleTableEntries`で打ち切るのみとした(詳細はParsingレイヤー節)。

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
- **ヘッダー/フッターの画像(`&G`)未対応**: `&G`は読み飛ばしている(文字色`&K`は対応済み。
  要件3.7, 3.8補足参照)。`&G`が参照する画像は、シートの`legacyDrawingHF`要素が指す
  VMLパート(`xl/drawings/vmlDrawingN.vml`)内の`v:shape`/`v:imagedata`経由で
  `xl/media/`の画像に辿り着く仕組み自体は特定できているが、どの`v:shape`がヘッダー/
  フッターのどのセクション(左/中央/右)・どのページ種別(奇数/偶数/先頭)に対応するかを
  示す紐付け規則を、ECMA-376の一次資料(Microsoft Learn経由で確認できる範囲)からも
  参照実装のソースからも確認できていない(本サンドボックスのネットワーク制限により
  GitHub上のOSS実装のソース取得ができなかったため)。誤った対応付けで実装すると
  画像が意図しないセクションに出る、または見つからずに終わるリスクがあるため、
  対象帳票で実際に`&G`入りのファイルが出てきた時点で、その実物を解析して
  紐付け規則を確定させたうえで対応する。
  `&F`(ファイル名)は、Utsushi が Stream を入力に取り元のファイル名を持たないため、
  帳票コードを代わりに展開している(`&Z`(ファイルパス)も同様の理由で同じ値を使う)。
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
- **図形プリセットのさらなる拡張**(要件10.1, 10.7): 拡張フェーズで星形4種・
  フローチャート記号7種・吹き出し4種(雲形・引き出し線1〜3本)・接続線5種を追加したが、
  フローチャート記号の残り(`flowChartOr`等)・自由曲線(`custGeom`)・より複雑な星形
  (`star10`以上)は引き続き「サポート外要素」である。対象帳票で実際に必要になった時点で
  一覧に追記する。
- **接続線の接続点(コネクションサイト)解決の残存する限界**(要件10.11): 要件10.11で
  `stCxn`/`endCxn`の解決に対応したが、以下の点は引き続き限界として残る。
  - 接続点の位置は、既定では配置矩形の上下左右の中点(4方向)で近似する。
    `flowChartInputOutput`の左右の接続点のみ実際の輪郭に合わせて補正するが、
    それ以外のプリセット(星形・矢印・吹き出し・三角形等)・画像・グループは4方向の
    近似のままであり、実際にExcel上で図形の辺・頂点以外の位置に接続点を作っている場合
    (例: 矢印の先端、星形の頂点)は見た目がずれる。
  - 接続先が接続線と異なるページに配置される場合(改ページで分割された場合)は解決せず、
    要件10.9の既定動作にフォールバックする。
  - 接続点の位置(4方向の近似、`flowChartInputOutput`の左右の接続点の補正)は
    ECMA-376の一次資料(`presetShapeDefinitions.xml`)への当たり直しができておらず、
    実装時にPDFをラスタライズして目視確認した推定値である。対象帳票で実際に見た目が
    ずれる場合に改めて検証する。
  - 接続点解決テーブル(`BuildConnectionTargetTable`)は参照先図形(画像も接続先になりうる。
    要件9.7)の回転前の`RectPt`のみを保持する。参照先自身、またはそれを含むグループが
    回転している場合、実際に見えている(回転後の)辺の位置と計算上の接続点はずれる。
    回転した図形・画像を接続先にする帳票が実際に出てきた場合に改めて対応を検討する。
  - 接続線自身が`RotationDegrees`を持ち、かつ両端点が解決済み(`ResolvedStart`/
    `ResolvedEnd`が共に非`null`)の場合、`ConnectorGeometryBuilder`は解決済みの
    絶対座標をそのまま線分の両端として使い、接続線自身の回転は適用しない
    (回転前提だった`Rect`基準の中心点が、絶対座標で指定された両端点に対しては
    意味を持たなくなるため)。この場合に`RotationDegrees`が非ゼロの`.xlsx`が
    実際に存在するかは未確認であり、対象帳票で問題になった場合に改めて検証する。
- **グループの回転と子要素の回転の合成の精度**(要件10.10): `GroupCommand`による
  `canvas`変換の入れ子でグループの回転・子要素個別の回転を合成する設計は、単体の
  目視確認では正しく動作することを確認したが、「グループ自身が回転しており、かつ
  グループの子座標空間の拡大縮小が非一様(縦横で倍率が異なる)」という組み合わせでは、
  回転と非一様スケールの適用順序によって見た目が変わりうる(アフィン変換は一般に
  可換ではないため)。この限界は子要素が図形(`GroupChildShape`)の場合と同じ経路
  (`ToGroupChildRect`で非一様スケール適用後、`canvas`回転を適用)を通る画像
  (`GroupChildImage`。要件9.7)にも同様に当てはまる。Excel自身がこの組み合わせを
  どう扱うかの一次資料での裏取りはしていない。対象帳票で実際に問題になった場合に
  改めて検証する。
- **雲形吹き出し(`cloudCallout`)・星形の近似精度の残存する限界**(要件10.12, 10.13):
  要件10.12/10.13で星形の既定内側半径比・雲形の引き出し位置の精度を改善したが、以下は
  引き続き近似のままである。雲形の輪郭(バンプの個数・半径)自体は固定値のままで
  `a:avLst`による微調整には対応しない(ECMA-376上も雲形の輪郭自体は調整ガイドを
  持たないため、これは近似ではなく仕様どおりである)。星形の既定内側半径比は二次資料の
  突き合わせによる推定値であり、ECMA-376一次資料そのものへの当たり直しはできていない。
  Excel側でファイルに`a:avLst`の指定がある場合はその値をそのまま使うため、既定値の
  精度が問題になるのはファイルに指定が無い場合のみである。
- **グループのネストに対するOpenXml SDK自体のDOM構築コスト**(要件10.8。security-reviewer指摘):
  `MaxShapeNestingDepth`(既定5段)は`ReadGroupChildGroup`のアプリケーションコード側の
  再帰にのみ効き、`WorksheetDrawing`への初回アクセス時に`DocumentFormat.OpenXml` SDKが
  XMLツリー全体を型付き`OpenXmlElement`ツリーへ変換する処理(SDK内部の再帰)には及ばない。
  理論上、数万段にネストした極小`<xdr:grpSp>`(1段あたり数十バイト)を仕込んだ`.xlsx`は、
  本プロダクトの`MaxShapeNestingDepth`チェックが実行される前にSDK側のXML→DOM変換の
  再帰でネイティブスタックを消費し、`StackOverflowException`(.NETでは捕捉不能・
  プロセスクラッシュ)を引き起こす可能性がある。これは`xdr:sp`/`xdr:pic`のみを扱っていた
  従来のスコープには無かった攻撃面で、ネスト可能な`xdr:grpSp`を読み取り対象に加えた
  今回の変更で新たに生じたものである。SDKに渡す前段でXMLの再帰深さを検査する、または
  変換処理をタイムアウト付きの別プロセスで実行する等の対策が考えられるが、
  現時点では「登録済み自社帳票のみを対象とする」という製品スコープ上のリスク許容として
  対応を見送る。対象帳票の運用形態が変わり任意のExcelファイルを受け付ける可能性が
  出てきた場合は、実装前に必ず再評価すること。
- **放射状グラデーションの`a:path`種別**(要件10.6): `a:path type="circle"`のみを
  厳密に扱い、`"rect"`/`"shape"`は同じ放射状近似にフォールバックする。矩形・図形に
  沿った塗りつぶしの厳密な再現は行わない。
- **画像・図形の上限がシート単位でありワークブック単位の合算上限が無い**(security-reviewer指摘):
  `MaxImagesPerSheet`/`MaxShapesPerSheet`はシートごとにリセットされるカウンタであり、
  ワークブック全体でシートをまたいだ合算上限は無い。`ReportPdfConverter.Convert`は
  常に帳票定義の`sheetName`1枚に処理対象を絞る(`WorkbookReadOptions.SheetNameFilter`)ため
  現状の呼び出し経路では実害は無いが、`IWorkbookReader`/`OpenXmlWorkbookReader`は
  `public`であり、`SheetNameFilter`を指定しない(全シート読み取り)呼び出し方をする
  将来のコードが現れた場合はシート数倍の積み上げに対する上限が無い。対象帳票で
  実際に必要になった時点でワークブック単位の合算上限を検討する。
