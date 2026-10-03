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
- `sheetFormatPr/@defaultColWidth` が無いシート(Excelが通常保存するブック)の既定列幅は、保存値ではなく
  「`baseColWidth` 文字分(×MDW)+ 余白5px」を8ピクセルの倍数に切り上げたピクセル数とする
  (`ExcelUnitConverter.DefaultColumnWidthToPixels`。Calibri 11 = MDW 7 で64px、ＭＳ Ｐゴシック 11 = MDW 8 で72px。
  Excelの画面上の既定列幅「8.43」「8.38」と一致する)。Parsingは既定列幅を知らない(MDWを持たない)ため、
  幅の指定が無い列の幅を `NaN` とし、`SheetModel.BaseColumnWidth` とともにLayoutへ渡す。以前は保存値 8.43 を
  既定列幅として換算しており、MDW 7 で59px(Excelより5px狭い)になっていた(code-reviewer指摘)。

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
- **セルが無い位置の書式(要件1.10)**: `row/@s`(`@customFormat` が真の行だけ)を `SheetModel.RowStyles`(行番号 → 書式)、
  `col/@style` を `SheetModel.ColumnStyles`(`ColumnStyleRange(FirstColumn, LastColumn, Style)` の範囲のリスト。
  `col` 要素をそのまま持ち、16,384 列ぶんに展開しない)、`cellXfs` の0番を `SheetModel.DefaultCellStyle` に読み取る。
  `SheetModel.GetEffectiveStyle(address)` が「セルの書式 → 行の書式 → 列の書式 → 標準の書式」の順で解決し、
  Layout(セルが無い位置の塗りつぶし・罫線、結合範囲の外周の罫線)と Substitution(セルが無い位置への置換)が
  同じ規則を使う。3つのプロパティは `init` とし、既定値(書式なし・`CellStyle.Default`)で既存の生成箇所を変えずに済むようにした。
  行・列の書式は `GetUsedRange`(使用範囲)に含めない(要件1.10補足)。行の書式の件数は `row` 要素の数と行番号の上限、
  列の書式の件数は `col` 要素の数の既存の上限に収まる。
- **数式セル**: 数式は評価せず、OOXMLにキャッシュされている計算結果の値のみを読み取る(要件1.6)。
- **数値書式**: `numFmt` を適用した表示文字列を `CellModel.FormattedValue` に持たせる。
  表示に使うセクションの色の指定(要件4.12)は `CellModel.FormatColor` に持たせる(下記「数値書式の色」)。
  汎用の数値書式エンジンではなく、自社帳票が使う範囲(金額・数量・日付・パーセント)のサブセット実装とする。
  解釈できない書式(指数表記・分数表記など)は例外にせず General 相当へフォールバックし、
  表示の崩れはゴールデンテストで検出する。
  - **組み込み書式(要件4.8)**: `NumberFormatter.BuiltInFormats` は ECMA-376 Part 1, 18.8.30 の既定値のうち、
    日本語(ja-JP)ロケールのExcelの表示に合わせる(14=`yyyy/m/d`、22=`yyyy/m/d h:mm`、5〜8=円記号付きの通貨、
    27〜36・50〜58=和暦を含む日本語ロケール固有の日付・時刻書式)。以前はロケールに依存しない既定値だけを持ち、
    14がゼロ埋めの`yyyy/mm/dd`になり、5〜8・27〜36・50〜58は標準(General)の表示に落ちていた。
  - **経過時間**: `[h]`・`[m]`・`[s]`(同じ文字の連続)は24時間・60分・60秒を超えて数える。時・分・秒の表示と同じく
    ミリ秒に丸めてから数える。負の値は表示できないため General にする。秒の小数部(`ss.0`〜`ss.000`)は切り捨てで表示する。
  - **和暦**: `[$-411]` 等のロケール指定は表示に反映しないが、`g`(英字1文字 `R`)・`gg`(漢字1文字 `令`)・`ggg`(元号名 `令和`)・
    `e`(元号の年)・`ee`(2桁)を、明治(Excelの扱いに合わせ1868/1/1開始)〜令和の元号表で解釈する。明治より前の日付は General にする。
    1年は「1」と表示し、「元年」には対応しない。解釈できない英字の書式指定子(`b` 等)は、書式文字をそのまま出さず General にする。
  - **数値書式の色(要件4.12)**: `FormatSection` が、セクション内の `[...]` のうち色名(英語の8色・日本語版の8色)と
    `[ColorN]`(1〜56。既定のインデックスカラーの N+7 番)を色として解析して持つ。`NumberFormatter.ResolveColor(value, formatCode)` が
    `FormatNumber` と同じ規則でセクションを選び、その色を返す(色が無い・General にフォールバックする場合は null)。
    `OpenXmlWorkbookReader` は数値セルにだけこれを `CellModel.FormatColor` として持たせ、`CellModel.WithText`(置換)は
    `FormatColor` を消す。Layout は `FormatColor` があればフォントの色の代わりに使う。
  - **セクションと負号**: `@` を含むセクションは文字列の表示用であり、数値の表示には使わない(数値に使えるセクションが無ければ General)。
    セクションが複数あり負数セクションが選ばれた場合は絶対値に書式を適用し、セクションが1つだけの場合は負号を出力全体の
    先頭に付ける(`"¥"#,##0` で `-¥1,000`)。小数部の `#` は不要な0を出さず、`?` は空白にする。
    `h"時"mm"分"` のように引用符・角括弧のリテラルを挟んでも、前後の書式指定子を見て `m` を「分」と判定する。
  - **解析結果のキャッシュ**: 書式コードごとの解析結果(不変)を `ConcurrentDictionary` に持ち、数値セルごとの再解析を避ける。
    キャッシュするのは255文字(Excelの書式コードの上限)以下の書式コードだけで、件数は1,024件までとする
    (`MaxCacheableFormatLength`・`MaxCacheEntries`)。件数が上限に達したらキャッシュを空にして登録し直す
    (先着の書式で埋まったままにすると、常駐プロセスでは1回の入力で埋められた後、以降の変換の書式がすべて都度解析になる。
    security-reviewer指摘)。プロセス内で共有する静的なキャッシュのため、入力ファイルの内容で際限なく大きくならないようにしている。
- **行・セルの番地と行高(要件1.7)**: `row/@r`・`c/@r` が省略されていれば、ECMA-376どおり直前の行の次の行・
  同じ行の直前のセルの次の列(行の先頭ならA列)とする。補った行番号にも `MaxCellsPerSheet` による行番号の上限を当てる。
  `r` があるのに番地として解釈できないセルは読み飛ばす。行高は `customHeight` の有無によらず `row/@ht` を採用する
  (Excelは自動調整された行にも `ht` を書く)。
- **ファイル内の数値の検証(要件6.12)**: 行高・列幅・既定の行高/列幅・余白・フォントサイズ・先頭ページ番号・図形の
  テキストの余白は、NaN・無限大・負・上限超えを既定値(または指定なし)に戻す(上限は「信頼できない入力に対する安全弁 一覧」)。
  NaNの余白で描画の原点がNaNになり、白紙のPDFが出ていたため。列幅換算(`ExcelUnitConverter.ColumnWidthToPixels`)も
  NaN を0として扱う。
- **用紙サイズ(要件1.8)**: `PaperSizeTable` のコードの意味は Windows の `DEVMODE.dmPaperSize`(`wingdi.h` の `DMPAPER_*`)に
  合わせる(Excelはプリンタの用紙コードをそのまま保存する)。ECMA-376の説明と食い違う箇所(12/13)も `wingdi.h` に従う。
  以前は43を「B4(JIS)」、62を「B4(JIS)回転」としていたが、`wingdi.h` では43ははがき、62はB5(JIS)横送りである。
  収録するのは自社帳票で使いうる A/B判・Letter/Legal・はがき・往復はがき・封筒(角形2号・3号、長形3号・4号、洋形4号)と、
  それらの回転(75〜89・92)・横送り(55・61・62・67)。回転・横送りの用紙は幅と高さを `wingdi.h` の定義どおりに持ち、
  印刷の向きは `orientation` が別に決める。一覧に無いコードはA4とする。
- **ページ中央・先頭ページ番号**: `printOptions/@horizontalCentered`・`@verticalCentered` を `PageSetupModel.HorizontalCentered`・
  `VerticalCentered` に、`pageSetup/@firstPageNumber` を `@useFirstPageNumber` が真の場合だけ `PageSetupModel.FirstPageNumber` に読む
  (`MaxFirstPageNumber` を超える値・解釈できない値は指定なし)。xsd:boolean の属性は、型付きプロパティの `Value` が不正な値で
  例外を投げるため、属性の文字列("1"/"true")を直接解釈する。
- **手動改ページ(要件3.11)**: `brk/@id` は0始まりの行(列)番号で、その行の上(列の左)で改ページする(ECMA-376 Part 1, 18.3.1.3)。
  モデルは「この1始まりの番号の手前で改ページ」を表すため `id + 1` にする。`id = 0` と最終行(列)より後ろを指すものは無視する。
- **印刷タイトルの範囲外の行・列**: `_xlnm.Print_Titles` の行番号・列番号が範囲外(`$0` や最大行を超える値)の場合は、
  その方向の印刷タイトルなしとして無視する(以前は Layout で `ArgumentOutOfRangeException` になっていた)。
- **セルの文字列と色**: インライン文字列のリッチテキストは共有文字列と同じく `r/t` だけを連結し、ふりがな(`rPh`)を本文に混ぜない。
  セルの色(`rgb`・`indexedColors`)のアルファは Excel と同じく無視して不透明にする(`"00FF0000"` を透明として読むと、文字・罫線・塗りが消える)。
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
  罫線合成は、当初セルごとに結合範囲の個数に比例する線形走査(`SheetModel.FindMergedRange`)を行っていたため、
  個数を無制限に許すと処理量がページ内セル数×結合範囲数で増大する(画像・図形の個数上限
  (`MaxImagesPerSheet`/`MaxShapesPerSheet`)と同じ理由によるDoS対策。security-reviewer指摘)。
  現在は印刷範囲ごとに1度だけ作るセル→結合範囲の索引(`MergedCellIndex`。Layoutレイヤー節)を引くため、
  処理量は結合範囲の個数に比例しないが、上限は索引のメモリを抑える安全弁として残す。
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
    - 同じ画像パートを参照する画像(同じロゴを複数箇所に貼った場合など)は、読み取り・検証済みのバイト列を共有する
      (画像パートごとに `ConditionalWeakTable` で保持)。参照ごとに読み直すと、圧縮後は小さい画像を多数の `xdr:pic` から
      参照させるだけで、展開後の大きさ×参照数のメモリを確保させられる(security-reviewer指摘)。
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
  - **テーマの色・スタイル参照(要件10.15, 10.16)**: 色の読み取りは`DrawingColorResolver`
    (Parsing内部)に集約する。`a:srgbClr`/`a:schemeClr`/`a:sysClr`を基本色として読み、子要素の修飾
    (`lumMod`/`lumOff`はHSL輝度、`tint`/`shade`は線形RGBで白/黒へ補間、`alpha`は不透明度)を順に適用する。
    `a:schemeClr`の`phClr`は、スタイル参照(`lnRef`等)の色を差し込む位置を表す。テーマの配色は
    `ColorResolver`が読み込んだものを使い、色名(`dk1`/`lt1`/`dk2`/`lt2`/`accent1`〜`6`/`hlink`/`folHlink`
    と別名`tx1`/`bg1`/`tx2`/`bg2`)から引けるようにする。
    `spPr`に塗りつぶし(`solidFill`/`gradFill`/`noFill`)・枠線の色が無い場合は`xdr:style`を見る。
    - `fillRef/@idx`: 0は塗りなし。1以上はテーマの`fillStyleLst`(1000以上は`bgFillStyleLst`)の書式を指すが、
      Excelの既定テーマではいずれも`phClr`を基にした塗りのため、スタイルの色の単色で近似する。
    - `lnRef/@idx`: 0は線なし。1以上はテーマの`lnStyleLst[idx-1]`の`@w`を太さに、スタイルの色を線の色にする。
      `a:ln`があって色だけ無い場合は、`a:ln/@w`を優先する。`lnStyleLst`に該当する番号が無い場合の太さは0.75ptとする。
      `a:ln`に色があってスタイルも`@w`も無い場合は1ptとする(従来の既定の細線)。
    - `fontRef`: 文字の`a:rPr`に色が無いときの既定の文字色にする。`a:rPr`の色が未対応の指定(`a:prstClr`等)で
      解決できない場合も、この既定の文字色で代用する(塗りつぶし・線は「なし」にできるが、文字は色なしにできないため)。
    - テーマが無いブックでは、配色はOffice既定テーマの配色、`lnStyleLst`の太さはOffice既定テーマの値(0.5/1/1.5pt)を使う。
    テーマの書式設定(`a:fmtScheme`)は`lnStyleLst`の太さだけを読み、要素数は各リストの先頭の数個
    (Excelのテーマは3個)に限る。
  - **反転(要件10.17)**: 図形・グループ・グループ内図形の`a:xfrm/@flipH`・`@flipV`を読み取り、`ShapeModel`・
    `GroupChildShape`・`GroupShapeModel`・`GroupChildGroup`に持たせる。Layoutはグループの反転を、子要素の
    配置矩形をグループの矩形内で鏡映し、子要素の反転フラグを反転させることで展開する(接続線は
    始点・終点の座標を鏡映する)。Renderingは図形の中心を軸に、回転の内側で`canvas.Scale(-1, 1)`等を
    かけて形状を描き、文字はその変換の外(回転の内側)で描く。上下反転の図形では、文字を図形の中心で180°回して描く
    (Officeは上下反転した図形の文字を逆さにし、鏡像にはしない。この挙動はOfficeの一般的な仕様に基づくもので、
    Excelで保存した実物との突き合わせは未実施。未決事項を参照)。接続点の解決(要件10.11)は、反転前の図形上の
    接続点を配置矩形の中心で鏡映する(左右反転なら接続点1と3、上下反転なら0と2が入れ替わるのと同じ)。反転した
    グループ自身・グループ内の画像を接続先にした場合も、グループの反転を反映する。グループ内の画像の中身は鏡像にしない。
  - **矢印(要件10.18)**: `a:ln/a:headEnd`・`a:tailEnd`の`@type`・`@w`・`@len`を`ShapeOutline`の
    `HeadEnd`/`TailEnd`(`LineEndStyle`)として読み取る。`headEnd`は線の始点、`tailEnd`は終点に付く。
    大きさは線の太さの倍数(`sm`=2倍、`med`=3倍、`lg`=5倍。線が細い場合でも見えるよう最小寸法を設ける)で
    近似する(倍率・最小寸法はExcelの出力との実測での突き合わせが未実施。未決事項を参照)。
    接続線は`Outline`がnullのときRenderingが既定の黒い線を補うため、線を明示的に消している場合(`a:ln/a:noFill`、
    または図形のスタイルがあって線の色が決まらない場合)は透明の線として読み取り、既定の線と区別する。
    逆に、色の無い`a:ln`に矢印だけがある場合は、既定の黒い線(`a:ln/@w`、無ければ1pt)に矢印を付けて読み取る。
    Renderingは接続線・線吹き出しの引き出し線の端点と、端での線の向きから矢印の形状を組み立て、
    `triangle`/`stealth`/`diamond`/`oval`は線の色で塗り、`arrow`は開いた線で描く。
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
- **帳票定義なしモード(要件12)**: 定義をJSONからロードせず、`ReportDefinition.CreateWithoutDefinition(documentName, sheetName, maxDigitWidthPx)`
  で定義を合成する(`SubstitutionFields`は空、`UnsupportedElements = Ignore`、`ToleranceMm`は既定値、
  `MaxDigitWidthPx`は後述の見積もりまたは呼び出し元の指定、`PrintAreaOverride = null`)。合成した定義は以降の
  `IReportModelBuilder.Build`・Substitution・Layout・Renderingに定義ありと同じ形で渡るため、
  後段のレイヤーには分岐を入れない(帳票固有の分岐を共通レイヤーに置かない方針と同じ考え方)。
  - ファサード `ReportPdfConverter` は帳票コードを取らないオーバーロード
    `ConvertWithoutDefinition(Stream, Stream, cellOverrides?, documentName?, maxDigitWidthPx?)` /
    `ConvertFileWithoutDefinition(string xlsxPath, string outputPath, cellOverrides?, documentName?, maxDigitWidthPx?)` /
    `ComputeLayoutWithoutDefinition(Stream, cellOverrides?, documentName?, maxDigitWidthPx?)` を持つ。
    メソッド名を分けるのは、既存の `Convert(string reportCode, …)` と引数の並びが似ており、
    オーバーロード解決で意図しない方が選ばれるのを避けるため。置換キーの辞書は受け取らない(要件12.4)。
    帳票定義を使わない呼び出し元のために、帳票定義ルートを取らない `CreateDefault(fontOptions, renderOptions)` も用意する
    (このとき帳票コードを指定した変換は `ReportDefinitionNotFoundException` になる)。
  - 対象シートの決定(要件12.2)はParsingレイヤーの責務とし、`WorkbookReadOptions.ActiveSheetOnly = true` で
    指定する。`OpenXmlWorkbookReader` は `workbookView/@activeTab` のシート(非表示、またはグラフシート等の
    ワークシートでない場合は、表示されている最初のワークシート)1枚だけを読み、サポート外要素の検出・シート単位の安全弁もそのシートにだけ適用する
    (全シートを読んでから1枚を選ぶと、使わないシートの上限超過で失敗しうるため)。ただしXMLの入れ子の深さ・パートの
    大きさの検査(要件6.7)は、DOMを組み立てる前にパッケージ内の全XMLパートに対して行う。
  - 列幅換算の最大数字幅(要件12.3)は、呼び出し元の指定(API の `maxDigitWidthPx`、CLI の `--max-digit-width`)が
    無ければ `ReportDefinition.EstimateMaxDigitWidthPx` でブックの標準フォントから見積もる。見積もりは、Excelが既定の
    列幅として表示する値(英語版 Calibri 11 の「8.43(64ピクセル)」、日本語版 ＭＳ Ｐゴシック/游ゴシック 11 の
    「8.38(72ピクセル)」)を列幅の換算式に当てはめて逆算した固定の表であり、実行環境のフォントには依存しない
    (環境のフォントの版の違いで結果が変わらないようにするため。帳票定義の `maxDigitWidthPx` を自動導出しない
    理由と同じ)。表に無いフォント・サイズは既定値7を使う。
  - 文書名(要件12.5)は合成定義の `ReportCode` に入れる。これにより PDF タイトル・`&F`/`&Z`・例外の
    `ReportCode` に既存の経路のまま反映される。

- **使用範囲の外の差し込みセル(要件1.11)**: `ReportModelBuilder.ValidateSubstitutionCells` は、置換キーの対象セルが
  `SheetModel.GetUsedRange`(値か書式のあるセル・結合範囲の範囲)の外なら `ReportStructureMismatchException` とする。メッセージには
  対処(そのセル自体に値か書式を設定する。行全体・列全体の書式は使用範囲に含まれない(要件1.10補足))を含める。

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
- 必須キー未指定・未知キー指定はここで例外(要件2.3, 2.4)。必須キーに空文字・空白のみの値が渡された場合も
  `RequiredSubstitutionValueMissingException` とする(要件2.12)。
- **差し込み値の検証・正規化(要件2.10, 2.11)**: `Apply`・`ApplyCellOverrides`の両経路で、書き換えの前に
  全値を`NormalizeValue`に通す。`null`、改行以外の制御文字(`char.IsControl`。タブ・NUL・DEL・C1を含む)、
  不可視の書式文字(`UnicodeCategory.Format`。ゼロ幅スペース・BOM・ソフトハイフン等)、
  対になっていないサロゲートは`InvalidSubstitutionValueException`(Stage=Substitution。`Target`に置換キー
  またはセル番地の文字列、`CellAddress`に対象セル)とする。CRLF・CR単独・U+2028・U+2029はLFへ統一する。
  `ApplyCellOverrides`でも、必須キーのセルを空文字・空白のみにする指定は`RequiredSubstitutionValueMissingException`とする(要件2.12)。
  値の「内容」の正規化のみで、書式には触れない(要件2.2)。
- **差し込み済みセルの記録(要件2.13, 2.14)**: 空でない値を差し込んだセルを`ReportModel.SubstitutedCells`に
  記録する(空文字で値を消したセルは外す)。Layoutレイヤーは、テンプレート自身の文字列とは区別して、
  この集合のセルに限り欠落の検出を行う。`ReportModel`の位置パラメータには加えず`init`プロパティとし、
  既存の`ReportModel.Create`/`with`式の呼び出し側を変えずに済むようにした。
- はみ出し時の挙動(`overflow: overflow|shrink|clip|wrap`)は帳票定義の値をそのままLayoutレイヤーに引き渡すためのフラグとして `ReportModel` に保持する(実際の折り返し/縮小計算はLayoutレイヤーの責務)。
- `ApplyCellOverrides` は、帳票定義の置換キー(`SubstitutionFields`)を経由せず、セル番地(A1形式の文字列。キーは `CellAddress.TryParse` で解釈する)を直接指定して値を書き換える第二の経路(要件2.7, 2.8)。
  - `Apply` と同じく `CellModel.WithText` で値のみを差し替え、書式には触れない。対象セルが未存在(空セル)の場合は
    `SheetModel.GetEffectiveStyle` の書式(行・列・ブックの標準の書式。要件1.10)で新規セルを作る点も `Apply` と同一。
    以前は `CellStyle.Default`(Calibri 11pt)で作っていたため、テンプレートで使っていない Calibri を厳格モードのフォント解決が要求していた。
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
  - 手動改ページの適用(要件3.2, 3.11): `PageBandCalculator.Split` は「直前の行(列) < 改ページ位置 <= 現在の行(列)」を
    満たす改ページがあれば現在の行(列)の手前で改ページする。改ページ位置の行(列)そのものが本文の並びに無い
    (非表示・高さ0・印刷タイトル)場合も改ページが失われない。
  - 自動改ページ計算(1ページの印字可能領域 = 用紙サイズ - 余白、を拡大縮小率で除した論理サイズに対し、行高/列幅の累積が収まる位置で分割)。
    「収まるか」の比較には許容誤差 `PageBandCalculator.TolerancePt`(1e-6pt)を設け、行高の合計や倍率での割り算の浮動小数点誤差で
    ちょうど収まる行(列)が次のページへ送られないようにする。1ページに使える大きさは帯の先頭の行(列)によって変わってよい
    (タイトルを付けるページだけタイトルの分狭くなる)ため、`Split` は帯の先頭を受け取って使える大きさを返す関数を取る。
  - 印刷タイトル行/列の複製(印刷タイトルは印刷範囲と独立に指定でき、印刷範囲の外の行/列でも各ページに繰り返す。要件3.4, 3.12)。
    タイトルを付けるのは、本文の帯の先頭がタイトルの最終行(列)より後ろにあるページだけとする(`ReportLayoutEngine.TitlePlan`)。
    印刷範囲より上(左)のタイトルは全ページに付き、印刷範囲の中のタイトルはそれ自体を本文として印刷したページより後ろのページに
    だけ付き、印刷範囲より下(右)のタイトルはどのページにも付かない。本文は印刷範囲の行(列)だけとし、範囲内のタイトル行(列)を
    本文から取り除かない。どのページにタイトルを付けるかはページ分割の結果で決まるため、印刷範囲ごとに先に分割の計画
    (`RangePlan`: 格子・タイトル・行帯・列帯・拡大縮小率・結合範囲の索引)を立て、差し込みセルの検証(要件2.6, 2.13)は
    その後に行う。実際にはどのページにも付かないタイトル行(列)にある差し込みセルは、印刷範囲外と同じくエラーにする。
  - 「次のページ数に合わせて印刷」(要件3.13): 本文の合計の大きさから求まる倍率(整数%に切り捨て、100%以下)を上限とし、
    下限 `MinFitScalePercent`(10%)〜上限の整数%のうち、実際に行帯・列帯へ分割したページ数が指定以下になる最大の値を
    二分探索で求める(倍率を下げても帯の数は増えないため単調。1%ずつ下げると大きなシートで最大90回の分割になる)。
    どれも収まらなければ下限の10%とする。タイトルの分や行・列の区切り位置で、合計から求めた倍率では指定のページ数に
    収まらないことがあるため、実際に分割して数える。この場合、手動改ページは無視する(Excelの仕様)。
  - 「ページ中央」(要件3.14): 本文(タイトルを含む)の大きさ×倍率が印字可能領域より小さければ、差の半分だけ本文全体を
    平行移動する(`PageCommandBuilder` の原点 = 余白 + 中央寄せの移動量)。大きい場合は移動しない。
  - 高さ0の行・幅0の列(要件4.10): 非表示の行・列と同じく印刷しない。判定は `SheetGrid.PrintedRowHeightPt`・
    `PrintedColumnWidthPt`(非表示または0以下・NaNなら0)に集約し、格子の構築・結合セルの矩形計算・差し込みセルの検証
    (`SheetGrid.IsRowNotPrinted`/`IsColumnNotPrinted`)で同じ基準を使う。
  - セル→結合範囲の索引(`MergedCellIndex`): 印刷範囲ごとに1度だけ作り、その印刷範囲の全ページで使い回す。
    行方向を「かかる結合範囲の組が変わらない行の区間」(結合範囲の個数の2倍+1以下)に分け、区間ごとにその区間にかかる
    結合範囲を先頭列の順に並べて持つ。格子の列にかからない結合範囲は登録しない。メモリは結合範囲の行数や格子の行数に
    比例しない(行ごとに登録すると、全行にわたる縦長の結合範囲が多いシートで「格子の行数×結合範囲数」になる。
    layout-fidelity-reviewer指摘)。結果は `SheetModel.FindMergedRange` と同じ(不正なファイルで結合範囲が重なる場合も、
    `MergedRanges` の並びで先に現れる範囲を返す)。
  - 複数の印刷範囲をそれぞれ独立したページ群として計算(要件3.6)
  - 必須の置換フィールドの対象セルが、印刷範囲・印刷タイトルのいずれにも含まれない場合はエラーとする(要件2.6)
  - 空でない値を差し込んだセル(`ReportModel.SubstitutedCells`。任意キー・セル番地直接指定を含む)が
    印刷範囲・印刷タイトルのいずれにも含まれない場合もエラーとする(要件2.13、`ValidateSubstitutedCellsAreInPrintRanges`)
  - **セル内テキストの改行と折り返し(要件4.6, 4.7)**: 折り返し表示のセルは`WrapLines`で、LF・CRLF・CR単独を
    段落区切りとしたうえで、`StringInfo.GetTextElementEnumerator`による書記素クラスタ単位で幅を測って
    折り返す(UTF-16の1単位ごとに区切るとサロゲートペアや異体字セレクタの途中で改行されるため)。
    図形内テキストも同じ`WrapLines`を使う。実際の折り返しは`TextWrapper`が行い、1文字足すごとに行全体を
    測り直す(行の長さの2乗に比例する計測になる)代わりに、計測済みの幅から改行位置の見当を付けて行全体の計測で
    確かめる。文字列を後ろへ伸ばしても幅が減らない(単調である)前提のもとで、1文字ずつ測り直す方法と同じ位置で改行する。
    折り返し表示でないセル(はみ出し・切り取り・縮小)では、
    Excelと同様に改行文字を取り除いて1行として配置する(`RemoveLineBreaks`)。
  - **折り返した差し込み値の欠落検出(要件2.14)**: 折り返し表示で、かつ`SubstitutedCells`に含まれるセルは、
    各行の字面(アセント+ディセント)が、字面の高さの4分の1の許容量を超えてセルの外に出ないことを確かめ
    (`EnsureSubstitutedTextFits`)、出る行があれば`LayoutComputationException`(`CellAddress`に対象セル)とする。
    字面全体が収まることを条件にしないのは、フォントの行送りがExcelの行高よりわずかに大きいことが多く、
    Excel上で収まっている住所まで誤検出するため。判定はページ上に見えている矩形ではなく、セル(結合範囲なら
    範囲全体。非表示の行・列は0として合計する)のシート上の本来の大きさで行う。結合範囲が改ページ・印刷範囲の端・
    非表示行にかかって一部しか見えないページでも判定を省かないため(layout-fidelity-reviewer指摘)。
    同じ関数で、差し込みセルの幅(余白・インデントを除く)または高さが0以下で何も描画できない場合もエラーにする(要件2.13)。
    非表示の行・列にある差し込みセルは、ページ上に行・列自体が現れないため、`ValidateSubstitutedCellsAreInPrintRanges`で
    印刷範囲外と同様にエラーにする。判定の結果はセルの番地・書式・(印刷範囲で共通の)拡大縮小率で決まるため、
    確かめ終えたセルは印刷範囲ごとに記録し、印刷タイトルの行/列として複数ページに繰り返し現れるセルを
    ページごとに検証し直さない。
  - **文字の収まりの確認(要件13)**: `PageCommandBuilder.EmitText` で、描画と同じフォント・配置を使って
    収まらない箇所を `FitIssueCollector` に記録する。描画命令は変えない。判定は次のとおり(幅は `MeasureTextWidth`)。
    結合範囲は、改ページで一部しか見えないページでも、シート上の本来の大きさ(`SumBounded`)で判定する。見えている部分で判定すると、
    範囲全体なら収まる文字を誤って報告するため(要件2.14の `EnsureSubstitutedTextFits` と同じ基準。code-reviewer指摘)。
    - 数値セル(`CellValueKind.Number`)で縮小表示でないもの: 1行の幅が内容矩形の幅を超えれば `NumberTooWide`。
      数値はこの判定だけを行う(Excelは数値を折り返さず、はみ出させもせず `####` にするため)。表示形式が標準(General)の数値は、
      Excel が桁を減らすか指数で表示するため、説明の文をそれに合わせる(layout-fidelity-reviewer指摘)。
    - はみ出し表示(結合範囲でない): 文字がはみ出す側(左揃え → 右、右揃え → 左、中央揃え → 両側)の、同じ行で
      このページに並ぶ隣の列を、はみ出した幅に届くまで順に見る。値を持つセル(結合範囲ならアンカーの値)に
      かかれば `OverlapsNeighborValue`。数式のセルは、結果が空文字でも Excel ははみ出しを止めるため、値があるものとして扱う
      (layout-fidelity-reviewer指摘)。はみ出した先がページの本文の矩形の外に出れば `CutAtPageEdge`。
      均等割り付け・両端揃え・繰り返しは描画が左揃えのため(上記「未決事項」)、Excel でははみ出さないこれらも描画どおり判定する。
    - 切り取り表示、およびはみ出し表示の結合範囲: 1行の幅が内容矩形の幅を超えれば `Clipped`。
    - 折り返し表示: 描画と同じ縦位置(`ResolveFirstBaselineY`)に置いた各行の字面が、字面の高さの4分の1を超えてセルの外に出れば
      `ExceedsCellHeight`(`WrappedLinesFit`。要件2.14の `EnsureSubstitutedTextFits` と共通の判定)。差し込みセルは
      `EnsureSubstitutedTextFits` が先にエラーにするため、テンプレートの文字だけが対象になる。
    記録は印刷範囲をまたいで1つの `FitIssueCollector` で行い、(セル, 種類) が同じものは最初に現れたページの1件にまとめる。
    記録しないものの説明の文を組み立てないよう、判定の前に `Accepts` で確かめる。
    説明の文(`FitIssue.Message`)にはセル番地・置換キー・対処だけを入れ、セルの文字列は入れない(宛名などの個人情報がログへ流れないように。
    文字列は `FitIssue.Text`。security-reviewer指摘)。置換キーは、置換キーで差し込んだセルにだけ付ける(セル番地直接指定で上書きした
    セル `ReportModel.OverriddenCells` は除く)。
    `ReportLayoutEngine.Compute(report, checkFit)` が `PagedLayout.FitIssues` に入れて返す(`Compute(report)` は確認する)。
    ファサードの `ComputeLayout` と `CheckFit` / `CheckFitWithoutDefinition` は確認し、PDFへの変換(`Convert` 等)は結果を使わないため
    確認を省く(長い文字列のセルが多いと計測が変換の時間を大きく増やすため。security-reviewer指摘)。`IReportLayoutEngine.Compute(report, checkFit)` は
    既定の実装を持つ(既存の実装を壊さない)。`CheckFit` はレイアウト計算までを行うため、PDFの描画で初めて分かるエラー(字形の無い文字の
    `MissingGlyphException` など)は送出しない。
    ゴールデンのスナップショットにも出力し、検出結果の回帰を確かめる。
    `PagedLayout.FitIssues`・`CellModel.FormatColor`・`SheetModel.RowStyles` 等は `init` プロパティのため record の等値比較に含まれる
    (リスト・辞書は参照比較)。これらのモデルを等値比較で使う箇所は無い。
  - **塗りつぶしの統合**: 同じ色の塗りつぶしが同じ行の高さで左右に接していれば1つの矩形にまとめる(`AddFill`)。
    行全体・列全体の書式(要件1.10)で並ぶ塗りつぶしをセルの数だけの矩形にしない(PDFのサイズと、ビューアで境目に見える細い線を防ぐ)。
  - **ヘッダー/フッターの改行**: 複数行のヘッダー/フッターは未対応のため、改行を取り除いて1行に配置する
    (`HeaderFooterCommandBuilder.ScaleRuns`。以前は改行文字がそのまま描画され豆腐や空白になっていた)。
  - ページヘッダー/フッターの書式コード展開と配置(要件3.7〜3.9, 3.15, 3.16)。
    「先頭ページ番号」が指定されていれば、`&P` は `先頭ページ番号 + ページの通し番号 - 1`、`&N` は最終ページの番号
    (`先頭ページ番号 + 総ページ数 - 1`)とする。「奇数/偶数ページで別指定」はこの表示上のページ番号の奇数・偶数で選び、
    「先頭ページのみ別指定」は文書の1枚目のページ(`HeaderFooterContext.IsFirstPage`)で選ぶ(先頭ページ番号を指定すると、
    1枚目のページ番号が1とは限らないため)。`&P+n`・`&P-n` は直後の数字(最大9桁。`MaxPageOffsetDigits`)を足し引きし、
    数字の続かない `&P+` はページ番号の後に `+` をそのまま残す。フォント指定 `&"フォント名,スタイル"` のスタイルは
    英語名(`Bold`/`Italic`/`Oblique`)と日本語版Excelの名前(`太字`/`斜体`)の両方を解釈する。
    先頭ページ番号を指定した場合の `&N` と奇数/偶数の選び方は、Excelの実機との突き合わせが未実施である。
  - 結合セルの矩形統合
  - **はみ出し表示のクリップ(要件4.9)**: 切り取り・折り返し・縮小のセルに加え、結合範囲の文字もセル(結合範囲)の矩形で
    切り取る(Excelは結合セルの文字を隣のセルへはみ出させない)。結合範囲でないセルのはみ出し表示は、隣のセルへは描くが、
    このページの本文(タイトルを含む)の矩形を印字可能領域(余白の内側)で切り詰めた矩形(`PageCommandBuilder._pageBodyRect`)の
    外へは描かない(1列だけで印字可能領域を超えるページで、文字が余白へ描かれていた)。
  - **二重線の罫線(要件4.11)**: Excelの二重線は「1pxの線・1pxの空き・1pxの線」で描かれるため、2本の線の中心の間隔を
    2px(`BorderMetrics.DoubleLineCenterSpacingPt` = 1.5pt)に拡大縮小率を掛けた値とする(以前は中心の間隔を1ptとしていた)。
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
  - **図形内テキストの余白・改行・空の段落(要件10.19)**: 矩形の内側の余白は `ShapeTextBody.Insets`(`a:bodyPr` の
    `lIns`/`tIns`/`rIns`/`bIns`。指定が無ければ DrawingML の既定値 `ShapeTextInsets.Default` = 左右7.2pt・上下3.6pt)に、
    拡大縮小率(グループ内なら追加の係数も)を掛けて使う。以前は全辺に固定の4ptを使っていた(`ShapeTextPaddingPt`)。
    Parsingレイヤーは `a:br` を `"\n"` 1文字のラン、`a:fld` を保存時点の文字列のランとして読み、ランの無い段落も
    空の段落として残す。Layoutは `"\n"` を行の区切りとし、空の段落は1行分の空行にする。空の段落の行の高さは、
    直前(無ければ直後)のランを持つ段落の先頭ランのフォントに拡大縮小率を掛けて決める。
  - **図形・接続線の枠線の倍率(要件10.20)**: 図形・グループ内図形・接続線の枠線の太さ(矢印の大きさはこれに比例する)に
    拡大縮小率を掛ける。枠線の指定が無い接続線は、既定の黒い実線1pt(`DefaultConnectorOutline`)をLayoutで補ってから
    倍率を掛ける(Renderingで補うと倍率が掛からないため)。
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
    (`a:bodyPr`の余白。下記「図形内テキストの余白」)・フォントサイズにも`scaleX`と`scaleY`の幾何平均を追加の係数として
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

- **印刷範囲が無いシートの使用範囲(要件3.10)**: `UsedRangeResolver` が、セルの使用範囲
  (`SheetModel.GetUsedRange`)に、各描画オブジェクトが占めるセルの範囲を合わせる。二セルアンカーは終端のセル
  (終端のオフセットが0ならその1つ手前)まで、一セルアンカー(固定サイズ)はアンカーセルから列幅・行高をたどって
  終端が掛かるセルまでとする。たどる回数は4,096回で打ち切る(幅・高さ0の行・列が続く入力への安全弁)。
  非表示の行・列は幅・高さ0として扱い、描画オブジェクトの寸法は描画時と同じ上限(固定サイズは5,000pt、
  二セルアンカーの終端は始点から4,096行・列)に抑える(描画では切り詰められる寸法で使用範囲だけが広がらないように)。
  帳票定義の `printArea` による上書き・Excelの印刷範囲がある場合は使わない。

### 5. Rendering レイヤー (`Utsushi.Rendering`)

- **責務**: `PagedLayout` を `SkiaSharp` の `SKDocument`(PDFバックエンド)で描画し、PDFファイルを生成する。
- **主なインターフェース**:
  ```csharp
  public interface IPdfRenderer
  {
      void Render(PagedLayout layout, Stream output);
  }
  ```
- **罫線の端の形(要件4.11)**: 線端は常に `Butt` で描き、角の継ぎ目は Layout が埋める
  (`PageCommandBuilder.ExtendBorderEndsAtJunctions`)。実線の水平・垂直の罫線の端に、直交する罫線の端が接している(角・T字)場合だけ、
  接している罫線の線幅の半分だけ端を延ばす。端をそろえたままだと、太い罫線(2.25ptなど)が角で交わるところの外側に欠けができる。
  一律に線端を四角く(`Square`)すると、直交する罫線の無い端(合計欄の下罫線だけを引いた場合など)まで隣のセルへ突き出すため
  採らない(layout-fidelity-reviewer指摘)。罫線はセルの辺ごとに出力するため、継ぎ目は端点どうしの一致で判定する。
  破線は延ばすと破線の周期がずれ、二重線・斜線はセル境界の上に無いため延ばさない(継ぎ目の相手としては破線も数える)。
  T字に交わる箇所の見た目のExcelとの一致は未検証である。
- 1ページ = 1 `SKCanvas` への描画。矩形塗りつぶし(背景)→罫線→テキスト→
  **画像・図形・接続線・グループ**(`drawing.xml`の出現順)の順で描画する。Excelはシート上に
  浮かぶ描画オブジェクト(画像・図形等)をセルの内容より上のレイヤーとして描画するため、
  これらは他のセル内容と重なる場合に最前面へ来るようにする(要件9.3, 10.3)。
- **画像の描画(要件9)**: `ImageCommand` は `SKBitmap.Decode(byte[])`でデコードして`SKImage`に変換し、
  `SKCanvas.DrawImage(image, destRect)` で `ImageCommand.Rect` へ描画する
  (SkiaSharp 2.88.8で利用可能な標準API)。全ページを通して2回以上描く画像(同じバイナリの配列を
  参照する画像)だけは、デコード結果を1回の出力の間`RenderContext`に記録して使い回す。印刷タイトルの行にある画像のように
  全ページに現れる画像も、デコードは1回で済み、PDFには1回だけ埋め込まれる。1回しか描かない画像は描画後すぐに解放する
  (全画像を記録すると、出力を終えるまで「画像の枚数×1枚のピクセル数の上限」のメモリを抱えるため)。既存の `ToSkRect(RectPt)` をそのまま使う。
  ページ境界外にはみ出す部分は `SKCanvas` が自然にクリップするため、追加のクリップ処理は不要。
  `SKBitmap.Decode` で実際に展開する前に `SKBitmap.DecodeBounds` で宣言上のピクセル寸法を確認し、
  上限(既定4096px)を超える場合はデコードせず `PdfRenderingException` とする(要件9.6。
  ピクセル爆弾対策。詳細はParsingレイヤー節「信頼できない入力に対する安全弁」を参照)。
  回転がある場合は、図形(要件10.5)と全く同じ`canvas.Save()` →
  `canvas.RotateDegrees(RotationDegrees, centerX, centerY)`(中心は`Rect`の中心)→
  `DrawImage` → `canvas.Restore()` のパターンで`RotationDegrees`を反映する(要件9.7。
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
    - 線吹き出し(要件10.14): `callout1`〜`3`/`borderCallout1`〜`3`/`accentCallout1`〜`3`/
      `accentBorderCallout1`〜`3`の12種は、折れ数N(1〜3)・本体枠線の有無・強調線の有無の
      組み合わせとして1つのロジック(`LineCalloutOutlinePath`)で描く。塗りつぶし用の`Build`は
      本体の`rect`のみ、枠線用の`BuildOutline`は「本体の枠線(`borderCallout`系・
      `accentBorderCallout`系のみ)」「強調線(accent系のみ。引き出し線の始点のX位置に本体の
      上端から下端までの縦線)」「引き出し線(N+1個の頂点を結ぶ開いた折れ線)」を持つ。
      ECMA-376上、`callout`(枠なし)系は本体の枠線を描かない(`stroke="false"`)ため、
      以前の実装(`callout1`〜`3`にも本体の枠線を描いていた)から改めた。
      引き出し線の頂点は調整ガイド`(adj1=y1, adj2=x1), (adj3=y2, adj4=x2), …`(本体の高さ・幅に
      対する比率)で決まり、ファイルに無い位置はECMA-376 presetShapeDefinitionsの既定値
      (折れ数1: 18750, -8333, 112500, -38333。折れ数2: 18750, -8333, 18750, -16667, 112500, -46667。
      折れ数3: 18750, -8333, 18750, -16667, 100000, -16667, 112963, -8333。いずれも1/100000単位)を補う。
      ファイル由来の極端な値はwedge系と同じ`CalloutTipAdjLimit`(±5倍)に収める。
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
  - `Outline`があれば`SKPaintStyle.Stroke`で描画する(Excel上は既定の黒い実線1ptで表示されるため、
    `Outline`が無い接続線にはLayoutレイヤーが既定の線色・太さを補い、拡大縮小率を掛けて渡す。要件10.20)。
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
- **字形の欠落検出(要件5.5)**: `DrawText`(セル・ヘッダー/フッター・図形内テキストの共通経路)で、
  描画前に各文字(`string.EnumerateRunes`)を`SKFont.GetGlyph`で字形IDに変換し、0(.notdef)になる文字が
  あれば`MissingGlyphException`(`FontName`・`CodePoint`・`Text`、帳票コード・シート名)を送出する。
  制御文字は判定しない。例外メッセージには宛名・住所などの個人情報がログへ丸ごと流れないよう描画文字列の
  先頭10文字だけを載せ、全文は`Text`プロパティで参照させる。`PdfRenderOptions.MissingGlyphs`(`MissingGlyphPolicy.Error`が既定、
  `Render`で従来どおり描画)で切り替えられる。`PdfRenderOptions`の位置パラメータには加えず`init`
  プロパティとし、既存の`with`式や`Default`/`OutlineText`の利用側を変えずに済むようにした。
  セル番地は`TextCommand`が持たないため例外に含まれない(`Text`で該当セルを特定する)。
- **フォント解決のスレッド安全性**: `FontResolver.Resolve`は、キャッシュに無い書体の生成をロックの内側で
  1回だけ行う(以前は`ConcurrentDictionary.GetOrAdd`の引数で生成済みの値を渡していたため、同時実行時に
  書体が重複して生成されていた)。インストール済みフォントから解決した書体は`Dispose`でも解放しない。
  SkiaSharp は同じネイティブ書体に対してプロセス全体で同一のマネージドオブジェクトを返す(実測で確認)ため、
  解放は同じプロセス内の別の`FontResolver`が使っている書体への操作になる(2.88.8では解放後も書体が使え続ける
  ことを実測したが、その内部挙動に依存しないよう解放しない)。
- **フォント埋め込み方針(決定済み)**: PDF内の文字列検索・コピーを維持するため、
  **フォントを埋め込む `PdfTextRendering.EmbedFont` を既定とする**。
  SkiaSharp の PDF バックエンド(NuGetで配布されるネイティブビルド)はフォントのサブセット化を行わず
  使用フォントを丸ごと埋め込む(2.88 系・3.x 系の双方で実測確認)ため、Utsushi 自身が描画前に
  使った字形だけのTrueTypeフォントを作って埋め込ませる(要件11.5。下記「日本語の文字」)。
  これにより請求書サンプルは 4.3MB → 約54KB になった。
  文字をベクタのアウトラインとして出力する `PdfTextRendering.Outline` も引き続き選べる
  (PDF内の文字列検索・コピー・テキスト抽出はできなくなる)。
- **日本語の文字(外字・異体字。要件11)**: `Utsushi.Rendering.Fonts` に置く。
  - **同梱フォント**(`BundledFonts`): BIZ UDPゴシック Regular(SIL OFL 1.1、4.6MB)を埋め込みリソースとして持つ。
    `FontResolver` はフォント名を 明示登録 → インストール済み → 同梱 の順に探す。代替フォント名を省略した
    `AllowFallback()` の代替先も、実行環境の既定書体(Linuxでは多くの場合日本語の字形を持たない DejaVu Sans)
    ではなく同梱フォントにする(要件11.1)。これにより日本語フォントの無いサーバー・CIでも日本語を描画できる
    (fontconfig のフォントを0件にした環境で全テストが通ることを確認済み)。同名のフォントがインストール
    されていればそちらを使う(太字の実字形もそちらで得られる)。
  - **日本語のフォント名**(`FontFamilyAliases`): 「ＭＳ Ｐゴシック」↔「MS PGothic」のような主要な日本語フォントの
    名前の対応表で照合する(要件11.6)。全角英数・全角空白は半角にそろえてから比べる。帳票固有の情報ではなく
    OS・Office・主要な無償フォントに共通する名前のため、共通レイヤーに置く。
  - **字形の並び**(`GlyphShaper`): 文字列を書記素クラスタごとに「どの書体のどの字形で描くか」の並び(`GlyphRun`)に
    変換する。計測(`SkiaFontMetricsProvider.MeasureTextWidth`)と描画(`SkiaPdfRenderer`)が同じ結果を使うため、
    外字用の代替フォントで描く文字の送り幅もレイアウト(折り返し・縮小・揃え)に反映される。
    - 外字(要件11.2): セルのフォント → `FontResolverOptions.GlyphFallbackFamilies`(既定: BIZ UDPゴシック →
      IPAmj明朝〈インストールされている場合〉)の順に、字形を持つ最初の書体を使う。太字・斜体の指定は
      代替書体にも引き継ぐ(実字形が無ければ合成)。
    - 異体字(要件11.3): 基底文字 + 異体字セレクタ(IVS/SVS)は、`cmap` format 14 を直接読んで
      (`VariationSequenceTable`。SkiaSharp の文字→字形変換は異体字セレクタを解釈しない)、登録のある最初の書体の
      字形を使う。Default UVS は基底文字の通常の字形、Non-Default UVS は登録された字形番号。どの書体にも登録が
      無ければ要件5.5のエラーとする(`MissingGlyphPolicy.Render` の場合は基底文字だけを描く)。
    - 濁点の合成(要件11.4): 仮名 + U+3099/U+309A の2文字からなるクラスタに限り NFC で合成する。
      NFC を文字列全体にかけると CJK互換漢字(例「﨑」以外の U+F900 台の多く)が統合漢字に置き換わり、
      人名の字形が変わるため、それ以外は正規化しない。
    - 大半の文字列はセルのフォントだけで字形がそろうため、その場合はクラスタ分割をせずに1つの並びにする(計測が頻繁なため)。
  - **描画前の準備**(`RenderContext`): 出力のたびに、全ページの文字列(セル・ヘッダー/フッター・図形内テキスト)を
    字形の並びに変換する。字形の欠落はページを描き始める前にここで検出する(要件5.5)。フォント埋め込みの場合は
    書体ごとに使った字形を集めてサブセットを作り、描画は字形番号で行う(`SKTextBlob`)。
  - **サブセット**(`TrueTypeSubsetter`): `glyf` 形式の書体から、使った字形(と複合字形の部品、.notdef)だけの
    sfnt を組み立てる。字形番号は振り直し、`glyf`/`loca`(長形式)/`hmtx`/`maxp`/`hhea` を作り直す。`cmap` は
    使った文字だけの format 4 + format 12 に、`post` は字形名なしの format 3 にする。`OS/2`・`name`・ヒンティング系
    (`cvt `/`fpgm`/`prep`/`gasp`)は写し、組版用の表(`GSUB`/`GPOS` 等。描画に使わない)や埋め込みビットマップは落とす。
    字形の形・送り幅は元の表をそのまま写すため、見た目・レイアウトは変わらない(テストで送り幅と輪郭の外接矩形の一致を確認)。
    次の書体はサブセット化せず、元の書体を SkiaSharp にそのまま渡す(本サブセット化を導入する前と同じ出力になる)。
    security-reviewer・code-reviewer の実測では、SkiaSharp はそれぞれ括弧内の形で出力した。
    OS/2 `fsType` の用途ビットが制限付きライセンス(0x0002)のみ(Type 3)、サブセット化禁止 0x0100(フォント全体を埋め込み)、
    ビットマップのみ埋め込み可 0x0200(Type 3)、`glyf` を持たない書体(CFF形式。例: Noto Sans CJK は Type 3)、
    可変フォント(`fvar` を持つ。表は既定のインスタンスのものしか取り出せず、太さ等を指定した書体の字形・送り幅が変わるため)、
    表が想定外の形の書体(表の読み出し失敗、`loca` の範囲の重なりで出力が元より大きくなる場合を含む)。
    Type 3 はPDF内の文字列検索がしにくく、異体字(下記)の文字列抽出も欠けるため、外字用の代替フォントを含め
    TrueType形式のフォントを使う。
    フォントの表はインストール済み・利用者登録のファイルから読むため、件数・オフセットは読む前に表の長さと照合し、
    `cmap` format 14 は共有された表を一度だけ読み、読み取る項目の総数を表の長さで打ち切る(security-reviewer指摘。
    300KBの細工した表で約6GBを使わせることができた)。SkiaSharp の `GetTableData` は失敗時に基底の `Exception` を
    投げるため、`TryGetTableData` を使う。サブセットのフォント名(`name` 表)は元のまま写すため、PDFの BaseFont に
    サブセットを示す接頭辞(`ABCDEF+`)は付かない(一般のビューアでは問題ないが、印刷前検査で指摘されうる)。
  - **異体字とPDFの文字列抽出**: `cmap` は1文字に1字形しか対応づけられない。同じ文書で「葛」と「葛」+IVS の両方を
    使う場合は、後から出てきた字形を同じ書体の別のサブセット(バケット)に入れ、どちらも「葛」として抽出できるようにする。
    `cmap` に対応の無い字形があると、SkiaSharp は ToUnicode をU+0000にし、pdftotext などはそこで行の抽出を
    打ち切る(実測)。異体字セレクタそのものは抽出されない(基底文字として抽出される)。
    サブセット化しない書体(上記)では、異体字の字形は `cmap` に対応が無く、抽出から欠ける(Noto Sans CJK で確認)。
  - **太字斜体の合成の判定**: Bold のファイルはあるが Bold Italic が無い書体に太字斜体を求めると、fontconfig は Bold に
    斜体の傾きを付けて返す。この傾きは書体のデータに含まれず、サブセットで作り直すと失われる(layout-fidelity-reviewer指摘の退行)。
    このため、別ファイルの書体が返った場合は OS/2(無ければ head)の書体情報でデータ自体が持つ太字・斜体を読み、
    足りない装飾だけを描画側で合成する(`FontResolver.ReadEmbeddedStyle`)。
  - **厳格モードとの関係**: 厳格モードは「フォント自体が見つからない」場合の方針であり、フォントが見つかったうえで
    字形の無い個々の文字は、厳格モードでも外字用の代替フォントで補う。英字フォント(Arial・Calibri 等)のセルに
    日本語を入れた場合も、日本語の部分は同梱のBIZ UDPゴシックで描かれ、エラーにならない(以前は字形欠落のエラー)。
    Excel はこの場合ブックの東アジア用フォント(テーマの `ea`。例: ＭＳ Ｐゴシック)で描くため、文字幅が異なり、
    折り返し・縮小の結果が変わりうる(請求書の例で 100pt → 112pt)。テンプレートでは日本語を入れるセルに日本語フォントを指定する。
    文字単位の代替を行わない場合は `GlyphFallbackFamilies` を空にする。
  - 制御文字は計測・描画のどちらからも除く(以前は .notdef の送り幅を持っていた。改行などは Layout で取り除き済み)。
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
    IReadOnlyList<DrawingObjectModel> DrawingObjects) // シート上の画像・図形・接続線・グループ(要件9, 10)。drawing.xmlの出現順(=重なり順)。
{
    // セルが無い位置の書式(要件1.10)。GetEffectiveStyle(address) がセル → 行 → 列 → 標準の順に解決する。
    public CellStyle DefaultCellStyle { get; init; }                      // cellXfs[0]
    public IReadOnlyDictionary<int, CellStyle> RowStyles { get; init; }   // customFormat の行
    public IReadOnlyList<ColumnStyleRange> ColumnStyles { get; init; }    // col/@style
}

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
    BorderCallout1, BorderCallout2, BorderCallout3,         // 追加(要件10.14)
    AccentCallout1, AccentCallout2, AccentCallout3,         // 追加(要件10.14)
    AccentBorderCallout1, AccentBorderCallout2, AccentBorderCallout3, // 追加(要件10.14)
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
// 段落内の改行(a:br)はランの Text 中の "\n"、空の段落(ランの無い a:p)は Runs が空の段落で表す(要件10.19)。
public sealed record ShapeTextBody(
    IReadOnlyList<ShapeTextParagraph> Paragraphs,
    VerticalAlignment VAlign,             // a:bodyPr/@anchor
    ShapeTextInsets? Insets = null);      // a:bodyPr/@lIns等。nullは指定なし(ShapeTextInsets.Default)
public sealed record ShapeTextInsets(double LeftPt, double TopPt, double RightPt, double BottomPt); // Default = 7.2/3.6/7.2/3.6
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
    bool HasFormula = false)
{
    public ArgbColor? FormatColor { get; init; }   // 数値書式の色の指定(要件4.12)。WithText で消える
}

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
    IReadOnlyList<PageLayout> Pages, string ReportCode, string SheetName)
{
    public IReadOnlyList<FitIssue> FitIssues { get; init; }  // 文字の収まりの確認(要件13)。描画には使わない
    public bool FitIssuesTruncated { get; init; }            // FitIssues を件数の上限(10,000件)で打ち切ったか
}

// ファサードの CheckFit / CheckFitWithoutDefinition の戻り値
public sealed record FitCheckResult(IReadOnlyList<FitIssue> Issues, bool IsTruncated);

public sealed record FitIssue(
    FitIssueKind Kind,          // OverlapsNeighborValue / CutAtPageEdge / NumberTooWide / Clipped / ExceedsCellHeight
    CellAddress Cell,
    int PageNumber,             // 最初に現れるページ
    string Text,                // 描画する文字列
    bool IsSubstituted,         // 置換キー・セル番地直接指定で値を置き換えたセル
    string? SubstitutionKey,    // 置換キーによる置換ならそのキー
    string Message);            // 説明の文(ログ用)

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
    結合セル範囲の個数上限超過を示す`"TooManyMergedRanges"`(要件2.9)、画像の参照先パートが無いことを示す
    `"MissingImagePart"`(要件6.10)を持つ)
  - `SubstitutionKeyNotFoundException` / `RequiredSubstitutionValueMissingException`(要件2.3, 2.4, 2.12)
  - `InvalidSubstitutionValueException`(要件2.10, 2.16。差し込み値が`null`、改行以外の制御文字、対になっていないサロゲートを含む、
    値の長さ・セル番地直接指定の件数が上限を超える(件数超過の`Target`は`"cellOverrides"`))
  - `InvalidCellOverrideAddressException`(要件2.8, 2.15。セル番地直接指定がA1形式として解釈できない場合、
    表記の異なる複数の番地が同じセルを指す場合。列名はASCIIの英字に限る(`char.IsLetter`だと`"é1"`が無関係なセルになっていた))
  - `NonAnchorMergedCellOverrideException`(要件2.9。セル番地直接指定の対象が結合セル範囲の非アンカー位置の場合)
  - `InvalidExcelFileException`(要件6.1, 6.2, 6.5, 6.6, 6.7, 6.8。`Reason` で非xlsx/破損/パスワード保護/
    ファイルを開けない(存在しない・アクセス不可)/ワークシートが無い/ファイル・共有文字列・
    セル数が上限超過(`TooLarge`)を区別する)
  - `ReportDefinitionSchemaException`(要件6.3。問題のあったプロパティパスを保持する)
  - `LayoutComputationException`(要件2.6, 2.13, 2.14, 6.9 を含む) / `PdfRenderingException` / `FontNotAvailableException`
    - `PdfRenderingException` は、出力先への書き込みの失敗(要件6.13)も表す。`AtomicFileWriter` はパスの正規化・出力先ディレクトリの
      作成・一時ファイルへの書き込み・置き換えで起きた `IOException`・`UnauthorizedAccessException`・`NotSupportedException`・
      不正なパスの `ArgumentException`(`ArgumentNullException` を除く)を、ファサードの `Convert` 系は出力ストリームへの書き込みの
      `IOException` を、`InnerException` に元の例外を入れた `PdfRenderingException` で包む(`IPdfRenderer` の実装によらない)。
    - `FontNotAvailableException` は、帳票を知らない `FontResolver` が帳票コード・シート名なしで送出する。ファサードがこれを捕捉し、
      帳票コード・シート名と処理段階(レイアウト計算中の文字幅の計測なら `Layout`、描画中なら `Rendering`)を補った同じ型の例外で
      包み直す(元の例外は `InnerException`)。
- ファサード(`ReportPdfConverter`)の引数の誤り(必須の引数が`null`、出力先パスが空・空白のみ、`maxDigitWidthPx`が正の有限の数でない)は、
  呼び出し元のプログラムの誤りとして `ArgumentNullException`・`ArgumentException`・`ArgumentOutOfRangeException` を送出し、
  `UtsushiException` 階層には含めない(要件6.14)。入力ファイルの読み込み・レイアウト計算を始める前にすべて検証する
  (以前は出力先の `null` が、時間のかかるレイアウト計算の後に分かっていた)。
  - `MissingGlyphException`(要件5.5。描画する文字の字形がフォントに無い)
- すべての例外は、帳票コード・シート名・セル番地・処理段階(Parsing/Substitution/Layout/Rendering)を構造化プロパティとして保持し、ログ出力時に特定できるようにする(要件6.4)。ただし`MissingGlyphException`は描画命令(`TextCommand`)がセル番地を持たないためセル番地を含まず、代わりに該当文字列(`Text`)で特定する。
- Renderingレイヤーは一時ファイル/一時ストリームに書き込み、正常終了時のみ最終出力先へ確定させる(要件5.4: 不完全PDFを残さない)。
  一時ファイルは出力先と同じディレクトリの `<出力先>.<GUID>.utsushi-tmp` で、失敗時は削除する。プロセスの強制終了など
  後始末が動かない異常終了では残ることがある。

### 信頼できない入力に対する安全弁 一覧

`security-reviewer` が横断確認を行う際の起点として、既知の安全弁(上限定数)を一覧にする。
新しく上限を追加した場合はここに追記し、`src/Utsushi.Parsing/OpenXml/` 配下の
`foreach`/`Elements<...>()` ループを新規に追加した場合は、対応する上限がこの表に
載っているかを確認する。根拠要件は表ごとに異なる点に注意(単一の要件に対応する表ではない)。

| レイヤー / クラス | 定数 | 既定値 | 超過時の挙動 | 根拠要件 |
|---|---|---|---|---|
| Parsing / `OpenXmlWorkbookReader` | `MaxXlsxPackageBytes` | 1 GiB | `InvalidExcelFileException(TooLarge)` | 6.6 |
| Parsing / `OpenXmlWorkbookReader` | `MaxXmlElementDepth` | 256段(全XMLパートの要素の入れ子) | `InvalidExcelFileException(TooLarge)`(DOM構築前に`XmlReader`で流し読みして検査) | 6.7 |
| Parsing / `OpenXmlWorkbookReader` | `MaxXmlPartBytes` | 64 MiB/XMLパート(展開後) | `InvalidExcelFileException(TooLarge)`(同上。`OpenSettings.MaxCharactersInPart`にも同じ値を設定) | 6.7 |
| Parsing / `OpenXmlWorkbookReader` | `MaxXmlElementsPerPart` | 5,000,000個/XMLパート | `InvalidExcelFileException(TooLarge)`(同上。DOMのメモリを抑える) | 6.7 |
| Parsing / `OpenXmlWorkbookReader` | `MaxRelationshipsPerPart` | 10,000件/関係パート | `InvalidExcelFileException(TooLarge)`(`SpreadsheetDocument.Open`より前にZIPを直接流し読みして検査。関係パートはエントリごとに検査し、展開できないものは`Corrupted`にして後ろの関係パートの検査を省かない) | 6.7, 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxRelationshipsPerPackage` | 50,000件(全関係パートの関係の数の合計) | `InvalidExcelFileException(TooLarge)`(同上) | 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxZipEntries` | 10,000個 | `InvalidExcelFileException(TooLarge)`(`ZipArchive`を作る前にZIPの終端レコード(ZIP64を含む)の申告件数で、作った後に実際の件数で検査) | 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxContentTypesBytes` | 4 MiB(`[Content_Types].xml`の展開後) | `InvalidExcelFileException(TooLarge)`(パートではないため`MaxXmlPartBytes`等の対象にならない。`Open`より前に流し読みで検査。エントリ名は`System.IO.Packaging`と同じく`ToUpperInvariant`した名前どうしで比べる(`OrdinalIgnoreCase`では`[Content_Typeſ].xml`を見逃す)) | 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxContentTypesEntries` | 10,000個(`[Content_Types].xml`の`Default`/`Override`の合計) | `InvalidExcelFileException(TooLarge)`(同上。入れ子の深さも`MaxXmlElementDepth`で検査) | 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxXmlElementsPerPackage` | 8,000,000個(全XMLパートの要素の数の合計) | `InvalidExcelFileException(TooLarge)`(`MaxXmlElementsPerPart`の直前まで詰めたパートを複数並べる入力への対策) | 6.11 |
| Parsing / `OpenXmlWorkbookReader` | `MaxCellTextLength` | 32,767文字(共有文字列・インライン文字列。Excel自体の上限) | `InvalidExcelFileException(TooLarge)` | 6.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxHeaderFooterTextLength` | 1,024文字/ヘッダー・フッター | `InvalidExcelFileException(TooLarge)` | 6.8 |
| Parsing / `OpenXmlWorkbookReader` | `MinPrintScalePercent`/`MaxPrintScalePercent` | 10〜400%(Excelと同じ) | 範囲外の拡大縮小率を丸める(例外化なし) | 6.9 |
| Parsing / `StyleTable` | `MaxNumberFormatCodeLength` | 255文字(Excel自体の上限) | その数値書式を読み取らず既定書式にフォールバック(例外化なし) | 6.8 |
| Parsing / `NumberFormatter` | `MaxCacheableFormatLength` / `MaxCacheEntries` | 255文字以下の書式コードのみ / 1,024件(プロセス内で共有する解析結果のキャッシュ) | 長すぎる書式コードはキャッシュせず都度解析する。件数が上限に達したらキャッシュを空にして登録し直す(例外化なし) | 設計判断(静的キャッシュの肥大化防止、要件番号なし) |
| Parsing / `OpenXmlWorkbookReader` | `MaxRowHeightPt` | 409pt(Excel自体の上限。`row/@ht`・`sheetFormatPr/@defaultRowHeight`) | 範囲外・NaN・無限大・負は無視し既定の行高(既定行高が不正なら`DefaultRowHeightPt` = 15pt)にする(例外化なし) | 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MaxColumnWidthChars` | 255文字(Excel自体の上限。`col/@width`・`defaultColWidth`) | 範囲外・NaN・無限大・負の`col/@width`は幅の指定が無いものとして、`defaultColWidth`は無いものとして`baseColWidth`から求める(例外化なし) | 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MaxPageMarginInches` | 100インチ/余白1辺 | 範囲外・NaN・無限大・負はその辺の既定の余白に戻す(例外化なし。用紙に収まらない余白の検出はLayout) | 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MinFontSizePt`/`MaxFontSizePt` | 1〜409pt(セルのフォント。Excelと同じ) | 範囲外・NaN・無限大は既定のサイズに戻す(例外化なし) | 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MinShapeFontSizeHundredths`/`MaxShapeFontSizeHundredths` | 100〜400,000(1/100pt。図形の文字の`a:rPr/@sz`、1〜4,000pt。ECMA-376の`ST_TextFontSize`) | 範囲外は既定のサイズに戻す(例外化なし) | 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MaxFirstPageNumber` | 1,000,000,000 | 超える値・解釈できない値は先頭ページ番号の指定なし(1から数える)とする(例外化なし。`&P`・`&N`の計算がintを溢れないように) | 3.15, 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapeTextInsetPt` | 5,000pt/辺(`a:bodyPr`の`lIns`等) | 範囲外・解釈できない辺はその辺だけ既定値(左右7.2pt・上下3.6pt)にする(例外化なし) | 10.19, 6.12 |
| Parsing / `OpenXmlWorkbookReader` | `MaxColumnExpansionsPerSheet` | 65,536列(`<col>`の`min`〜`max`の展開の合計) | `InvalidExcelFileException(TooLarge)` | 6.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxPageBreaksPerSheet` | 1,026件(行・列それぞれ。Excel自体の上限) | `InvalidExcelFileException(TooLarge)` | 6.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxPrintAreasPerSheet` | 1,000個 | `InvalidExcelFileException(TooLarge)` | 6.8 |
| Layout / `ReportLayoutEngine` | `MaxPrintRangeCells` | 2,000,000(印刷範囲ごとの(行数+タイトル行数)×(列数+タイトル列数)の合計) | `LayoutComputationException`(`SheetGrid`を作る前に判定) | 6.9 |
| Layout / `StyledBlankPositionCounter` | `MaxStyledBlankPositions` | 200,000か所(文書全体で、ファイルにセルが無い位置を行・列・標準の書式で塗りつぶし・罫線を描いた回数。印刷タイトルの繰り返しも数える) | `LayoutComputationException` | 1.10(セルの描画命令の数は従来 `MaxCellsPerSheet` で抑えられていたが、`col` 要素1つで 16,384 列に効く行・列の書式は、小さなファイルで `MaxPrintRangeCells` に比例する描画命令・PDFサイズを作らせられる。security-reviewer指摘) |
| Layout / `FitIssueCollector` | `MaxIssues` | 10,000件(文字の収まりの確認の結果) | 打ち切り、`PagedLayout.FitIssuesTruncated`(`FitCheckResult.IsTruncated`)を立てる(例外化なし) | 13.4 |
| Layout / `ReportLayoutEngine` | `MaxPagesPerDocument` | 5,000ページ | `LayoutComputationException`(ページを組み立てる前に判定) | 6.9 |
| Parsing / `OpenXmlWorkbookReader` | `MaxSharedStringCount` | 200,000件 | `InvalidExcelFileException(TooLarge)` | 6.6 |
| Parsing / `OpenXmlWorkbookReader` | `MaxCellsPerSheet` | 500,000(シート内セル数、および行番号の上限を兼ねる) | `InvalidExcelFileException(TooLarge)` | 6.6 |
| Parsing / `OpenXmlWorkbookReader` | `MaxImagesPerSheet` | 50枚/シート | `UnsupportedWorkbookElementException(TooManyImages)` | 9.6 |
| Parsing / `OpenXmlWorkbookReader` | `MaxImageDataBytes` | 10 MiB/画像 | `UnsupportedWorkbookElementException(ImageTooLarge)` | 9.6 |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapesPerSheet` | 50個/シート(図形・接続線・グループ合計) | `UnsupportedWorkbookElementException(TooManyShapes)` | 10.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapeTextLength` | 2,000文字(段落・ランを連結した文字数。段落の区切りも1文字と数える) | `UnsupportedWorkbookElementException(ShapeTextTooLong)` | 10.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxShapeNestingDepth` | 5段 | `UnsupportedWorkbookElementException(GroupNestingTooDeep)` | 10.8 |
| Parsing / `OpenXmlWorkbookReader` | `MaxGradientStopsPerFill` | 64個/塗りつぶし | 打ち切り(超過分は無視、例外化なし) | 設計判断(DoS対策、要件番号なし) |
| Parsing / `OpenXmlWorkbookReader` | `MaxMergedRangesPerSheet` | 1,000件/シート | `UnsupportedWorkbookElementException(TooManyMergedRanges)` | 2.9 |
| Parsing / `StyleTable` | `MaxStyleTableEntries` | 10,000件(フォント/塗りつぶし/罫線/数値書式/cellXfs各々) | 打ち切り(超過分は既定書式にフォールバック、例外化なし) | 要件6.6の対象外(6.6補足文の通り、既定書式フォールバックで吸収できるため意図的に例外化していない) |
| Parsing / `ColorResolver` | `MaxIndexedColorCount` | 10,000件 | 打ち切り(超過分は呼び出し元のfallback色、例外化なし) | 要件6.6の対象外(上記と同様の理由) |
| Layout / `PageCommandBuilder` | `MaxSpanCells` | 4,096セル | 結合セル・可視範囲の走査を打ち切り(例外化なし) | 設計判断(DoS対策、要件番号なし) |
| Layout / `PageCommandBuilder` | `MaxDrawingObjectDimensionPt` | 5,000pt | 描画オブジェクトの寸法をクランプ(例外化なし) | 設計判断(DoS対策、要件番号なし) |
| Rendering / `SkiaPdfRenderer` | `MaxDecodedImageDimensionPx` | 4,096px | `PdfRenderingException`(デコード前に宣言サイズを検査) | 設計判断(DoS対策、要件番号なし) |
| Substitution / `CellSubstitutor` | `MaxValueLength` | 32,767文字/値(UTF-16のコード単位数。Excelのセルの上限。置換キー経由・セル番地直接指定とも) | `InvalidSubstitutionValueException` | 2.16 |
| Substitution / `CellSubstitutor` | `MaxCellOverrideCount` | 10,000件/変換(セル番地直接指定) | `InvalidSubstitutionValueException`(`Target = "cellOverrides"`。番地の解釈より前に判定) | 2.16 |
| Layout / `HeaderFooterParser` | `MaxPageOffsetDigits` | 9桁(`&P+n`/`&P-n`の数字) | それより後ろの数字は加減算に使わず文字のまま残す(例外化なし。桁あふれ防止) | 3.16 |

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
- ~~PDFのフォントサブセット化~~ → SkiaSharp のネイティブビルドが対応していないため、Utsushi 自身が
  描画前に使った字形だけのTrueTypeフォントを作って埋め込ませる(要件11.5、`TrueTypeSubsetter`)。
  請求書サンプルで 4.3MB → 約54KB。外字用のIPAmj明朝(46MB。そのまま埋め込むと2文字で約31MB)も同様に数KBになる。
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

- **差し込み文字列の配置で未対応のExcel書式**(要件4.4): 主用途(宛先・住所・氏名・担当者の差し込み)の
  品質確認(タスク21)で、以下が未対応であることを確認した。いずれもテンプレート側で避けられるため、
  対象帳票で必要になった時点で対応する。
  - 文字の方向・回転(`textRotation`。縦書き`255`を含む)は読み取っておらず、横書きで出力される。
    サポート外要素としての検出対象にもなっていない。
  - 横位置「均等割り付け」「両端揃え」「繰り返し」は左揃え、「選択範囲内で中央」は自セル内の中央揃え、
    縦位置「均等割り付け」「両端揃え」は中央揃えとして扱う。インデントは常に左側に加算される。
  - はみ出し表示(`overflow`)は、Excelと違い隣のセルに値があっても止まらず重なって描画される(描画は変えず、
    要件13の文字の収まりの確認で重なる箇所を検出できるようにした)。
  - 禁則処理・英単語単位の折り返しは行わない(書記素クラスタ単位の単純な幅基準)。Excelとは行数が
    1行ずれることがあり、要件2.14の判定もこの行数に基づく。
  - 横方向の欠落は例外にしない(要件13の文字の収まりの確認で検出できる)。`clip`のセルでは長い氏名・住所が右側で切り詰められ、`overflow`のセルでは
    このページの本文の範囲(余白の内側。要件4.9)の外にはみ出した部分が欠ける。結合セルの文字ははみ出さず結合範囲で切れる。要件2.14は縦方向(折り返し行)のみを対象とする。
  - 複数行のヘッダー/フッターは未対応(改行を取り除いて1行に配置する)。
  - 文字の合成(シェーピング)は、異体字セレクタ(要件11.3)と仮名の濁点・半濁点(要件11.4)に限って行う。
    それ以外の結合文字(例: ラテン文字のアクセント記号)や字形の置き換え(`GSUB`)は行わない。
- **大量発行時の性能**: 変換のたびに`definition.json`と`.xlsx`全体を解析し直す(キャッシュ無し)。
  折り返しは1文字追加するたびに行頭からの幅を測り直すため、処理量は1段落の文字数の2乗に比例する。
  開発環境の実測では、同一プロセス内で1つのコンバータを使い回して請求書サンプルを20件連続変換した場合に
  1件あたり約0.1秒(初回のフォント解決を除く。フォント埋め込み・サブセット化込み、約60KB。Releaseビルド)。
  発行件数が問題になった時点で、テンプレート解析結果のキャッシュ(`WorkbookModel`は不変モデルのため
  使い回せる)を検討する。

- **全体監査(タスク28)で残した未検証事項・未対応事項**(layout-fidelity-reviewer指摘):
  - Excelの実機(PDF出力)との突き合わせが未実施: 先頭ページ番号を指定した場合の`&N`と奇数/偶数ページの選び方、
    はみ出し表示をページ本文(余白の内側)で切り取ること、ページ中央の対象に印刷タイトルを含めること、
    印刷範囲より下(右)にあるタイトルを付けないこと、太い実線の罫線がT字に交わる箇所の端の形、
    回転(75〜89・92)・横送りの用紙と横向き(`landscape`)の組み合わせ。
  - 未対応: 数値がセル幅に収まらない場合の`####`表示(数値の文字列をそのまま配置する。要件13の確認で検出できる)、45〜135度回転した画像の
    配置矩形(この角度ではアンカーが幅と高さを入れ替えた矩形を表すとされるが、Utsushiは入れ替えずに扱う。実機では未確認)、最終行(列)が非表示の結合範囲で外周の罫線の辺が欠けること、
    和暦の「元年」表示。
- **改ページをまたぐ結合セルの文字位置**: 結合範囲がページ境界をまたぐ場合、現状は
  「そのページに見えている部分」を1つの矩形として扱い、その中にテキストを配置する。
  Excel は結合範囲全体を基準に文字を配置したうえでページ境界で切り取るため、
  両ページに同じ文字列が現れる点が異なる。印刷タイトル行にある結合セル
  (各ページに繰り返すことが正しいもの)とは区別が必要なため、対象帳票で
  「本文中の結合セルが改ページをまたぐ」ケースが出てから実装する。
  なお罫線については、範囲の本当の上端/下端/左端/右端がそのページに含まれる場合のみ
  対応する辺を描画するよう対応済み(改ページの切れ目には罫線を出さない)。文字位置のみ未対応。
- **外字・異体字対応の残る限界**(要件11):
  - 外字用の代替フォントで描いた文字は、セルのフォントと書風が異なることがある(例: ゴシック体の宛名の中の
    「𠮷」だけが、IPAmj明朝の明朝体になる)。IPAmj明朝は明朝体しか無いため。社内に同じ書風の外字フォントがあれば、
    `GlyphFallbackFamilies` の先頭に登録する。
  - CFF形式(`glyf` を持たない)のフォント(例: Noto Sans CJK)はサブセット化できず、SkiaSharp が Type 3 として出力する。
    外字用の代替フォントにはTrueType形式のフォントを使う。
  - 異体字の字形の文字列抽出は基底文字になり、異体字セレクタは抽出されない。
  - 縦書き用の字形(`vert`)には対応しない(縦書き自体が未対応)。
  - 行の高さ・縦位置はセルのフォントのメトリクスだけで決める。外字用の代替フォントの em ボックスの設計が大きく
    異なる場合(社内外字フォントなど)、その文字だけがクリップ矩形からはみ出して切れる可能性がある(未検証)。
  - 合成済みの字形が無い仮名 + 濁点・半濁点(例「き」+ U+309A)は、2つの字形として並べて描く(Excel は重ねて表示する)。
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
  帳票定義なしモード(要件12.5)では帳票コードの代わりに文書名(ファイル変換では拡張子を除いた入力ファイル名)を
  展開する。Excelの`&F`は拡張子付きのファイル名を出すため、この点はExcelと異なる。
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
  フローチャート記号7種・吹き出し4種(雲形・引き出し線1〜3本)・接続線5種を追加し、
  その後、要件10.14で線吹き出しを12種(`callout`/`borderCallout`/`accentCallout`/`accentBorderCallout`の1〜3)に広げたが、
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
  帳票定義なしモード(要件12)の追加時に再評価し、security-reviewerが実測した: 「数万段」「理論上」という
  上の見積もりは誤りで、スタック1MB(Windowsのメインスレッドの既定)では**深さ5,000段・4.2KBの`.xlsx`**で
  `OpenXmlCompositeElement.Populate`の再帰によりスタックオーバーフローし、**呼び出し元のプロセスごと**落ちる
  (Linux既定の8MBでは落ちないが、処理時間が深さの二乗で増え、200,000段(75KB)で400秒以上応答しない)。
  また、定義ありモードでは登録済みのシート名に一致するファイルしか事実上通らないのに対し、定義なしモードは
  任意の`.xlsx`を受け付けるため、社内の担当者が社外から受け取ったファイルをそのまま変換する経路が現実に生じる。
  このため「帳票定義の有無は入力の信頼性を変えない」とは言えず、リスクは定義なしモードで大きくなる。
  タスク25で、SDKのDOMに触れる前に、パッケージ内の全XMLパートを`XmlReader`で流し読みして要素の深さ
  (`MaxXmlElementDepth`)とパートの大きさ(`MaxXmlPartBytes`)を検査するようにした(要件6.7)。流し読みは
  再帰しないため、深いネストでもスタックを消費しない。この対策で上記のスタックオーバーフローは起きなくなった。
- **放射状グラデーションの`a:path`種別**(要件10.6): `a:path type="circle"`のみを
  厳密に扱い、`"rect"`/`"shape"`は同じ放射状近似にフォールバックする。矩形・図形に
  沿った塗りつぶしの厳密な再現は行わない。
- **画像・図形の上限がシート単位でありワークブック単位の合算上限が無い**(security-reviewer指摘):
  `MaxImagesPerSheet`/`MaxShapesPerSheet`はシートごとにリセットされるカウンタであり、
  ワークブック全体でシートをまたいだ合算上限は無い。`ReportPdfConverter.Convert`は
  常に帳票定義の`sheetName`1枚に処理対象を絞り(`WorkbookReadOptions.SheetNameFilter`)、帳票定義なしモード(要件12)も
  アクティブシート1枚に絞る(`WorkbookReadOptions.ActiveSheetOnly`)ため
  現状の呼び出し経路では実害は無いが、`IWorkbookReader`/`OpenXmlWorkbookReader`は
  `public`であり、`SheetNameFilter`を指定しない(全シート読み取り)呼び出し方をする
  将来のコードが現れた場合はシート数倍の積み上げに対する上限が無い。対象帳票で
  実際に必要になった時点でワークブック単位の合算上限を検討する。
  なお、同じ名前の`<sheet>`を大量に並べると名前での絞り込みを素通りして同じシートを何度も読めてしまう問題は、
  絞り込んだ1枚を読んだ時点でループを打ち切ることで解消した(タスク24.9)。
- **ファイルの内容だけで処理量が決まる経路の上限漏れ**(帳票定義なしモード追加時のsecurity-reviewer指摘):
  タスク25で対応した(要件6.7〜6.10、上の安全弁一覧)。当時残っていた「1ページ内のセルごとに結合範囲を線形に探す
  (`SheetModel.FindMergedRange`)ため、処理量が印刷範囲のセル数×結合範囲の個数に比例する」制約は、タスク28で
  セル→結合範囲の索引(`MergedCellIndex`)を印刷範囲ごとに作ってページ間で使い回すことで解消した。
  その後の security-reviewer の実測で見つかった経路(型として不正な属性値による SDK の例外の漏れ、検証に失敗した
  画像が枚数に数えられない、印刷タイトルでの上限の迂回、文字列の長さ、DOM の要素数、関係パートの数)も、
  上の安全弁一覧のとおり対応した(タスク25.6)。属性値の例外は、個々の読み取り箇所ではなく `ReadCore` で
  `FormatException`/`OverflowException` をまとめて `InvalidExcelFileException(Corrupted)` に読み替える。
  タスク28の全体監査では、パッケージ構造に残っていた迂回経路(ZIPエントリの数、関係の数の合計、`[Content_Types].xml`、
  パッケージ全体の要素の数、壊れた関係パートの後ろの関係パートが検査されない)を塞いだ(要件6.11)。パートのルート要素が
  想定と違う場合に OpenXml SDK が投げる `InvalidDataException` も `Corrupted` に読み替える。
- **図形の色・反転・矢印の未検証事項・既知の差分**(要件10.15〜10.18。layout-fidelity-reviewer指摘):
  (1) 上下反転した図形の文字を180°回す挙動と、矢印の大きさ(線の太さの2/3/5倍、元の太さの最小1pt)は、Excelで保存した
  ファイルのPDF出力との実測での突き合わせが未実施。LibreOfficeの実装では元の太さの最小が約2ptで、開いた矢印(`arrow`)の
  倍率が大きいとされており、細い線の矢印はExcelより小さく描かれている可能性がある。対象帳票で矢印・上下反転を使う場合は、
  実物と見比べて調整する。(2) `fillRef`の2・3番目(Office既定テーマではグラデーション)はスタイルの色の単色で近似するため、
  「グラデーション」系のスタイルの図形は実際と色合いが異なる。(3) `fontRef`は色だけを使い、書体(`minor`/`major`)は使わない。
  (4) 未対応の色の指定(`a:prstClr`/`a:scrgbClr`/`a:hslClr`)と修飾(`satMod`/`hueMod`等)は通知なく無視される。
