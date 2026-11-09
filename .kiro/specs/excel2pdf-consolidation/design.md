# 設計書: 入口の Excel2Pdf への一本化と診断の強化

対象要件: `.kiro/specs/excel2pdf-consolidation/requirements.md`
前提となる設計: `.kiro/specs/excel-report-pdf-conversion/design.md`(以下「中核の設計」。本書に書かない点は中核の設計のまま)

## 概要

変換の流れ(Parsing → 変換設定の合成 → Substitution → Layout → Rendering)は変えない。主な変更は次の4つである。

1. **入口**: 公開する入口を `Excel2Pdf` と `Excel2PdfEngine` に絞る。`ReportPdfConverter` は internal の
   変換パイプライン(`ConversionPipeline`)に縮める。
2. **帳票定義**: JSONの帳票定義を廃止する。`ReportDefinition` は、帳票定義なしの変換で合成していた
   「変換設定」だけを持つ型(`ConversionSettings`)に縮める。
3. **未対応要素**: Parsing は扱いの指定(無視/エラー)を受け取らなくなる。常に要素を読み飛ばしたうえで記録し、
   扱いの判断はファサードが行う。
4. **フォントの代替**: `FontResolver` が代替の結果を問い合わせられるようにする。ファサードはレイアウトの結果に
   現れたフォントについて問い合わせ、診断に載せる。

```
Excel2Pdf ──(設定した値・Excel2PdfOptions)──▶ ConversionPipeline(internal。旧 ReportPdfConverter)
  1. Parsing: IWorkbookReader.Read ──▶ WorkbookModel(+ UnsupportedElements)
  2. [保存のみ] 扱いがエラーで UnsupportedElements が1件以上 → UnsupportedWorkbookElementException(ここで止める)
  3. ConversionSettings(合成) + IReportModelBuilder ──▶ ReportModel
  4. Substitution: ApplyCellOverrides / ApplyNumericCellOverrides
  5. Layout: Compute(report, checkFit) ──▶ PagedLayout(+ FitIssues)
  6a. [保存] Rendering: SkiaPdfRenderer ──▶ PDF
  6b. [診断] PagedLayout のフォント → FontResolver.GetSubstitution ──▶ Excel2PdfDiagnostics
```

## 1. 公開API(ファサード `Utsushi`)

### 1.1 `Excel2PdfEngine`(要件1.2, 1.4)

```csharp
public sealed class Excel2PdfEngine : IDisposable
{
    public static Excel2PdfEngine Create(FontResolverOptions? fonts = null, PdfRenderOptions? render = null);
    public void Dispose();
}
```

- 現在の `ReportPdfConverter.CreateDefault(FontResolverOptions?, PdfRenderOptions?)` の役割を引き継ぐ。
  内部に `ConversionPipeline` を1つ持ち、`FontResolver` を Layout(`SkiaFontMetricsProvider`)と描画で共有する。
- スレッドセーフである。`ConversionPipeline` は変換ごとの状態を持たず、`FontResolver` は内部でロックする(現在の `ReportPdfConverter` と同じ)。
  `Excel2Pdf` の `Dispose` は、渡されたエンジンを解放しない(要件1.2)。
- 変換のメソッドは公開しない。`Excel2Pdf` はエンジンの internal メンバーを通じて変換する。
- 部品(`IWorkbookReader`・`IReportLayoutEngine`・`IPdfRenderer`・`FontResolver`)を差し替えて作るコンストラクタは internal にする。
  現在 `ReportPdfConverter` の公開コンストラクタで行っている差し替え(ゴールデンテストの決定的なフォントメトリクス・固定時刻)は、
  既存の `InternalsVisibleTo`(`Directory.Build.props`。`<プロジェクト名>.Tests` と `Utsushi.Golden.Tests`)でテストにだけ許す。
  このコンストラクタは、フォントの代替の問い合わせ(3)のために `FontResolver` を明示的に受け取る
  (現在の `ReportPdfConverter` は `IDisposable? ownedResources` としてしか受け取っていない)。
- `ReportPdfConverter` という名前を残さないのは、変換のメソッドを持たない型に「Converter」の名前は合わず、
  帳票定義を前提にした旧APIと区別するためである。

### 1.2 `Excel2Pdf`(要件1.1, 1.3, 4)

現在のコンストラクタは9つある(帳票コード必須が3つ、コンバータと任意の帳票コードを受け取るものが3つ、それ以外が3つ)。
このうち帳票コード必須の3つを削り、コンバータを受け取る3つをエンジンを受け取る形に置き換えて、次の6つにする。

```csharp
public Excel2Pdf(string xlsxPath, Excel2PdfOptions? options = null);
public Excel2Pdf(byte[] xlsxContent, Excel2PdfOptions? options = null);
public Excel2Pdf(Stream xlsxStream, Excel2PdfOptions? options = null);
public Excel2Pdf(string xlsxPath, Excel2PdfEngine engine, Excel2PdfOptions? options = null);
public Excel2Pdf(byte[] xlsxContent, Excel2PdfEngine engine, Excel2PdfOptions? options = null);
public Excel2Pdf(Stream xlsxStream, Excel2PdfEngine engine, Excel2PdfOptions? options = null);
```

| メンバー | 変更 |
|---|---|
| `SetText`・`SetValue`(5種)・`Clear()`・`Clear(cell)`・`Save(path)`・`Save(stream)`・`ToPdfBytes()`・`Dispose()`・`XlsxPath` | そのまま |
| `SetField`・`ClearField`・`ReportCode` | 削除 |
| `CheckFit()` | 削除し、`Check()` に統合する(要件4.5) |
| `Check()` | 新設。`Excel2PdfDiagnostics` を返す(要件4.1) |

- エンジンを渡した場合に `Excel2PdfOptions.Fonts`・`Render` を指定すると `ArgumentException`(現在のコンバータを渡す場合と同じ)。
- `Save`・`ToPdfBytes` は、未対応要素の扱いがエラーで未対応要素が1件以上あれば、`UnsupportedWorkbookElementException` を送出する(要件2.2)。
  判定は概要の図の手順2で行う。

### 1.3 `Excel2PdfOptions`(要件2.1)

| プロパティ | 変更 |
|---|---|
| `Fonts`・`Render`・`DocumentName`・`MaxDigitWidthPx`・`SheetName` | そのまま(帳票定義ありの場合の制約の説明を除く) |
| `UnsupportedElements`(`UnsupportedElementPolicy`) | 新設。既定値は `UnsupportedElementPolicy.Error` |

```csharp
public enum UnsupportedElementPolicy
{
    Error = 0,   // 既定。保存の時点で、未対応要素の一覧を持つ例外を送出する
    Ignore = 1,  // 未対応要素を出力せずに変換を続ける
}
```

- 既定値は `init` の初期値で `Error` にする。そのうえで列挙の0も `Error` にし、`default(UnsupportedElementPolicy)` が無視にならないようにする。
- 現在 `Utsushi.ReportDefinitions` に同名の列挙(`Ignore = 0`)がある。帳票定義を削除するまで(tasks.md のタスク7)は両方が存在する。
  `ReportPdfConverter`(名前空間 `Utsushi`)の中で旧い列挙を参照している箇所は、新しい列挙と取り違えないよう完全修飾名で書く。

### 1.4 `Excel2PdfDiagnostics`(要件4)

```csharp
public sealed record Excel2PdfDiagnostics(
    IReadOnlyList<FitIssue> FitIssues,
    bool FitIssuesTruncated,
    IReadOnlyList<UnsupportedElement> UnsupportedElements,
    bool UnsupportedElementsTruncated,
    IReadOnlyList<FontSubstitution> FontSubstitutions)
{
    public bool HasIssues => FitIssues.Count > 0 || UnsupportedElements.Count > 0 || FontSubstitutions.Count > 0;
}

public sealed record FontSubstitution(string RequestedFamily, bool Bold, bool Italic, string ActualFamily);
```

- `Check()` は未対応要素の扱いによらず、未対応要素を例外にしない(要件4.2)。
  手順2を飛ばして最後まで計算するため、それ以外のエラーは、扱いを無視にした場合の `Save` と同じになる(要件4.4)。
- フォントの代替は、変換で使ったフォントの数(フォント名・太字・斜体の組)を超えないため、打ち切りを設けない。
- `FitCheckResult`(`src/Utsushi.Layout/Model/FitIssue.cs`)は削除する(`Excel2PdfDiagnostics` に置き換える)。
- `FitIssue.SubstitutionKey` は削除する(置換キーの廃止)。`IsSubstituted` は「`SetText`・`SetValue` で設定したセルか」の意味で残す。

## 2. 未対応要素(要件2)

### 2.1 モデル(`Utsushi.Core`)

例外(`Utsushi.Core`)から参照するため、モデルは `Utsushi.Core` に置く。`Utsushi.Core` は値型と例外だけを持つ方針(`.kiro/steering/structure.md`)に合う。

```csharp
public enum UnsupportedElementKind
{
    Chart,                    // グラフ(グラフを参照する図表枠)                                      要件2.7(a)
    UnsupportedDrawingObject, // グラフ以外の図表枠・その他の非対応の描画オブジェクト・解決できない描画パート (a)
    UnsupportedAnchor,        // 絶対位置(absoluteAnchor)で配置された描画オブジェクト                     (a)
    UnsupportedShapePreset,   // 非対応プリセット・custGeom の図形・接続線                                 (b)
    UnsupportedGroupChild,    // 非対応の子要素を含むグループ(グループ全体を出力しない)                   (c)
    UnsupportedImageFormat,   // 非対応形式、または内容が形式と一致しない画像                              (d)
    MissingImagePart,         // 参照先が無い・空・画像でない画像                                          (d)
    LimitExceeded,            // 枚数・大きさ・文字数・入れ子・結合セル数の上限超過                          (e)
    ExternalReference,        // 外部参照                                                                  (f)
    LegacyDrawing,            // コメント・フォームコントロールなどのVML                                    (f)
    MalformedDrawingObject,   // 位置・大きさが欠けている・壊れていて読み飛ばす描画オブジェクト            (g)
}

public sealed record UnsupportedElement(
    UnsupportedElementKind Kind,
    string? SheetName,        // 外部参照は null(ブック全体)
    CellAddress? Anchor,      // アンカーの左上のセル。分からなければ null
    string? Detail,           // プリセット名・content-type・上限の種類と値など
    string Message);          // 人が読む説明
```

- 現在の種類の文字列(`ElementKind`)のうち、上限の超過を表すもの(`TooManyImages`・`TooManyShapes`・`TooManyMergedRanges`・`ImageTooLarge`・
  `ShapeTextTooLong`・`GroupNestingTooDeep`)は `LimitExceeded` にまとめる。どの上限かは `Detail` に入れる(例: `"TooManyShapes: 上限 2000"`)。
- 記録の単位(要件2.5)は次のとおりにする。
  - 枚数・個数の上限(`TooManyImages`・`TooManyShapes`・`TooManyMergedRanges`)は、シートと上限の種類ごとに1件だけ記録する。
    超えた後の要素を1つずつ記録すると、上限が意味をなさなくなるためである。
  - 1つの要素の上限(`ImageTooLarge`・`ShapeTextTooLong`・`GroupNestingTooDeep`)は、要素ごとにアンカーとともに記録する。
    要素の数は枚数・個数の上限で抑えられているため、件数は増えすぎない。
- `Message`・`Detail` には、セル・図形の文字列を入れない(要件2.4)。ファイルから読んだ値(content-type・プリセット名など)は任意の文字列になりうる。
  そこで、制御文字を取り除き、64文字で切り詰めてから入れる(処理ログに流れる値の大きさを抑えるため)。
- 現在 1623 行で、グループ内の非対応の子要素の種類を `"UnsupportedShapePreset"` としているのは誤りである。これを `UnsupportedGroupChild` に直す。

### 2.2 Parsing の変更

- `WorkbookReadOptions.UnsupportedElementBehavior` と列挙 `UnsupportedElementBehavior` を削除する。
  読み込みは常に、現在の「無視」の動作(要素を読み飛ばす)をする。
- `OpenXmlWorkbookReader` で扱いの指定を見ている箇所はすべて書き換える。内訳は、`== Error` で例外を投げる23か所と、
  `!= Error` で検知を省く2か所(`DetectUnsupportedWorkbookElements`・`DetectUnsupportedElements`。2662〜2730行)である。
  前者は、読み込み1回ぶんの収集器 `UnsupportedElementCollector`(internal)への `Add` に置き換え、無視の側の処理(`null`/`false` を返す、`break` する)は残す。
  後者は、検知を常に行う。
- **描画オブジェクトの検知はアンカーごとに行う**。現在のシート単位の真偽の判定(`ChartParts.Any()`・`HasUnsupportedDrawingObject`)では、
  アンカーも件数も得られないためである。描画パートのアンカー(`twoCellAnchor`・`oneCellAnchor`・`absoluteAnchor`)を1つずつ見て、次のように1件ずつ記録する。
  - `absoluteAnchor` → `UnsupportedAnchor`(中身が画像・図形でも)
  - グラフを参照する図表枠(`graphicFrame` の `c:chart`) → `Chart`
  - それ以外の図表枠と、4種(画像・図形・接続線・グループ)以外の要素 → `UnsupportedDrawingObject`

  同じアンカーを `Chart` と `UnsupportedDrawingObject` の両方で記録しない(グラフは図表枠のアンカーに置かれるため)。
  シートから参照されているが解決できない描画パートは、アンカーの無い `UnsupportedDrawingObject` とする。
  `legacyDrawing` は、VML を解析しないため、シートごとにアンカーの無い1件とする。
- **黙って読み飛ばしている経路**(要件2.7(g))は、行番号で列挙せず、「描画オブジェクトを読む処理(`Read*`)が、例外も記録もせずに値を返さずに読み飛ばすすべての経路」と定める。
  現時点で分かっているものは次のとおりである。
  - アンカーの位置の破損(1077行)
  - 画像の大きさを解決できない(1213行)
  - 図形の大きさを解決できない(1394〜1396行)
  - グループの `chOff`/`chExt` の欠落(1535行)
  - 接続線・グループにある同じ形の経路

  これらは `MalformedDrawingObject` として記録する。
  `embedId` が空(1241行)・画像以外のパートを参照(1267行)は、参照先の問題なので `MissingImagePart`(要件2.7(d))とする。
  漏れが無いことは、security-reviewer の横断確認(`src/Utsushi.Parsing/OpenXml/` 全体)で確かめる。
- 収集器は件数の上限 `MaxUnsupportedElements = 10_000` を持つ。上限に達したら記録をやめ、打ち切りの印を立てる(要件2.5)。
  上限の比較は `FitIssueCollector` と同じ形にし、上限を渡せる internal のオーバーロードでテストする(`.kiro/steering/tech.md`「テスト戦略」)。
- 結果は `WorkbookModel` に載せる。

  ```csharp
  public sealed record WorkbookModel(
      IReadOnlyList<SheetModel> Sheets,
      FontStyle DefaultFont,
      IReadOnlyList<UnsupportedElement> UnsupportedElements,
      bool UnsupportedElementsTruncated);
  ```

  変換対象のシート(`ActiveSheetOnly`・`SheetNameFilter` で読む1枚)と、ブック全体の要素(外部参照)だけを含む。読まないシートの要素は含まない
  (帳票定義ありの経路も、`SheetNameFilter` で1枚だけを読むため同じである)。
- 安全弁(中核仕様の要件6.6〜6.9・6.11)の検査は、これまでどおり例外で直ちに中止する(要件2.6)。収集の対象にはしない。

### 2.3 例外

`UnsupportedWorkbookElementException` を次のように変える。

```csharp
public sealed class UnsupportedWorkbookElementException : UtsushiException
{
    public IReadOnlyList<UnsupportedElement> Elements { get; }
    public bool IsTruncated { get; }
    // ElementKind(string)は削除する(tasks.md のタスク2で、参照しているテストとともに)
}
```

- メッセージは「未対応の要素が N 件あるため中止します(UnsupportedElements を Ignore にすると、出力せずに変換します)」とし、
  先頭の20件を1行ずつ続ける。21件目以降は「ほか M 件」とする。
- `ToDiagnosticString()` には、要素の一覧をすべて(最大10,000件)含める。処理ログに残し、テンプレートを一度で直せるようにするためである(要件2のユーザーストーリー)。
  1件は1行で、セルの文字列を含まないため、1件あたりの大きさは抑えられている。

### 2.4 帳票定義ありの経路の移行期間の扱い

帳票定義を削除するまで(tasks.md のタスク7)は、帳票定義ありの経路が残る。この間、`unsupportedElements: "error"` の帳票定義では、
ファサードの共通の経路(`ReportPdfConverter.ReadWorkbook`。`Convert*`・`ComputeLayout*`・`CheckFit*` のすべてが通る)で、
`WorkbookModel.UnsupportedElements` が1件以上なら例外を送出する。これにより中核仕様の要件13.5(収まりの確認は変換と同じエラー)を保つ。
移行期間には、同じ帳票定義でも次の2点が変わる。

- 例外の中身が、最初の1件から全件になる。
- 壊れた描画オブジェクト(`MalformedDrawingObject`)でも止まるようになる。

中核仕様は tasks.md のタスク1で先に改訂しておく。

## 3. フォントの代替(要件3)

- `FontResolver` は、代替した `FontKey` から「実際に使った書体のファミリ名」への対応を、キャッシュと同じ寿命で持つ。
  現在の `_substituted`(`HashSet<FontKey>`)を辞書にし、外字の補いの判定(`IsSubstituted`)は辞書のキーで行う。
- 次のメソッドを公開する。

  ```csharp
  // 指定のフォントを解決し(キャッシュ済みならそれを使う)、代替したならその書体のファミリ名を、そうでなければ null を返す。
  public string? GetSubstitution(FontStyle font);
  ```

  問い合わせの時点で解決する。Layout に `SkiaFontMetricsProvider` 以外の計測(テストの `ApproximateFontMetricsProvider`)を使う構成でも、
  未解決のフォントを漏らさないためである。厳格モードで無いフォントを問い合わせると `FontNotAvailableException` になる。
  これは保存の描画の段階で起きるエラーと同じである。
- ファサードは診断のときに、`PagedLayout` の全ページから、使ったフォントの集合(ファミリ名・太字・斜体)を作る。
  対象は `TextCommand.Font` と `ShapeTextLine.Font` で、`GroupCommand` の子は再帰的にたどる。ヘッダー/フッターも `TextCommand` として出力されるため含まれる。
  そのうえで、それぞれを `GetSubstitution` で問い合わせる。キャッシュ済みのフォントも、その変換で使っていれば必ず問い合わせることになるため、
  要件3.2(前の変換で解決済みのフォントも漏らさない・使っていないフォントは返さない)を満たす。
- 外字の補い(`ResolveGlyphFallbacks`)は別の経路で、上の辞書に入らない(要件3.4)。
- 厳格モードでは代替が起きないため、一覧は常に空になる(要件3.3)。

## 4. 変換設定と帳票定義の廃止(要件1.3, 1.5, 1.6)

### 4.1 削除するもの

| レイヤー | 削除するもの |
|---|---|
| ReportDefinition | `ReportDefinitionJsonReader`、`IReportDefinitionRepository`・`FileSystemReportDefinitionRepository`、`SubstitutionFieldDefinition`、`OverflowBehavior`、`UnsupportedElementPolicy`(旧)、`ReportDefinition.SubstitutionFields`・`PrintAreaOverride`・`ToleranceMm`・`UnsupportedElements`・`TryGetField`・`RequiredFields` |
| Substitution | `ICellSubstitutor.Apply`(置換キー)、必須キー・未知のキーの検証 |
| Layout | `ValidateRequiredFieldsAreInPrintRanges`、`FindSubstitutionKey`、`ReportModel.OverflowByCell` による上書き、`PrintAreaOverride` による印刷範囲の上書き、`FitCheckResult`、`FitIssue.SubstitutionKey` |
| Core(例外) | `ReportDefinitionNotFoundException`・`ReportDefinitionSchemaException`・`ReportStructureMismatchException`・`SubstitutionKeyNotFoundException`・`RequiredSubstitutionValueMissingException` |
| ファサード | `ReportPdfConverter` の公開API、`Excel2Pdf` の帳票コード付きコンストラクタ・`SetField`・`ClearField`・`ReportCode`・`CheckFit` |
| samples・docs | `samples/reports/*/definition.json`、`samples/edge-cases/*/definition.json`、`docs/帳票定義スキーマ.md` |

### 4.2 残すもの

- `ReportDefinition` は `ConversionSettings` に名前を変える。持つ値は、文書名・シート名・最大数字幅だけにする
  (現在の `CreateWithoutDefinition` で合成している値)。`EstimateMaxDigitWidthPx` はこの型に残す。
- `ReportModel` と `IReportModelBuilder` は残す(シートと変換設定を組にして Substitution・Layout に渡す役割は変わらない)。
- プロジェクト名 `Utsushi.ReportDefinition`・名前空間 `Utsushi.ReportDefinitions` は、本仕様では変えない。
  プロジェクトの改名はソリューション・テスト・ドキュメントへの影響が大きく、機能を変えずに単独で行える。
  そこで、別の変更として扱う(「未決事項」)。

### 4.3 セルの文字のはみ出し

キーごとの `overflow` を廃止すると、書き込んだセルの文字の扱いは、セルの書式だけで決まる。
これは現在の帳票定義なしの変換・セル番地直接指定と同じ規則である。

| セルの書式 | 扱い |
|---|---|
| 「縮小して全体を表示する」(`shrinkToFit`) | 縮小表示 |
| 「折り返して全体を表示する」(`wrapText`) | 折り返し(収まらなければ中核仕様の要件2.14のエラー) |
| どちらも無い | はみ出し表示(数値は切り取り) |

帳票定義の `overflow: "clip"` に相当するExcelの書式は無い。`clip` のセルは、はみ出し表示になる。
例として、請求書の `InvoiceNo`・`TotalAmount`、エッジケースの送付状の `Department` がある。
`overflow: "overflow"` と `overflow` の省略は、もともとセルの書式に従うため変わらない。ゴールデンへの影響は 6.1 で扱う。

### 4.4 「帳票コード」から「文書名」への改名(要件1.6)

`UtsushiException.ReportCode`・`PagedLayout.ReportCode`・`WorkbookReadOptions.ReportCode` は、帳票定義の廃止後は文書名を入れる欄になる。
そこで、名前を `DocumentName` に変える。`ToDiagnosticString()` の見出しも「文書」に変える。
ゴールデンのスナップショットの1行目(`report=...`)は `document=...` になる。

## 5. CLI(要件5)

### 5.1 コマンドとオプション

```
utsushi check   <入力.xlsx>              [オプション]
utsushi convert <入力.xlsx> <出力.pdf>   [オプション]
utsushi --help

  --text <セル>=<文字列>          SetText(複数指定可)
  --number <セル>=<数値>          SetValue(double。小数点は「.」、カルチャに依存しない)(複数指定可)
  --sheet <シート名>              Excel2PdfOptions.SheetName
  --document-name <名前>          Excel2PdfOptions.DocumentName
  --max-digit-width <px>          Excel2PdfOptions.MaxDigitWidthPx
  --ignore-unsupported            Excel2PdfOptions.UnsupportedElements = Ignore
  --allow-font-fallback [書体名]  FontResolverOptions.AllowFallback(書体名)
  --allow-missing-glyphs          PdfRenderOptions.MissingGlyphs = MissingGlyphPolicy.Render
  --outline-text                  PdfRenderOptions.OutlineText(文字のアウトライン化)
  -h, --help
```

- `--text`・`--number` を合わせて、同じセル(表記が違っても同じセル)を2回指定したら使い方エラー(要件5.4)。
  セル番地の正規化は `CellAddress.TryParse` を使う。
- 日時・時間の指定は設けない(必要ならシリアル値を `--number` で渡す。要件5.4の補足)。オプションを増やさず保守を軽くするためである。

### 5.2 出力と終了コード

| 終了コード | 意味 |
|---|---|
| 0 | `check`: 問題なし。`convert`: 保存した。`--help`: 使い方を標準出力に出した |
| 1 | `check`: 問題が1つ以上ある(`Excel2PdfDiagnostics.HasIssues`。`--ignore-unsupported`・`--allow-font-fallback` の指定によらない) |
| 2 | 使い方エラー(標準エラーに理由と使い方) |
| 3 | 変換・診断のエラー。`UtsushiException` は `ToDiagnosticString()` を、`IOException`・`UnauthorizedAccessException` は例外のメッセージを標準エラーに出す |

`check` の標準出力の形式は次のとおりである(1行1件。種類ごとにまとめ、件数を見出しに出す)。

```
未対応の要素: 2件
  [Chart] シート '注文書' D10: グラフは出力されません
  [UnsupportedShapePreset] シート '注文書' B30: 図形のプリセット 'heart' に対応していません
文字の収まらない箇所: 1件
  [NumberTooWide] C8(1ページ目): 数値の表示文字列がセルの幅に収まりません
フォントの代替: 1件
  ＭＳ Ｐゴシック → BIZ UDPゴシック
```

- 打ち切った種類には「(上限で打ち切り)」を見出しに付ける。
- 文字の収まらない箇所の行には、`FitIssue.Message`(セルの文字列を含まない)だけを出す。描画する文字列(`FitIssue.Text`)は出さない。

### 5.3 構成

- `Program.Main` は `Run(string[] args, TextWriter stdout, TextWriter stderr)` を呼ぶだけにし、テストから終了コードと出力を確かめられるようにする。
- 引数の解釈(`CommandLineOptions.Parse`)は、コマンドごとの必須の引数の検証を含めて純粋な関数にする(現在と同じく単体テストの対象)。

## 6. テスト

### 6.1 ゴールデンテストの移行(要件6)

1. `tools/Utsushi.SampleGenerator` で、サンプル帳票の書き込むセルに、廃止する `overflow` に相当する書式を付ける。
   - `shrink` → `shrinkToFit`
   - `wrap` → `wrapText`
   - `clip` → 何も付けない(4.3)
   - `overflow`・省略 → 変えない

   同じく `tools/edge-case-samples/generate.py` で、エッジケースのサンプルにも付ける(例: 注文書の `SupplierName`・売上一覧の `Department` に `shrinkToFit`)。
2. `ReportConversionGoldenTests` を、帳票コードと置換キーの辞書から、「サンプル名 → セル番地と値の辞書」に書き換える。
   - セル番地は各 `definition.json` の `cell` から写す。
   - 変換は `Excel2Pdf` で行う。エンジンは、決定的な `ApproximateFontMetricsProvider` と固定時刻を使う internal のコンストラクタ(1.1)で作る。
   - 帳票定義が指定していた値を `Excel2PdfOptions` で渡す。`MaxDigitWidthPx` には各定義の値(サンプル帳票は7。
     指定しないと、標準フォントの MS PGothic 11pt から8と見積もられ、全列の幅が変わる)を、`DocumentName` には帳票コードを渡す
     (指定しないと、ファイル名から `template` になる)。
   - エッジケースのサンプルは定義の値が8で、標準フォント(游ゴシック 11pt)からの見積もりと同じだが、同じく明示する。
3. フィクスチャのディレクトリ名(`invoice` など)は変えない。この段階(帳票定義の削除より前)のスナップショットの差分は、次の2種類だけになるはずである。
   - `FitIssue` の行の `key=<置換キー>` が `key=-` になること
   - `clip` のセルの描画がはみ出し表示になること

   これ以外の差分が出たら原因を調べる。差分は layout-fidelity-reviewer のレビューを経てから更新する。
   帳票定義の削除(4.4の改名・`SubstitutionKey` の削除)で、1行目が `document=` になり、`key=` の列が無くなる。
4. テンプレートを再生成すると、同じテンプレートを使う他のテスト(`WithoutDefinitionConversionTests`・`Excel2PdfTests` など)の結果も変わりうる。
   これらも同じ段階で確かめる。
5. この移行を、帳票定義の削除より前に行う(要件6.1)。移行したテストが緑であることを、削除の前後で確かめる。

### 6.2 エッジケース検証用サンプル

- `EdgeCaseSampleTests` の、帳票定義・置換キーを使うテストは、セル番地の指定に書き換える。
- 置換キーだけを対象とする観点(必須キー・未定義のキー)のテストは削除する。
- `unsupportedElements` の ignore/error のテストは、`Excel2PdfOptions.UnsupportedElements` と `Check()` のテストに置き換える。
  グラフを含む注文書(`edge-order-form`)で、次の3点を確かめる。
  - 既定ではエラーになり、例外の一覧にグラフ(アンカーのセル番地つき)が含まれる。
  - `Check()` は例外にせずに一覧を返す。
  - `Ignore` を指定すれば変換できる。
- 既存の `Excel2PdfTests` のうち、未対応要素を含むサンプル(`edge-order-form`)を既定の設定で開いているテストには、既定をエラーにする段階で `Ignore` を明示する。

### 6.3 追加するテスト

| 対象 | 観点 |
|---|---|
| Parsing | 2.2 の各経路で、要素が読み飛ばされ、かつ正しい種類・シート・アンカー・`Detail` で記録される。枚数の上限の超過は1件だけ、1つの要素の上限の超過は要素ごとに記録される。グラフは `Chart` の1件だけになる。収集の件数の上限(小さい上限を渡すオーバーロード)。読まないシートの要素を含まない。図形のテキストを持つ要素で、`Message`・`Detail` に文字列が入らない。長いcontent-typeが切り詰められる |
| Rendering | `GetSubstitution`: 厳格モードでは常に null(無いフォントは例外)。代替の後は書体名。外字の補いでは null のまま |
| ファサード | 既定でエラー・一覧の全件・`Ignore` で変換できる・`Check()` は例外にしない。フォントの代替が2回目の変換でも返る・使わないフォントは返らない。エンジンを渡す場合の `Fonts`・`Render` の指定はエラー。`Excel2Pdf` を破棄してもエンジンは使い続けられる。1つのエンジンを複数のスレッドの `Excel2Pdf` で同時に使っても結果が変わらない |
| CLI | `Run` の終了コード(0/1/2/3)と出力。`--help`。重複したセルの指定。`--number` の不正な値。未対応要素での中止の出力 |

## 7. ドキュメントと steering

- `.kiro/steering/`(structure.md の構成・帳票の追加手順、tech.md の記述)、`CLAUDE.md`、`README.md`、`docs/ライブラリの使い方.md`、
  `docs/テンプレート作成ガイド.md`、`docs/エッジケース検証レポート.md`、`samples/edge-cases/README.md`、`.claude/agents/` の帳票定義への言及を改める。
- 「帳票固有の分岐は帳票定義データ側に置く」という方針(`CLAUDE.md`「必ず守る制約」・`tech.md`「コーディング規約」・`structure.md`「レイヤー構成」)は、
  「帳票ごとの差異はテンプレート(Excelの書式・印刷設定)と呼び出し元が書き込む値で表し、共通レイヤーのコードに帳票固有の分岐を持ち込まない」に改める。
- 移行が完了したら、product.md・structure.md・tech.md・CLAUDE.md の「移行中」の注記を外す。
- `docs/実装設計失敗事例集.md` の過去の事例は、記録としてそのまま残す(廃止した機能に関する節には、廃止した旨の注記だけを加える)。

## 未決事項

- **プロジェクト `Utsushi.ReportDefinition` の改名**: 中身は変換設定とモデルの組み立てだけになるため、名前が実態に合わなくなる。
  機能を変えない別の変更として、本仕様の実装後に判断する。
- **書式の近似の検知**: requirements.md「未決事項」を参照。
