# 実装計画: 自社帳票のExcel→PDF変換

対象要件: `.kiro/specs/excel-report-pdf-conversion/requirements.md`
対象設計: `.kiro/specs/excel-report-pdf-conversion/design.md`

> **状況**: タスク1〜15完了(タスク1〜11: 2026-09-19、タスク12: 2026-09-20、タスク13〜14:
> 2026-09-21、タスク15: 2026-09-21)。タスク15(接続線・グループ・追加プリセット・多段階/放射状
> グラデーション)では、データモデル定義、Parsing/Layout/Renderingの3層すべてでの接続線・
> グループ対応、多段階/放射状グラデーション、星形・フローチャート記号・雲形/引き出し線付き
> 吹き出しのジオメトリ、SampleGeneratorへの機能追加とサンプル帳票への配置・ゴールデンテスト
> 更新、各層のユニットテスト追加、最終レビュー対応まで完了した。CLIで生成したPDFを
> `pdftoppm`でラスタライズして目視確認済み。`dotnet build` / `dotnet test`(427件、
> 2026-09-21時点)/ `dotnet format` はグリーン。code-reviewer/security-reviewer/
> layout-fidelity-reviewer/doc-reviewerによるレビューを実施し、指摘(shapeCount上限の
> グループ経由での超過、画像検証順序の後退、グラデーションストップ数無制限、負のa:ext値に
> よる未処理例外、requirements.mdの補足の掲載順・あいまいな相互参照)はいずれも修正済み。
> 実装時に決定した事項・判明した制約は `design.md` に反映済み。タスク文面どおりに
> 実現できなかった項目には各タスクに注記を付けた。

- [x] 1. ソリューション基盤のセットアップ
  - `.kiro/steering/structure.md` の構成に従い、`Utsushi.sln` と各レイヤーの `.NET 5` クラスライブラリプロジェクト(Parsing/ReportDefinition/Substitution/Layout/Rendering)、および対応するテストプロジェクトを作成する
  - Nullable参照型・ファイルスコープ名前空間を既定で有効化する共通 `Directory.Build.props` を用意する
  - _Requirements: 7.2, 7.3_
  - 注記: ソリューションファイルは `Utsushi.sln` ではなく新形式の `Utsushi.slnx` とした。
    また、例外階層とレイヤー共通の値型を置く `Utsushi.Core`、呼び出し元が参照するファサード
    `Utsushi` を追加している(`.kiro/steering/structure.md` 更新済み)。
    さらに、呼び出し元プロダクトの開発環境がVisual Studio 2019(C# 9.0まで)であることが判明したため、
    ファイルスコープ名前空間は不採用とし、`LangVersion` を明示的に `9.0` に固定した
    (`.kiro/steering/tech.md`「Visual Studio 2019 対応」参照)。

- [x] 2. Parsingレイヤー: WorkbookModel の実装
  - [x] 2.1 `DocumentFormat.OpenXml` を依存に追加し、`.xlsx` からセル値・スタイル(フォント/罫線/配置/数値書式/背景色)を読み取る `IWorkbookReader` を実装する
    - _Requirements: 1.1, 1.2, 7.1_
  - [x] 2.2 列幅・行高・結合セル範囲の読み取りを実装する
    - _Requirements: 1.2_
  - [x] 2.3 印刷範囲・手動改ページ・用紙サイズ/余白/拡大縮小・印刷タイトル・印刷順序の読み取りを実装する
    - _Requirements: 1.3_
  - [x] 2.4 サポート外要素(図形・外部参照等)の検出とエラー化/無視の切り替えを実装する
    - _Requirements: 1.5_
  - [x] 2.5 破損ファイル・非xlsx形式・パスワード保護ファイルに対する明確な例外を実装する
    - _Requirements: 6.1, 6.2_
  - [x] 2.6 入力ファイルが存在しない/アクセス権限がなく開けない場合も `InvalidExcelFileException` に統一する
    - _Requirements: 6.5_
    - 注記: 当初 `ReportPdfConverter.ConvertToFile` は `File.OpenRead` の失敗(存在しない等)を
      未加工の `IOException` 系のまま呼び出し元へ伝播していた。ストリーム版 `Convert` との挙動統一のため、
      2026-09-21のリファクタリングで `UtsushiException` 階層(Stage=Parsing)へ変換するよう修正し、
      要件6.5として追記した。

- [x] 3. ReportDefinitionレイヤー: 帳票定義のロードと突合
  - [x] 3.1 帳票定義JSONのスキーマを定義し、ロード時のスキーマ検証を実装する
    - _Requirements: 6.3, 8.1_
  - [x] 3.2 `WorkbookModel` と帳票定義を突合して `ReportModel` を構築する `IReportModelBuilder` を実装し、シート名/セル番地の不一致を検出する
    - _Requirements: 1.4_
  - [x] 3.3 最初の対象帳票(例: 請求書)の帳票定義ファイルと、マスキング済みサンプルExcelを `samples/reports/` に追加する
    - _Requirements: 8.3_

- [x] 4. Substitutionレイヤー: セル値の置換
  - [x] 4.1 置換キー→値の辞書を受け取り、対象セルの値のみを更新する `ICellSubstitutor` を実装する(スタイルは変更しない)
    - _Requirements: 2.1, 2.2_
  - [x] 4.2 未知の置換キー・必須キー欠落時の例外を実装する
    - _Requirements: 2.3, 2.4_
  - [x] 4.3 はみ出し時の挙動フラグ(shrink/clip/wrap)を `ReportModel` に保持する処理を実装する(実際の描画反映はLayoutレイヤーで行う)
    - _Requirements: 2.5_

- [x] 5. Layoutレイヤー: ページ分割と座標計算
  - [x] 5.1 列幅(文字単位)・行高からポイント単位への換算ロジックをプロトタイプし、対象帳票での誤差を検証した上で実装する
    - _Requirements: 4.5_
  - [x] 5.2 印刷範囲によるクリッピングを実装する
    - _Requirements: 3.1_
  - [x] 5.3 手動改ページの適用を実装する
    - _Requirements: 3.2_
  - [x] 5.4 自動改ページ計算(用紙サイズ・余白・拡大縮小率から導く印字可能領域に基づく行/列分割)を実装する
    - _Requirements: 3.3_
  - [x] 5.5 印刷タイトル行/列の複製を実装する
    - _Requirements: 3.4_
  - [x] 5.6 印刷順序(行優先/列優先)に従ったページ順序決定を実装する
    - _Requirements: 3.5_
  - [x] 5.7 結合セルの矩形統合、罫線・背景色の描画命令生成を実装する
    - _Requirements: 4.2, 4.3_
  - [x] 5.8 テキストの水平/垂直配置・インデント・縮小表示(shrink指定時)の座標計算を実装する
    - _Requirements: 4.1, 4.4, 2.5_

- [x] 6. Renderingレイヤー: PDF描画
  - [x] 6.1 `SkiaSharp` を依存に追加し、`PagedLayout` を `SKDocument` PDFバックエンドで描画する `IPdfRenderer` を実装する(背景→罫線→テキストの順で描画)
    - _Requirements: 5.1, 5.2, 5.3, 7.1_
  - [x] 6.2 対象帳票が使用するフォントの埋め込み(サブセット化)を実装し、フォント未検出時をビルド/デプロイ時エラーとして検出する仕組みを用意する
    - _Requirements: 4.1_
    - 注記: **サブセット化は実現できていない**。SkiaSharp の PDF バックエンド(NuGet配布の
      ネイティブビルド)がサブセット化に対応しておらず、使用フォントを丸ごと埋め込む
      (2.88系・3.x系の双方で実測確認)。日本語帳票1ページで約4.3MBになる。
      **PDF内の文字列検索を維持するため、フォント埋め込みを既定とし、このサイズは許容する**
      という判断になった。サイズを優先する場合のために、文字をアウトライン化して出力する
      `PdfTextRendering.Outline`(同帳票で91KB、見た目は同一、ただしPDF内検索は不可)も選べる。
      サブセット化の根本対応は `design.md`「未決事項」に残している。
    - フォント未検出の検出は `FontResolver`(既定で厳格モード、フォールバックしない)と
      事前検証用の `FontResolver.EnsureAvailable` で実装した。
  - [x] 6.3 一時ストリームへの描画→正常終了時のみ最終出力先に確定させる仕組みを実装し、失敗時に不完全なPDFを残さないようにする
    - _Requirements: 5.4_

- [x] 7. エラーハンドリング/ログの横断対応
  - [x] 7.1 `UtsushiException` 基底クラスと各派生例外(帳票コード/シート名/セル番地/処理段階を保持)を実装する
    - _Requirements: 6.4_
  - [x] 7.2 各レイヤーの例外送出箇所を7.1の例外型に統一する
    - _Requirements: 1.4, 1.5, 2.3, 2.4, 6.1, 6.2, 6.3_

- [x] 8. テスト整備
  - [x] 8.1 Layoutレイヤー(改ページ計算・単位換算)、Substitutionレイヤー(必須/未知キー判定)のユニットテストを作成する
    - _Requirements: 3.3, 2.3, 2.4_
  - [x] 8.2 最初の対象帳票についてのエンドツーエンド(Parsing→Rendering)ゴールデンテストを作成する
    - _Requirements: 8.3_
  - [x] 8.3 帳票定義JSONのスキーマ検証テストを作成する
    - _Requirements: 6.3_

- [x] 9. エントリーポイントの実装
  - [x] 9.1 5レイヤーを組み立てて「帳票コード + 置換値辞書 + 入力Excel」から「PDFストリーム」を得るユースケースクラス/CLIを実装する
    - _Requirements: 7.2, 7.3_

- [x] 10. 2つ目以降の帳票への対応(拡張性の検証)
  - [x] 10.1 共通レイヤーのコード変更なしに、2つ目の帳票定義とサンプル・ゴールデンテストを追加できることを確認する
    - _Requirements: 8.1, 8.2, 8.3_

- [x] 11. PDF出力品質とページ構成の拡張
  - [x] 11.1 太字/斜体を Type 3 フォントにせず、フォント埋め込み(CID TrueType)のまま再現する
    - _Requirements: 4.1, 5.3_
    - 実フォントが該当字形を持つ場合は実字形を使い、持たない場合は通常字形に描画時の装飾を合成する。
      判定は通常字形とのフォントデータ一致で行う。
  - [x] 11.2 複数の印刷範囲をそれぞれ独立したページ群として出力する
    - _Requirements: 3.6, 3.1_
  - [x] 11.3 ページヘッダー/フッターの読み取り・書式コード展開・配置を実装する
    - _Requirements: 3.7, 3.8, 3.9_
  - [x] 11.4 上記を検証するサンプル帳票(領収書: 複数印刷範囲、納品書: ヘッダー/フッター)と
        ゴールデンテストを追加する
    - _Requirements: 8.3_

- [x] 12. セル番地直接指定によるオーバーライド機能
  - [x] 12.1 `ICellSubstitutor` に `ApplyCellOverrides(ReportModel, IReadOnlyDictionary<string, string>)` を追加し、`CellSubstitutor` で実装する(帳票定義未登録セルへのA1形式直接指定、書式は変更しない)
    - _Requirements: 2.7_
  - [x] 12.2 セル番地がA1形式として解釈できない場合の例外 `InvalidCellOverrideAddressException` を実装する
    - _Requirements: 2.8_
  - [x] 12.3 `Utsushi.ReportPdfConverter` の `Convert` / `ConvertToFile` / `ComputeLayout` に任意パラメータ `cellOverrides` を追加し、`Apply` の後段で `ApplyCellOverrides` を適用する
    - _Requirements: 2.7_
  - [x] 12.4 CLIに `--override <セル番地>=<値>` オプションを追加する
    - _Requirements: 2.7_
  - [x] 12.5 ユニットテスト(未登録セルの上書き、既存書式の維持、不正なセル番地でのエラー)とゴールデンテスト(既存の置換キーとの併用)を追加する
    - _Requirements: 2.7, 2.8_
  - [x] 12.6 結合セル範囲の非アンカー位置への直接指定を検出する例外 `NonAnchorMergedCellOverrideException` を実装し、名前付きキー方式が残した `OverflowByCell` エントリを上書き対象セルから取り除く
    - _Requirements: 2.9_
    - 注記(コードレビューで発見): 結合セルの非アンカー位置を直接指定すると、Layoutレイヤーはアンカーの値しか描画しないため値が静かに失われる。また同一セルが帳票定義の置換キーにも登録され`overflow`が明示されている場合、`ApplyCellOverrides`だけでは`OverflowByCell`のエントリが残ってしまい「Excel側の書式に従う」という設計と矛盾する。いずれも12.1のレビューで発見し、本タスクで合わせて修正した。

- [x] 13. シート内画像(ロゴ等)の再現
  - [x] 13.1 Parsingレイヤー: 画像のデータモデル(`ImageModel` / `ImageExtent`)を定義する
    - `ImageExtent` の派生として、固定サイズ(oneCellAnchor相当)の `FixedImageExtent`、
      対角セル指定(twoCellAnchor相当)の `CellSpanImageExtent` を実装し、`SheetModel.Images` として保持する
    - _Requirements: 9.1, 9.2_
  - [x] 13.2 Parsingレイヤー: `xdr:pic`(oneCellAnchor/twoCellAnchor)から画像を読み取る
    - `WorksheetPart.DrawingsPart.WorksheetDrawing` のアンカーから `a:blip` の `r:embed` を辿って
      `ImagePart` を解決し、バイナリと `ContentType`、0始まりのOOXMLマーカーを1始まりの
      `CellAddress` + ポイント単位オフセットへ変換する(`ReadImages` / `TryReadMarker`)
    - _Requirements: 9.1, 9.2_
  - [x] 13.3 Parsingレイヤー: `DetectUnsupportedElements` を改修し、画像のみのDrawingPartを誤検出しないようにする
    - 従来は `DrawingsPart` の存在だけで無条件に「サポート外要素(Drawing)」としていたが、
      アンカーを列挙して `xdr:pic` 以外(`xdr:sp`/`xdr:grpSp`/`xdr:cxnSp` 等)が
      1つでも含まれる場合のみ例外化するよう変更する(design.md「Parsing レイヤー」の必須修正点。
      これをしないと画像を1つでも含むシートが `unsupportedElements: "error"` で常に失敗し続ける)
    - _Requirements: 9.5_
  - [x] 13.4 Parsingレイヤー: 対応形式外の画像を `UnsupportedImageFormat` として扱う
    - 画像の `ContentType` が `image/png`/`image/jpeg`/`image/gif`/`image/bmp` 以外
      (EMF/WMF等)の場合は `UnsupportedWorkbookElementException`
      (`ElementKind = "UnsupportedImageFormat"`)を、既存の `unsupportedElements` ポリシー
      (ignore/error)に従って送出/無視する
    - _Requirements: 9.4_
  - [x] 13.5 Layoutレイヤー: 固定サイズ画像(oneCellAnchor相当)のページ座標変換を実装する
    - `SheetModel.Images` の各画像について、アンカーセルがページに含まれる場合のみ、
      結合セルと同じ座標変換の仕組み(`SheetGrid`)でページ左上原点のポイント座標に変換し
      `ImageCommand` を生成する(`PageCommandBuilder.EmitImages`)
    - _Requirements: 9.1, 9.2_
  - [x] 13.6 Layoutレイヤー: 2セルアンカー(twoCellAnchor相当)の幅・高さ計算を実装する
    - アンカーセルから対角セルまでの列幅/行高を `SheetGrid` で合算して幅・高さを求める
      (`SpanWidthPt` / `SpanHeightPt`)。対角セルが現在のページの行/列範囲外にあっても、
      印刷範囲全体の格子を保持する `SheetGrid` を使うため正しく計算できる
    - _Requirements: 9.2_
  - [x] 13.7 Layoutレイヤー: 改ページをまたぐ画像をアンカー側のページにのみ配置する
    - 結合セルのような「見えている部分だけ切り出す」対応はせず、アンカー左上セルが
      属するページにのみ画像全体を配置し、他のページには描画しない(design.md「未決事項」の割り切り)
    - _Requirements: 9.1_
  - [x] 13.8 Renderingレイヤー: `ImageCommand` を描画する
    - `SKBitmap.Decode` でデコードし `SKCanvas.DrawBitmap` で描画する。描画順を
      背景→罫線→テキスト→画像 とし、Excelと同様に画像が他のセル内容より最前面に来るようにする
    - _Requirements: 9.1, 9.3_
  - [x] 13.9 `Utsushi.SampleGenerator` に画像埋め込み機能を追加する
    - `SpreadsheetBuilder.SetImage`(oneCellAnchorでの画像配置)を実装する。画像処理ライブラリは
      追加せず、`PlaceholderPng`(PNGを直接組み立てる自前の最小エンコーダ)で単色矩形のロゴを生成する
    - _Requirements: 8.3_
  - [x] 13.10 サンプル帳票(invoice)にロゴ画像を配置し、ゴールデンテストを更新する
    - invoiceサンプルのF1セル(表題行の右上)にロゴ画像を配置して `template.xlsx` を再生成し、
      `UTSUSHI_UPDATE_GOLDEN=1 dotnet test` でゴールデンファイルを更新する
      (差分は実際にCLIでPDFを生成し目視確認したうえでコミットした)
    - _Requirements: 9.1, 9.2, 9.3, 8.3_
  - [x] 13.11 Parsing層のユニットテストを追加する
    - oneCell/twoCellアンカーのパース、`ContentType` 許可リスト判定によるignore/error、
      画像のみのシートで `unsupportedElements: error` でも例外にならないこと、
      画像以外の図形は引き続き例外になることを検証する
    - _Requirements: 9.1, 9.2, 9.4_
  - [x] 13.12 Layout層のユニットテストを追加する
    - 固定サイズ画像・2セルアンカー画像の座標変換、改ページをまたぐ画像がアンカー側の
      ページにのみ配置されることを検証する
    - _Requirements: 9.1, 9.2_
  - [x] 13.13 Layoutレイヤー: 2セルアンカーの幅・高さ計算が印刷範囲外の列/行を0扱いする不具合を修正する
    - `SpanWidthPt`/`SpanHeightPt`が印刷範囲にクリップされた`SheetGrid`を使っていたため、
      対角セルが印刷範囲のすぐ外にあるだけで画像が実際より小さく計算されていた
      (layout-fidelity-reviewer指摘)。`SheetModel`から直接列幅/行高を取得するよう修正し、
      回帰テストを追加する
    - _Requirements: 9.2_
  - [x] 13.14 Renderingレイヤー: 画像デコード失敗時の例外に帳票コード・シート名を含める
    - `DrawImage`が送出する`PdfRenderingException`にreportCode/sheetNameが渡っておらず、
      同ファイル内の他の送出箇所と一貫性がなかった(code-reviewer必須指摘、要件6.4)。
      `DrawPage`からreportCode/sheetNameを伝播させ、ユニットテストを追加する
    - _Requirements: 6.4_
  - [x] 13.15 セキュリティ対策: ピクセル爆弾・計算量DoSへの上限を追加する(要件9.6)
    - Renderingレイヤー: `SKBitmap.Decode`で実際に展開する前に`SKBitmap.DecodeBounds`で
      宣言上のピクセル寸法を確認し、上限(既定4096px)を超える場合は拒否する
    - Layoutレイヤー: 2セルアンカーの幅・高さ計算(列/行の合算)に上限(既定4096列/行)を設け、
      対角セルにセル番地の上限近くを指定された場合の計算量を抑える。画像の表示サイズにも
      上限(既定5000pt)を設ける
    - Parsingレイヤー: 1シートあたりの画像アンカー数の上限(既定50枚、超過は
      `ElementKind = "TooManyImages"`)、画像1枚あたりの読み取りバイト数の上限
      (既定10MB、超過は`ElementKind = "ImageTooLarge"`)を設ける
    - security-reviewer指摘。いずれもセキュリティレビューで発見された、悪意あるExcelファイルに
      よるDoS(数百バイトのファイルで大きなメモリ・CPU消費を引き起こせる)への対策
    - _Requirements: 9.6_
  - [x] 13.16 セキュリティ対策: 画像バイナリの先頭シグネチャ(マジックバイト)を検証する
    - `ContentType`はOPCパッケージ側の申告値に過ぎず実バイト列と一致する保証がないため、
      ネイティブコードのデコーダ(SkiaSharp)に渡す前にPNG/JPEG/GIF/BMPの先頭バイトを
      比較する(security-reviewer指摘)。不一致の場合は`ElementKind = "UnsupportedImageFormat"`
      として扱う
    - _Requirements: 9.4_
  - [x] 13.17 13.13〜13.16の追加分のユニットテストを追加する
    - 印刷範囲外の対角セルを持つ2セルアンカーの回帰テスト、画像デコード失敗時の
      帳票コード/シート名検証、ピクセル爆弾対策(宣言サイズ超過の拒否)、
      画像枚数上限、ContentType偽装の検出を検証する
    - _Requirements: 9.2, 9.4, 9.6, 6.4_

- [x] 14. シート内図形(シェイプ)の再現
  - [x] 14.1 Parsingレイヤー: 画像・図形共通のデータモデルへリファクタリングする
    - `ImageExtent`/`FixedImageExtent`/`CellSpanImageExtent` を `AnchorExtent`/
      `FixedAnchorExtent`/`CellSpanAnchorExtent` に改名し、`DrawingObjectModel`
      (`AnchorCell`/`AnchorOffset`/`Extent` を持つ抽象基底)を新設して `ImageModel` に
      継承させる。`SheetModel.Images` を `SheetModel.DrawingObjects` に置き換え、
      呼び出し側(`PageCommandBuilder`・SampleGenerator・既存テスト)を追随させる
      (design.md「Parsing レイヤー」の図形節参照)
    - _Requirements: 9.1, 9.2_
  - [x] 14.2 Parsingレイヤー: 図形のデータモデル(`ShapeModel` / `ShapePresetType` /
        `ShapeFill` / `ShapeOutline` / `ShapeTextBody`)を定義する
    - _Requirements: 10.1, 10.4, 10.5, 10.6_
  - [x] 14.3 Parsingレイヤー: `xdr:sp` から対応済みプリセットジオメトリの図形を読み取る
    - `xdr:spPr/a:prstGeom/@prst` を静的なマッピング表と突き合わせ、一致しないものは
      読み取らない(14.8で非対応要素として検出させる)。アンカー(セル位置・オフセット・
      固定/セル追従の範囲)は画像と共通の仕組み(14.1)を使う。旧`ReadImages`を
      `ReadDrawingObjects`に置き換え、1回のアンカー列挙の中で`ReadImage`/`ReadShape`を
      呼び分けて `drawing.xml` の出現順を保った `DrawingObjects` を構築する
    - _Requirements: 10.1, 10.2, 10.3_
  - [x] 14.4 Parsingレイヤー: 図形の調整ガイド値(`a:avLst`)・回転(`a:xfrm/@rot`)を読み取る
    - ガイド名(`adj`/`adj1`/`adj2`等)を `ShapePresetType` ごとに定めた順序で
      `IReadOnlyList<double>` に整形する。ファイルに該当ガイドが無い位置は`double.NaN`とし、
      Renderingレイヤーがその位置のECMA-376既定値を補う(既定値表はRenderingレイヤー側に置く)
    - _Requirements: 10.1, 10.5_
  - [x] 14.5 Parsingレイヤー: 図形の塗りつぶし(`a:solidFill`/`a:gradFill`/`a:noFill`)と
        枠線(`a:ln`)を読み取る
    - グラデーションは先頭・末尾の `a:gs` の色のみ採用し、`a:lin/@ang` があれば角度として読み取る。
      色は`a:srgbClr`(RGB直接指定)のみ対応し、テーマ/システム色は解決しない(未対応分は
      塗り/枠線を無しとして扱う)
    - _Requirements: 10.1, 10.6_
  - [x] 14.6 Parsingレイヤー: 図形内テキスト(`xdr:txBody`)を段落・ラン単位で読み取る
    - `a:bodyPr/@anchor` を垂直配置、各 `a:p/a:pPr/@algn` を段落ごとの水平配置として読み取り、
      `a:r/a:rPr`(サイズ・太字・斜体・色・書体)と `a:t` から `FontStyle` と同じ型で
      `ShapeTextRun` を構築する(折り返しはLayoutレイヤーの責務なので行わない)
    - _Requirements: 10.4_
  - [x] 14.7 Parsingレイヤー: `DetectUnsupportedElements` を拡張し、接続線・グループを
        引き続きサポート外要素として検出する
    - 「`xdr:pic` 以外はすべて `Drawing` として例外化」だった判定を「`xdr:pic` および
      `xdr:sp`(シェイプ)以外(接続線 `xdr:cxnSp`、グループ `xdr:grpSp`)が1つでもあれば
      `Drawing` として例外化」に拡張する(`HasNonPictureDrawingObject`を
      `HasUnsupportedDrawingObject`に改名)。ここでの`xdr:sp`判定は構造的なもの(要素の種類が
      シェイプかどうか)であり、プリセットが対応済みかどうかは問わない(画像の
      `ContentType`許可リスト判定が`DetectUnsupportedElements`ではなく`ReadImage`側の
      個別検証であるのと同じ位置付け)。非対応プリセットの`xdr:sp`はこの時点では
      素通りし、後段の図形読み取り(14.8)が個別に`UnsupportedShapePreset`として検出する
    - _Requirements: 10.7_
  - [x] 14.8 Parsingレイヤー: 対応済み一覧に無いプリセットジオメトリを `UnsupportedShapePreset`
        として扱う
    - `UnsupportedWorkbookElementException`(`ElementKind = "UnsupportedShapePreset"`)を、
      既存の `unsupportedElements` ポリシー(ignore/error)に従って送出/無視する
    - _Requirements: 10.7_
  - [x] 14.9 セキュリティ対策: 図形個数・テキスト文字数の上限を設ける(要件10.8)
    - 1シートあたりの図形アンカー数の上限(既定50個、超過は `ElementKind = "TooManyShapes"`)、
      図形1つあたりの全テキスト文字数の上限(既定2000文字、超過は
      `ElementKind = "ShapeTextTooLong"`)を、画像の上限(13.15)と同じ考え方で設ける
    - _Requirements: 10.8_
  - [x] 14.10 Layoutレイヤー: 図形のページ座標変換を実装する(`PageCommandBuilder.EmitDrawingObjects`)
    - `SheetModel.DrawingObjects`(画像・図形が混在)を出現順に処理し、共通の
      `TryComputeDrawingObjectRect`(旧`EmitImages`を一般化)で矩形を求める。画像・図形の
      `DrawCommand` を出現順のまま1つのリストに追加し、z-orderを保つ(design.md
      「Layout レイヤー」の図形節、要件10.3)
    - _Requirements: 10.1, 10.2, 10.3_
  - [x] 14.11 Layoutレイヤー: 図形内テキストの折り返し・配置を実装する
    - `ShapeModel.Text` を `IFontMetricsProvider` で図形の矩形幅を基準に折り返し、
      段落の水平配置・`VAlign` に基づく垂直位置から各行のローカル座標(回転前)を算出し
      `ShapeCommand.TextLines` を構築する(`BuildShapeTextLines`/`WrapShapeText`)。
      段落内の複数ランは先頭ランのフォントで折り返しを代表させる近似とした
    - _Requirements: 10.4_
  - [x] 14.12 Layoutレイヤー: 改ページをまたぐ図形をアンカー側のページにのみ配置する
    - 画像(13.7)と同じ割り切りを図形にも適用する
    - _Requirements: 10.1_
  - [x] 14.13 Renderingレイヤー: 対応済みプリセットのパス生成(`ShapeGeometryBuilder`)を実装する
    - `rect`/`roundRect`/`ellipse`/`triangle` と、`rightArrow`/`leftArrow`/`upArrow`/
      `downArrow`/`leftRightArrow`/`upDownArrow` のパス生成を、既定調整値をフォールバックとして
      実装した。矢印は「右向き・幅=進行方向」のローカル座標で組み立て、明示的なアフィン変換
      (`TransformArrowLocalPath`)で4方向に配置する。双方向矢印(leftRightArrow/upDownArrow)は
      単方向矢印と同じ既定矢尻長さ比(0.5)を使うと両端の矢尻で幅を使い切り菱形に潰れるため、
      専用の既定値(0.25)を設けた(単一の`.pdf`をラスタライズして目視確認済み)
    - _Requirements: 10.1_
  - [x] 14.14 Renderingレイヤー: 吹き出し(`wedgeRectCallout`/`wedgeRoundRectCallout`/
        `wedgeEllipseCallout`)のパス生成を実装する
    - 本体形状(矩形/角丸矩形/楕円)に、引き出し先端(調整値2つを本体の幅・高さに対する
      比率とみなす)から最も近い辺へ向けた引き出し三角形を追加する2段階のパス構築とした
    - _Requirements: 10.1_
  - [x] 14.15 Renderingレイヤー: 図形の塗りつぶし・枠線・回転を描画する
    - `SolidShapeFill`/`LinearGradientShapeFill`/`noFill`、`ShapeOutline` の描画と、
      `RotationDegrees` に応じた `canvas.Save`/`RotateDegrees`/`Restore` を実装した
    - _Requirements: 10.1, 10.5, 10.6_
  - [x] 14.16 Renderingレイヤー: 図形内テキストを描画する
    - `ShapeTextLine` を一時的な `TextCommand`(ClipRectなし)に変換し、セル内テキストと同じ
      `DrawText` を再利用することでフォント解決・太字/斜体合成ロジックを重複させずに実装した。
      図形本体と同じ回転変換の内側(`Save`/`Restore`の間)で描画するため回転が反映される
    - _Requirements: 10.4, 10.5_
  - [x] 14.17 `Utsushi.SampleGenerator` に図形埋め込み機能を追加する
    - `SpreadsheetBuilder.SetShape`(プリセット・塗り/枠線・回転・テキストを指定できる)を実装した。
      画像と図形が同じシートに混在する場合の重なり順(要件10.3)を保つため、両者を1つの
      `DrawingsPart`にまとめる`AppendDrawingObjects`に統合した(旧`AppendImage`を置き換え)
    - _Requirements: 8.3_
  - [x] 14.18 サンプル帳票に図形(注記の吹き出し等)を配置し、ゴールデンテストを更新する
    - invoiceサンプルの合計行と明細表の間に、基本図形+回転+テキスト(`RoundRect`、
      赤枠・回転-12度・「見本」)、矢印(`RightArrow`)、吹き出し+テキスト
      (`WedgeRoundRectCallout`、「ご確認ください」)を配置し、`UTSUSHI_UPDATE_GOLDEN=1 dotnet test`
      でゴールデンファイルを更新した(差分は実際にCLIでPDFを生成しラスタライズして
      目視確認したうえでコミットした)。グラデーション塗りはサンプルには含めず、
      Rendering層のユニットテスト(14.20)で個別に検証する
    - _Requirements: 10.1〜10.5, 8.3_
  - [x] 14.19 Parsing層のユニットテストを追加する
    - `ShapeWorkbookFixtures`/`ShapeReadingTests`を追加し、対応済み/非対応プリセットの判定、
      oneCell/twoCellアンカー、調整ガイド値(既定値へのフォールバック含む)・回転・
      単色/グラデーション塗り・noFill・枠線・テキスト(段落/配置/フォント)の読み取り、
      図形個数/テキスト文字数の上限のignore/error、接続線・グループが引き続き`Drawing`に
      分類されること、画像と図形が混在する場合の`DrawingObjects`の出現順維持を検証した
      (17件追加)
    - _Requirements: 10.1〜10.8_
  - [x] 14.20 Layout層・Rendering層のユニットテストを追加する
    - Layout層: `ReportLayoutEngineTests`に、固定/2セルアンカーの座標変換、改ページをまたぐ
      図形の配置、テキストの折り返し・垂直配置、画像と図形混在時の出現順維持を追加した(6件)。
      Rendering層: `ShapeGeometryBuilderTests`で全13プリセットの非空パス生成・矩形への
      収まり方(吹き出しは意図的に矩形外へ広がる)・矢印の先端位置・調整値のフォールバックを
      検証し(36件)、`SkiaPdfRendererTests`に全プリセットの描画・回転後の後続描画命令への
      影響がないこと・グラデーション/noFill/テキスト描画を追加した(18件)
    - _Requirements: 10.1〜10.6_
  - [x] 14.21 レビュー対応
    - `layout-fidelity-reviewer`指摘: 図形内テキストのフォントサイズ・矩形内側余白に
      印刷拡大縮小率(`_scale`)が適用されておらず、縮小率を持つ帳票でテキストが矩形から
      はみ出す不具合を修正した(`PageCommandBuilder.BuildShapeTextLines`/`WrapShapeText`)。
      縮小率によらず折り返し行数が一定になる回帰テストを追加し、修正前は実際に
      検出できる(2行→10行に増える)ことを確認した
    - `security-reviewer`指摘: (1) 図形内テキストの文字数上限チェックが全ラン読み取り・
      フォント解析後にしか効いていなかったため、累積文字数を数えながら上限超過時点で
      即座に打ち切るよう修正(`OpenXmlWorkbookReader.ReadShapeText`)。(2) 吹き出しの
      引き出し先端の調整値(ファイル由来、理論上Int32の全域に相当する値を取りうる)が
      クランプされておらず極端な座標になりうる問題を、絶対値5.0への丸めで修正
      (`ShapeGeometryBuilder.WedgeCalloutPath`)。(3) OOXMLパーツ全体のサイズ上限が無い点・
      画像/図形の上限がシート単位でワークブック単位の合算上限が無い点をdesign.mdの
      未決事項に記載(今回の主経路では実害無しを確認済み)
    - `code-reviewer`指摘: design.mdが吹き出しの`adj2`既定値を`0.25`と記載していたが、
      実装(`ShapeGeometryBuilder.DefaultCalloutTipYAdj`)は目視確認の結果`0.75`を採用して
      おり、design.mdを更新せず放置していた無断逸脱を修正(design.mdを実装値に合わせて
      更新し、選定経緯を明記)。`ReadShapeText`の重複した`<summary>`タグを1つに統合。
      レイヤー越境(`Utsushi.Parsing.Model`の型がLayout/Renderingから直接参照される点)は
      既存の`FontStyle`と同じ確立済みパターンであり本PR起因の新規逸脱ではないため対応不要
      と判断
    - _Requirements: 10.7, 10.8, 6.4_

- [x] 15. 図形対応の拡張(接続線・グループ・追加プリセット・多段階/放射状グラデーション)
  - [x] 15.1 Parsingレイヤー: 接続線のデータモデル(`ConnectorModel` / `ConnectorPresetType`)を定義する
    - _Requirements: 10.9_
  - [x] 15.2 Parsingレイヤー: グループのデータモデル(`GroupShapeModel` / `GroupChildModel`
        階層: `GroupChildShape`/`GroupChildImage`/`GroupChildConnector`/`GroupChildGroup`)を定義する
    - _Requirements: 10.10_
  - [x] 15.3 Parsingレイヤー: グラデーションのデータモデルを拡張する(`GradientStop`,
        `LinearGradientShapeFill`を複数ストップ対応に変更, `RadialGradientShapeFill`を追加)
    - 既存の`LinearGradientShapeFill(StartColor, EndColor, AngleDegrees)`を
      `LinearGradientShapeFill(IReadOnlyList<GradientStop> Stops, AngleDegrees)`に変更したため、
      既存の呼び出し箇所(Rendering層の`CreateShapeFillPaint`等)・テストを追随修正した
    - _Requirements: 10.6_
  - [x] 15.4 Parsingレイヤー: 追加プリセット(星形4種・フローチャート記号7種・
        追加吹き出し4種)を`ShapePresetType`と`SupportedShapePresets`マッピング表に追加する
    - 星形4種の調整ガイド名はECMA-376既定で`adj1`ではなく単一の`adj`であることを
      SDKの`Dr.ShapeTypeValues`一覧で裏取りして採用した。フローチャート記号・
      cloudCallout・callout1-3は固定比率/固定形状として描画するため調整ガイドを
      読み取らない(`ShapeAdjustmentGuideNames`は空配列。design.md参照)
    - _Requirements: 10.1_
  - [x] 15.5 Parsingレイヤー: 多段階・放射状グラデーションを読み取る
    - `a:gsLst`内の全`a:gs`(位置・色)を`GradientStop`のリストとして読み取り、
      `a:lin`/`a:path`(`@path="circle"`のみ厳密対応、`"rect"`/`"shape"`は放射状に
      フォールバック)を判別して`LinearGradientShapeFill`/`RadialGradientShapeFill`を
      構築する
    - _Requirements: 10.6_
  - [x] 15.6 Parsingレイヤー: 接続線(`xdr:cxnSp`)を読み取る(`ReadConnector`)
    - 対応済みプリセット(`straightConnector1`/`bentConnector2`/`bentConnector3`/
      `curvedConnector2`/`curvedConnector3`)の判定、反転(`flipH`/`flipV`)・回転・枠線の
      読み取りを行う。非対応プリセットは`UnsupportedShapePreset`として扱う
    - _Requirements: 10.7, 10.9_
  - [x] 15.7 Parsingレイヤー: グループ(`xdr:grpSp`)を読み取る(`ReadGroupShape`)
    - `grpSpPr/a:xfrm`から`ChildOffset`/`ChildExtent`/回転を読み取り、直接の子要素
      (`xdr:sp`/`xdr:pic`/`xdr:cxnSp`/入れ子の`xdr:grpSp`)を出現順に`GroupChildModel`へ
      変換する再帰処理を実装した(`ReadGroupChildren`/`ReadGroupChildShape`/
      `ReadGroupChildImage`/`ReadGroupChildConnector`/`ReadGroupChildGroup`)
    - _Requirements: 10.10_
  - [x] 15.8 Parsingレイヤー: `DetectUnsupportedElements`/`HasUnsupportedDrawingObject`を
        拡張し、`xdr:cxnSp`/`xdr:grpSp`を構造的に許容する
    - 既存の「`xdr:pic`/`xdr:sp`以外は`Drawing`」の判定に`xdr:cxnSp`/`xdr:grpSp`を追加した
    - _Requirements: 10.9, 10.10_
  - [x] 15.9 Parsingレイヤー: グループ内に非対応要素が1つでもあればグループ全体を
        サポート外要素として扱う
    - `ReadGroupChildren`が子孫(再帰的に)を検証し、対応済みプリセット一覧に含まれない
      図形・接続線・`xdr:graphicFrame`等が1つでもあれば、Errorモードは即座に例外を送出し、
      Ignoreモードは`null`を返してグループ全体を破棄する(トップレベル/ネストいずれも
      呼び出し元が`null`を伝播させて全体を`unsupportedElements`ポリシーに従わせる)
    - _Requirements: 10.10_
  - [x] 15.10 セキュリティ対策: 接続線・グループを含めた合計個数上限とグループのネスト
        段数上限を設ける
    - 図形・接続線・グループ(グループ内部の子孫要素を含む)の合計個数を共通の
      `MaxShapesPerSheet`でカウントする(画像は引き続き独立の`MaxImagesPerSheet`でカウント)。
      グループのネスト段数の上限(既定5段、`ElementKind = "GroupNestingTooDeep"`。
      トップレベルのグループ自身を1段目とする)を`MaxShapeNestingDepth`として追加した
    - _Requirements: 10.8_
  - [x] 15.11 Layoutレイヤー: 接続線のページ座標変換を実装する(`ConnectorCommand`生成)
    - 画像・図形と共通の`TryComputeDrawingObjectRect`をそのまま流用する
    - _Requirements: 10.9_
  - [x] 15.12 Layoutレイヤー: グループの子座標空間からページ座標への変換を実装する
    - グループ自身のページ矩形を求めたうえで、`ChildOffset`/`ChildExtent`から
      各子要素の`LocalRect`を比例変換(非一様倍率)しページ座標へ変換する再帰処理
      (`BuildGroupChildren`/`ToGroupChildRect`)を実装し、結果を`GroupCommand`にまとめた。
      子座標空間の大きさが0以下、または変換後の矩形が異常に大きい場合の安全弁も設けた
    - _Requirements: 10.10_
  - [x] 15.13 Layoutレイヤー: グループ内図形のテキスト折り返しを既存ロジックで対応する
    - `GroupChildShape.Text`を`BuildShapeTextLines`/`WrapShapeText`と同じロジックで
      折り返す(`BuildGroupChildShapeCommand`がトップレベルの`BuildShapeCommand`と
      同じ処理を再利用する)
    - _Requirements: 10.10, 10.4_
  - [x] 15.14 Renderingレイヤー: `DrawPage`のコマンド振り分けを再利用可能なヘルパーへ
        切り出す
    - `GroupCommand`の内部展開から個々の子コマンドを描画する際に同じ振り分けロジックを
      再帰的に使うための準備として、既存の`switch`文を`DrawSingleCommand`へ抽出した
      (抽象レコード型`DrawCommand`と紛らわしくなるため、型名とは別名にした)
    - _Requirements: 10.10_
  - [x] 15.15 Renderingレイヤー: 追加プリセット(星形・フローチャート記号)のパス生成を実装する
    - `star4`/`star5`/`star6`/`star8`は外接円半径と内側頂点の半径比から交互に結ぶ2N角形
      (`StarPath`)、フローチャート記号7種はそれぞれ専用のパス生成メソッド
      (`DiamondPath`/`StadiumPath`/`ParallelogramPath`/`DocumentPath`/`PredefinedProcessPath`、
      `flowChartProcess`/`flowChartConnector`は既存の`RectPath`/`EllipsePath`を再利用)を実装した。
      `pdftoppm`でラスタライズして目視確認済み
    - _Requirements: 10.1_
  - [x] 15.16 Renderingレイヤー: 追加の吹き出し(`cloudCallout`, `callout1`/`callout2`/`callout3`)の
        パス生成を実装する
    - 雲形(`CloudCalloutPath`)は円の和集合(`SKPath.Op(SKPathOp.Union)`)+
      wedge系と共通化した引き出し三角形(`AddWedgeTail`)、引き出し線付き吹き出しは
      N本の折れ線を汎用ロジック(`BuildLeaderPoints`)で生成した。callout1/2/3は
      本体(塗りつぶし対象)と引き出し線(塗りつぶし無し)でジオメトリが異なるため、
      `ShapeGeometryBuilder`に塗りつぶし用の`Build`とは別に枠線用の`BuildOutline`を
      新設し、`SkiaPdfRenderer.DrawShape`をFill/Outlineで別々のパスを使うよう変更した
    - _Requirements: 10.1_
  - [x] 15.17 Renderingレイヤー: 多段階・放射状グラデーションの描画を実装する
    - `SKShader.CreateLinearGradient`/`CreateRadialGradient`に複数ストップを渡すよう
      `CreateShapeFillPaint`を拡張した(`ToShaderStops`ヘルパーで位置昇順にソート)
    - _Requirements: 10.6_
  - [x] 15.18 Renderingレイヤー: 接続線のパス生成(`ConnectorGeometryBuilder`)と描画を実装する
    - `ShapeGeometryBuilder`とは別に新設。`Outline`が無い接続線には既定の黒い実線1ptを補う
      (`DefaultConnectorOutline`)。`pdftoppm`でラスタライズして5種のプリセットを目視確認済み
    - _Requirements: 10.9_
  - [x] 15.19 Renderingレイヤー: グループの描画(`GroupCommand`)を実装する
    - `canvas.Save`/`RotateDegrees(Center)`/子コマンドの再帰描画(`DrawSingleCommand`)/
      `Restore`で実装した。子要素の個別回転との合成を`pdftoppm`でラスタライズして
      目視確認し、想定どおり正しく合成されることを確認した
    - _Requirements: 10.10, 10.5_
  - [x] 15.20 `Utsushi.SampleGenerator` に接続線・グループ・追加プリセット・多段階/放射状
        グラデーションの埋め込み機能を追加する
    - `SetGradientShape`(多段階線形/放射状グラデーション)、`SetConnector`、`SetGroup`
      (`GroupChildShapeSpec`による子座標空間上の子図形指定)を追加した。追加プリセット
      (星形・フローチャート記号等)は既存の`SetShape`が`A.ShapeTypeValues`を直接受け取る
      ため追加変更不要だった
    - _Requirements: 8.3_
  - [x] 15.21 サンプル帳票に拡張分の図形を配置し、ゴールデンテストを更新する
    - invoiceサンプルの6行目(件名行と合計行の間の空白行)に、線形グラデーション星形・
      放射状グラデーションのflowChartTerminator・接続線(bentConnector3)・回転付き
      グループ(楕円+矩形の2要素)を配置した。CLIで実際にPDFを生成し`pdftoppm`で
      ラスタライズして目視確認し、既存の図形群(9行目)と重ならないことを確認した。
      `GoldenSnapshot`が`ConnectorCommand`/`GroupCommand`を`unknown`としてしか
      記述できていなかったため、専用の記述(接続線: rect/preset/rotation/flip/outline、
      グループ: center/rotation/children数+子コマンドを再帰的にインデント記述)を追加した
      うえでゴールデンファイルを更新した
    - _Requirements: 10.1, 10.6, 10.9, 10.10, 8.3_
  - [x] 15.22 Parsing層のユニットテストを追加する
    - 接続線・グループの読み取り、グループ内非対応要素の検出(型不明・非対応プリセット)、
      合計個数/ネスト段数の上限(境界値を含む)、多段階/放射状グラデーションの読み取りを
      `ConnectorAndGroupReadingTests.cs`で検証した(test-writerが作成、26件)
    - _Requirements: 10.1, 10.6, 10.8, 10.9, 10.10_
  - [x] 15.23 Layout層のユニットテストを追加する
    - 接続線の座標変換、グループの子座標空間変換(`ChildOffset`が非ゼロの場合の平行移動、
      非一様倍率、`ChildExtent`が0以下の壊れた入力での安全な空振り)、入れ子グループの
      再帰変換(2段ネストを手計算した期待値で検証)、改ページ境界での配置を
      `ConnectorAndGroupLayoutTests.cs`で検証した(test-writerが作成、10件)
    - _Requirements: 10.9, 10.10_
  - [x] 15.24 Rendering層のユニットテストを追加する
    - 追加プリセットのパス生成(星形・フローチャート記号のBounds、cloudCallout/calloutの
      引き出し部分)、callout1/2/3のBuild(本体のみ)とBuildOutline(引き出し線含む)の
      差異、それ以外のプリセットではBuildとBuildOutlineが一致すること、接続線5種の
      経路生成(反転による端点入れ替えを含む)、GroupCommand/ConnectorCommandを含む
      PagedLayoutのレンダリングが例外なく完了することを
      `ShapeGeometryBuilderTests.cs`/`ConnectorGeometryBuilderTests.cs`/
      `SkiaPdfRendererTests.cs`で検証した(test-writerが作成、Rendering層のテスト計163件)
    - _Requirements: 10.1, 10.6, 10.9, 10.10_
  - [x] 15.25 レビュー対応
    - `code-reviewer`(shapeCount上限のグループ経由での超過、画像検証順序の後退、
      TooManyShapesメッセージの不正確さを修正)、`security-reviewer`(グラデーション
      ストップ数無制限を修正、OpenXml SDK再帰のStackOverflowリスクをdesign.mdに記録)、
      `layout-fidelity-reviewer`(座標変換ロジックは問題なしと確認、Layout/Renderingの
      ユニットテスト不足を指摘→15.23/15.24で対応)の指摘にすべて対応した。
      最後に`doc-reviewer`でrequirements.md/design.md/tasks.mdの最終整合性を確認し、
      補足の掲載順(要件10.1→10.6→10.7→10.8→10.9→10.10)と「下記補足」のあいまいな
      使い回しをrequirements.mdで修正した(Critical該当なし)
    - _Requirements: 10.8, 10.9, 10.10, 6.4_

- [ ] 16. 接続線の接続点(コネクションサイト)解決と星形・雲形吹き出しの近似精度向上
  - [ ] 16.1 Parsingレイヤー: 図形・画像・グループのID読み取り
    - `NonVisualDrawingProperties/@id`を`ShapeModel`/`ImageModel`/`GroupShapeModel`、
      グループ内の`GroupChildShape`/`GroupChildImage`/`GroupChildGroup`に`Id: uint`
      として追加する(`GroupChildConnector`は不要)。既存の`ReadShape`/`ReadImage`/
      `ReadGroupShape`/`ReadGroupChildShape`/`ReadGroupChildImage`/`ReadGroupChildGroup`
      を拡張する
    - _Requirements: 10.11_
  - [ ] 16.2 Parsingレイヤー: 接続線の接続点参照(`stCxn`/`endCxn`)を読み取る
    - `ConnectionRef(uint ShapeId, uint SiteIndex)`を追加し、`ConnectorModel`/
      `GroupChildConnector`に`StartConnection`/`EndConnection: ConnectionRef?`を追加する。
      `ReadConnector`/`ReadGroupChildConnector`で`xdr:cNvCxnSpPr`配下の`a:stCxn`/`a:endCxn`
      (`@id`+`@idx`)を読み取る。要素が無ければ`null`
    - _Requirements: 10.11_
  - [ ] 16.3 Parsingレイヤー: 星形(star4/5/6/8)の既定内側半径比をプリセットごとに修正する
    - `DefaultStarInnerRadiusRatio`(単一の0.38)を廃止し、`star4`=0.25、`star5`=0.382、
      `star6`=0.577、`star8`=0.75をプリセットごとの定数として`ShapeGeometryBuilder`に
      持たせる(design.md「未決事項」の推定値であることの注記を残す)
    - _Requirements: 10.12_
  - [ ] 16.4 Parsingレイヤー: 雲形吹き出し(cloudCallout)の引き出し位置調整ガイドを読み取る
    - `ShapeAdjustmentGuideNames[ShapePresetType.CloudCallout]`を`["adj1", "adj2"]`に変更し、
      `CloudCalloutPath`のシグネチャに`adjustmentValues`を追加して`wedgeRectCallout`等と
      同じ`Adj`ヘルパーで読み取る(既定値は変更しない)
    - _Requirements: 10.13_
  - [ ] 16.5 Layoutレイヤー: 描画オブジェクトのID→ページ矩形解決テーブルを構築する
    - `PageCommandBuilder.Build`が画像・図形・グループ(接続線を除く)を処理する際、
      `Dictionary<uint, (RectPt Rect, ShapePresetType? Preset)>`(ページごとに独立)を
      同時に組み立てる。グループ内要素は`ToGroupChildRect`変換後の最終ページ矩形を記録する
    - _Requirements: 10.11_
  - [ ] 16.6 Layoutレイヤー: 接続点(コネクションサイト)を解決する
    - `ConnectionSiteResolver.Resolve(rect, preset, siteIndex)`(既定は矩形の上下左右の
      中点。`flowChartInputOutput`/`flowChartDocument`は実際の輪郭に合わせて補正)を実装し、
      接続線(トップレベル・グループ内)を16.5のテーブル構築後にまとめて処理して
      `ConnectorCommand.ResolvedStart`/`ResolvedEnd`を設定する。解決できない場合は
      `null`のままにする(要件10.9の既定動作へのフォールバックはRenderingレイヤーの責務)
    - _Requirements: 10.11_
  - [ ] 16.7 Renderingレイヤー: 接続線の描画で解決済み接続点を優先する
    - `ConnectorCommand.ResolvedStart`/`ResolvedEnd`が両方とも非nullの場合、
      `ConnectorGeometryBuilder.Build`がこの2点を始点・終点として使うよう拡張する
      (`Rect`/`FlipHorizontal`/`FlipVertical`は無視する)
    - _Requirements: 10.11_
  - [ ] 16.8 サンプル帳票への配置とゴールデンテスト更新
    - `SpreadsheetBuilder`に、IDを指定して図形と接続線を関連付けるAPI(`SetShape`等が
      返す/受け取るID、または`stCxn`/`endCxn`を組み立てる`SetConnector`の拡張)を追加し、
      invoiceサンプルに接続点解決が効くケース(フローチャート記号同士を接続)を1つ配置する。
      CLIでPDFを生成し`pdftoppm`でラスタライズして目視確認したうえでゴールデンファイルを
      更新する
    - _Requirements: 10.11, 8.3_
  - [ ] 16.9 Parsingレイヤーのユニットテストを追加する
    - ID読み取り、`stCxn`/`endCxn`の読み取り(要素の有無両方)、星形プリセットごとの
      既定内側半径比、雲形吹き出しの`adj1`/`adj2`読み取りを検証する
    - _Requirements: 10.11, 10.12, 10.13_
  - [ ] 16.10 Layoutレイヤーのユニットテストを追加する
    - 同一ページ内での接続点解決成功、参照先が異なるページにある場合のフォールバック、
      グループ内要素を参照先とする解決、`flowChartInputOutput`/`flowChartDocument`の
      補正、それ以外のプリセット・画像・グループでの4方向近似、参照先ID不在時の
      フォールバックを検証する
    - _Requirements: 10.11_
  - [ ] 16.11 Renderingレイヤーのユニットテストを追加する
    - `ConnectorGeometryBuilder.Build`が`ResolvedStart`/`ResolvedEnd`指定時にそれを
      使うこと(`Rect`/フラグを無視すること)、星形の内側半径比がプリセットごとに
      異なること、雲形吹き出しの引き出し位置が`adjustmentValues`に応じて変わることを検証する
    - _Requirements: 10.11, 10.12, 10.13_
  - [ ] 16.12 レビュー対応
    - `code-reviewer`/`layout-fidelity-reviewer`/`security-reviewer`の指摘に対応する
      (ID解決テーブルの構築コスト、接続点解決が改ページ・グループネストと絡む場合の
      エッジケースを重点的に確認する)
    - _Requirements: 10.11, 10.12, 10.13_
