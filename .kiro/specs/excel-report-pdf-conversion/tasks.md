# 実装計画: 自社帳票のExcel→PDF変換

対象要件: `.kiro/specs/excel-report-pdf-conversion/requirements.md`
対象設計: `.kiro/specs/excel-report-pdf-conversion/design.md`

> **状況**: 全タスク完了(タスク1〜11: 2026-09-19、タスク12: 2026-09-20、タスク13: 2026-09-21)。
> `dotnet build` / `dotnet test`(224件、2026-09-21時点)/ `dotnet format` はグリーン。
> 実装時に決定した事項・判明した制約は `design.md` に反映済み。
> タスク文面どおりに実現できなかった項目には各タスクに注記を付けた。

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
