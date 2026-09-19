# 実装計画: 自社帳票のExcel→PDF変換

対象要件: `.kiro/specs/excel-report-pdf-conversion/requirements.md`
対象設計: `.kiro/specs/excel-report-pdf-conversion/design.md`

- [ ] 1. ソリューション基盤のセットアップ
  - `.kiro/steering/structure.md` の構成に従い、`Utsushi.sln` と各レイヤーの `.NET 5` クラスライブラリプロジェクト(Parsing/ReportDefinition/Substitution/Layout/Rendering)、および対応するテストプロジェクトを作成する
  - Nullable参照型・ファイルスコープ名前空間を既定で有効化する共通 `Directory.Build.props` を用意する
  - _Requirements: 7.2, 7.3_

- [ ] 2. Parsingレイヤー: WorkbookModel の実装
  - [ ] 2.1 `DocumentFormat.OpenXml` を依存に追加し、`.xlsx` からセル値・スタイル(フォント/罫線/配置/数値書式/背景色)を読み取る `IWorkbookReader` を実装する
    - _Requirements: 1.1, 1.2, 7.1_
  - [ ] 2.2 列幅・行高・結合セル範囲の読み取りを実装する
    - _Requirements: 1.2_
  - [ ] 2.3 印刷範囲・手動改ページ・用紙サイズ/余白/拡大縮小・印刷タイトル・印刷順序の読み取りを実装する
    - _Requirements: 1.3_
  - [ ] 2.4 サポート外要素(図形・外部参照等)の検出とエラー化/無視の切り替えを実装する
    - _Requirements: 1.5_
  - [ ] 2.5 破損ファイル・非xlsx形式・パスワード保護ファイルに対する明確な例外を実装する
    - _Requirements: 6.1, 6.2_

- [ ] 3. ReportDefinitionレイヤー: 帳票定義のロードと突合
  - [ ] 3.1 帳票定義JSONのスキーマを定義し、ロード時のスキーマ検証を実装する
    - _Requirements: 6.3, 8.1_
  - [ ] 3.2 `WorkbookModel` と帳票定義を突合して `ReportModel` を構築する `IReportModelBuilder` を実装し、シート名/セル番地の不一致を検出する
    - _Requirements: 1.4_
  - [ ] 3.3 最初の対象帳票(例: 請求書)の帳票定義ファイルと、マスキング済みサンプルExcelを `samples/reports/` に追加する
    - _Requirements: 8.3_

- [ ] 4. Substitutionレイヤー: セル値の置換
  - [ ] 4.1 置換キー→値の辞書を受け取り、対象セルの値のみを更新する `ICellSubstitutor` を実装する(スタイルは変更しない)
    - _Requirements: 2.1, 2.2_
  - [ ] 4.2 未知の置換キー・必須キー欠落時の例外を実装する
    - _Requirements: 2.3, 2.4_
  - [ ] 4.3 はみ出し時の挙動フラグ(shrink/clip/wrap)を `ReportModel` に保持する処理を実装する(実際の描画反映はLayoutレイヤーで行う)
    - _Requirements: 2.5_

- [ ] 5. Layoutレイヤー: ページ分割と座標計算
  - [ ] 5.1 列幅(文字単位)・行高からポイント単位への換算ロジックをプロトタイプし、対象帳票での誤差を検証した上で実装する
    - _Requirements: 4.5_
  - [ ] 5.2 印刷範囲によるクリッピングを実装する
    - _Requirements: 3.1_
  - [ ] 5.3 手動改ページの適用を実装する
    - _Requirements: 3.2_
  - [ ] 5.4 自動改ページ計算(用紙サイズ・余白・拡大縮小率から導く印字可能領域に基づく行/列分割)を実装する
    - _Requirements: 3.3_
  - [ ] 5.5 印刷タイトル行/列の複製を実装する
    - _Requirements: 3.4_
  - [ ] 5.6 印刷順序(行優先/列優先)に従ったページ順序決定を実装する
    - _Requirements: 3.5_
  - [ ] 5.7 結合セルの矩形統合、罫線・背景色の描画命令生成を実装する
    - _Requirements: 4.2, 4.3_
  - [ ] 5.8 テキストの水平/垂直配置・インデント・縮小表示(shrink指定時)の座標計算を実装する
    - _Requirements: 4.1, 4.4, 2.5_

- [ ] 6. Renderingレイヤー: PDF描画
  - [ ] 6.1 `SkiaSharp` を依存に追加し、`PagedLayout` を `SKDocument` PDFバックエンドで描画する `IPdfRenderer` を実装する(背景→罫線→テキストの順で描画)
    - _Requirements: 5.1, 5.2, 5.3, 7.1_
  - [ ] 6.2 対象帳票が使用するフォントの埋め込み(サブセット化)を実装し、フォント未検出時をビルド/デプロイ時エラーとして検出する仕組みを用意する
    - _Requirements: 4.1_
  - [ ] 6.3 一時ストリームへの描画→正常終了時のみ最終出力先に確定させる仕組みを実装し、失敗時に不完全なPDFを残さないようにする
    - _Requirements: 5.4_

- [ ] 7. エラーハンドリング/ログの横断対応
  - [ ] 7.1 `UtsushiException` 基底クラスと各派生例外(帳票コード/シート名/セル番地/処理段階を保持)を実装する
    - _Requirements: 6.4_
  - [ ] 7.2 各レイヤーの例外送出箇所を7.1の例外型に統一する
    - _Requirements: 1.4, 1.5, 2.3, 2.4, 6.1, 6.2, 6.3_

- [ ] 8. テスト整備
  - [ ] 8.1 Layoutレイヤー(改ページ計算・単位換算)、Substitutionレイヤー(必須/未知キー判定)のユニットテストを作成する
    - _Requirements: 3.3, 2.3, 2.4_
  - [ ] 8.2 最初の対象帳票についてのエンドツーエンド(Parsing→Rendering)ゴールデンテストを作成する
    - _Requirements: 8.3_
  - [ ] 8.3 帳票定義JSONのスキーマ検証テストを作成する
    - _Requirements: 6.3_

- [ ] 9. エントリーポイントの実装
  - [ ] 9.1 5レイヤーを組み立てて「帳票コード + 置換値辞書 + 入力Excel」から「PDFストリーム」を得るユースケースクラス/CLIを実装する
    - _Requirements: 7.2, 7.3_

- [ ] 10. 2つ目以降の帳票への対応(拡張性の検証)
  - [ ] 10.1 共通レイヤーのコード変更なしに、2つ目の帳票定義とサンプル・ゴールデンテストを追加できることを確認する
    - _Requirements: 8.1, 8.2, 8.3_
