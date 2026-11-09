# 実装計画: 入口の Excel2Pdf への一本化と診断の強化

対象要件: `.kiro/specs/excel2pdf-consolidation/requirements.md`
対象設計: `.kiro/specs/excel2pdf-consolidation/design.md`

> **状況**: 実装中。実装前の実現可能性の調査で決めたことは design.md 9 にまとめた。
>
> **順序の方針**:
> - 中核仕様を最初に改訂する(CLAUDE.md「先に仕様側を更新する」)。
> - どの時点で止めても、今より欠落の検知が弱くならないようにする。そのため、新しい診断の仕組み(2〜5)、回帰テストの移行(6)、
>   CLIの作り直し(7)を先に行い、帳票定義の削除(8)はその後にする。
> - **各サブタスクの終わりで `dotnet build`・`dotnet test` が緑であることを保つ**(サブタスクの途中でビルドを壊したまま次へ進まない)。
>   そのため、既存の動作を変えずに足す → 中身を移す → 古いものを外す、の順に並べている。
> - テストは、対応する実装のサブタスクに含める(「テスト」とだけ書いたサブタスクは、横断的なテストに限る)。

## 1. 中核仕様の改訂

- [x] 1.1 中核の requirements.md の「はじめに」「用語」と要件1・2を、requirements.md「中核仕様への影響」のとおり改める。
      帳票定義を削除するまで(タスク8)残る機能の要件(1.4, 1.11, 2.1, 2.3〜2.6, 2.12)には「廃止予定(excel2pdf-consolidation のタスク8で削除)」と注記する
  - _Requirements: 1.3, 1.5, 2.1_
- [x] 1.2 中核の requirements.md の要件3〜11(3.7・3.8の補足、3.10、4.5、4.12、6.3〜6.15、8、9.4〜9.6、10.7・10.8・10の補足)を改める
  - _Requirements: 2.6, 2.7, 7.2_
- [x] 1.3 中核の requirements.md の要件12〜14を改める
  - _Requirements: 1.1, 1.2, 4.5, 5.8_
- [x] 1.4 中核の design.md を改める。「未決事項」の近似する書式には本仕様の要件7で検知する旨を加え、
      「帳票定義ありの経路では登録済みのシート名に合うファイルしか通らない」という入力の前提の記述を、任意の .xlsx を受け付ける唯一の経路になることに合わせて改める
  - _Requirements: 1.1, 7.2_
- [x] 1.5 中核の tasks.md に、本仕様で廃止・変更するタスクの注記を加える(完了済みのタスクは残す)
- [x] 1.6 spec-compliance-reviewer・doc-reviewer

## 2. Parsing: 未対応要素の収集

既存の扱いの指定(`UnsupportedElementBehavior`)を残したまま、無視の側で記録を始め(2.1〜2.7。エラーの側の判定の順序と `ElementKind` は変えない)、最後に指定を外す(2.8)。

- [x] 2.1 `Utsushi.Core` に `UnsupportedElementKind`・`UnsupportedElement` と、説明の文の正規化 `DiagnosticText.Normalize` を追加する(まだ使わない)
  - _Requirements: 2.4_
- [x] 2.2 `UnsupportedElementCollector` を追加する(記録の単位、件数の上限、`Message`/`Detail` の正規化。上限を渡せる internal のオーバーロード)。単体テストを含む
  - _Requirements: 2.4, 2.5_
- [x] 2.3 `WorkbookModel` に `UnsupportedElements`・`UnsupportedElementsTruncated` を加え、`OpenXmlWorkbookReader` で収集器を1回の読み込みごとに作って載せる(この時点では常に空)
  - _Requirements: 2.2_
- [x] 2.4 画像の未対応要素を記録する(非対応形式・内容の不一致・参照先が無い/空/画像でない・枚数の上限・大きさの上限)。無視の側で記録し、エラーの側の動作は変えない。テストを含む
  - _Requirements: 2.5, 2.7_
- [x] 2.5 図形・接続線・グループの未対応要素を記録する(非対応プリセット・`custGeom`・数の上限・文字数の上限・入れ子の段数)。
      グループ全体を出力しない場合はグループのアンカーで `UnsupportedGroupChild` を1件だけ記録し、子要素の側では重ねて記録しない。種類の誤り(1623行)を直す。テストを含む
  - _Requirements: 2.5, 2.7_
- [x] 2.6 描画オブジェクトの検知をアンカーごとにする(グラフ・図表枠・その他の非対応の要素・`absoluteAnchor`・解決できない描画パート)。
      `legacyDrawing`・`legacyDrawingHF`・外部参照・結合セル数の上限を記録する。検知を省く2か所(`!= Error`)の早期の戻りを外す。テストを含む
  - _Requirements: 2.2, 2.5, 2.7_
- [x] 2.7 黙って読み飛ばしている経路(design.md 2.2)を `MalformedDrawingObject` として記録する(`absoluteAnchor` と重ねない)。テストを含む
  - _Requirements: 2.7_
- [x] 2.8 `UnsupportedWorkbookElementException` に `Elements`・`IsTruncated` を持たせ、`ElementKind` を削除する(`ToDiagnosticString()` に一覧を含めるフックを `UtsushiException` に加える)。`WorkbookReadOptions.UnsupportedElementBehavior` を削除し、
      読み込みは常に記録だけにする。帳票定義ありの経路は `ReportPdfConverter.ReadWorkbook` で `unsupportedElements: "error"` のときに一覧から例外を作る(design.md 2.4)。
      既存の Parsing のテストの `Error`・`ElementKind` の検証を、一覧の検証に書き換える
  - _Requirements: 2.2, 2.3, 2.6_
- [x] 2.9 security-reviewer(収集の上限・`Message`/`Detail` の正規化・`src/Utsushi.Parsing/OpenXml/` 全体の、黙って読み飛ばす経路と未対応ループの横断確認)、code-reviewer
  - _Requirements: 2.4, 2.5, 2.6, 2.7_

## 3. Rendering: フォントの代替の問い合わせ

- [x] 3.1 `FontResolver` の代替の記録(`_substituted`)を、代替した書体のファミリ名を持つ辞書にする(外字の補いの判定は辞書のキーで行い、動作は変えない)
  - _Requirements: 3.4_
- [x] 3.2 `GetSubstitution(FontStyle)` を公開する(問い合わせの時点で解決する)。テスト(厳格モード・代替後・外字の補い)を含む
  - _Requirements: 3.1, 3.3, 3.4_

## 4. ファサード: エンジン・未対応要素の扱い・診断

- [x] 4.1 `Excel2PdfEngine` を追加する(`Create`・`Dispose`。内部に `ReportPdfConverter` を持つ)。部品と `FontResolver` を受け取る internal のコンストラクタを加える
  - _Requirements: 1.2, 1.4_
- [x] 4.2 `Excel2Pdf` にエンジンを受け取るコンストラクタ(パス・バイト列・ストリーム)を追加する。テスト(`Fonts`・`Render` の指定はエラー、破棄してもエンジンは使える、複数のスレッドで共有できる)を含む
  - _Requirements: 1.1, 1.2_
- [x] 4.3 `UnsupportedElementPolicy`(名前空間 `Utsushi`)と `Excel2PdfOptions.UnsupportedElements` を追加する(既定は `Ignore` のまま。旧い同名の列挙は完全修飾名で参照する)
  - _Requirements: 2.1_
- [x] 4.4 保存の時点で、読み込みの直後に未対応要素を判定する。テスト(エラーで全件・`Ignore` で変換できる)を含む
  - _Requirements: 2.2, 2.3_
- [x] 4.5 既定を `Error` に変える。既存の `Excel2PdfTests` のうち未対応要素を含むサンプルを既定の設定で開くテストに `Ignore` を明示する。`edge-order-form` で既定のエラーを確かめる
  - _Requirements: 2.1, 6.3_
- [x] 4.6 `Excel2Pdf.Check()` と `Excel2PdfDiagnostics` を追加する(文字の収まらない箇所・未対応要素。未対応要素では例外にしない)。テスト(要件4.2〜4.4)を含む
  - _Requirements: 4.1, 4.2, 4.3, 4.4_
- [x] 4.7 `FontSubstitution` を追加し、`Check()` でレイアウトの結果のフォントを集めて問い合わせる。テスト(2回目の変換でも返る・使わないフォントは返らない)を含む
  - _Requirements: 3.1, 3.2, 4.1_
- [x] 4.8 code-reviewer

## 5. 書式の近似の検知(design.md 3.5)

骨組み(5.1〜5.3)を先に作り、種類ごとに検知を足す(5.4〜5.13)。各サブタスクは Parsing の読み取り・Layout の記録・テスト(エッジケースのサンプルを含む)と、近似の行が変わる場合のゴールデンの更新をまとめて行う。

- [x] 5.1 `Utsushi.Core` に `FormatApproximationKind`・`FormatApproximation` を追加する
  - _Requirements: 7.1_
- [x] 5.2 Layout: `FormatApproximationCollector`(種類ごとの上限・要素の識別を含む重複の判定)を追加し、`checkFit` のときだけ作って `PagedLayout` に載せる(この時点では常に空)。単体テストを含む
  - _Requirements: 7.4, 7.5_
- [x] 5.3 ファサード: `Excel2PdfDiagnostics` に `FormatApproximations`・`FormatApproximationsTruncated`・`HasFormatApproximations` を加える。ゴールデンのスナップショットに近似の行を出す(この時点では差分なし)
  - _Requirements: 4.1, 4.3, 7.1_
- [x] 5.4 セルの配置: `Horizontal/VerticalAlignment`(読み済み)と `Indent`(読み済み)を記録する
  - _Requirements: 7.2, 7.3_
- [x] 5.5 文字の回転: セルの `textRotation`・図形の文字の向き(`bodyPr/@vert`・`@rot`)を読み、`TextRotation` を記録する
  - _Requirements: 7.2, 7.3_
- [ ] 5.6 条件付き書式: `conditionalFormatting`(拡張の `x14` を含む)の範囲を読み(ブロックの数・1つの範囲指定に含む範囲の数の上限)、印刷範囲と重なる部分を記録する
  - _Requirements: 7.2, 7.3_
- [ ] 5.7 リッチテキスト: ランの書式がセルの書式と異なる印を読み、記録する。`SetText`・`SetValue` で上書きしたセルの印を消す
  - _Requirements: 7.2, 7.3_
- [ ] 5.8 表示形式: 解釈できない書式・表せない値の印を `NumberFormatter` で載せ、記録する。`SetValue` 用に `NumberFormatting` の internal のオーバーロードを加え、上書きしたセルは書き込んだ値で求め直す
  - _Requirements: 7.2, 7.3_
- [ ] 5.9 セルの塗りつぶし: 網かけ・グラデーションの印を `StyleTable` で読み、`EmitCell` で記録する(セルが無い位置を行・列の書式で描く場合を含む)
  - _Requirements: 7.2, 7.3, 7.4_
- [ ] 5.10 図形の塗りつぶし・書体: テーマの `fillStyleLst` を読み、`ShapeFill` の印(design.md 3.5 の4つの場合)と `fontRef` の印を載せて記録する。既定のテーマの `fillRef idx=1` には印を付けない
  - _Requirements: 7.2, 7.3_
- [ ] 5.11 画像: 反転(`flipH`・`flipV`)を `ImageModel`・`GroupChildImage` に読み、`ImageFlip`・`ImageRotation` を記録する(反転したグループの中の画像を含む)
  - _Requirements: 7.2, 7.3_
- [ ] 5.12 ヘッダー/フッター: `HeaderFooterCommandBuilder.ScaleRuns` で改行を取り除くときに `MultiLineHeaderFooter` を記録する
  - _Requirements: 7.2_
- [ ] 5.13 ページの分割: `MergedCellAcrossPageBreak`・`DrawingAcrossPageBreak`・`MergedCellHiddenEdge` を記録する
  - _Requirements: 7.2, 7.3, 7.4_
- [ ] 5.14 タスク5の開始前と比べて、ゴールデンの描画命令の行が変わっていないことを確かめる。保存(`checkFit` なし)では集めないことを確かめる
  - _Requirements: 7.5_
- [ ] 5.15 layout-fidelity-reviewer・security-reviewer(新しい読み取りの上限)・code-reviewer
  - _Requirements: 7.2, 7.3_

## 6. ゴールデンテストの移行(帳票定義の削除より前)

- [ ] 6.1 `tools/Utsushi.SampleGenerator` で、サンプル帳票の書き込むセルに `shrinkToFit`・`wrapText` を付けてテンプレートを再生成し、同じテンプレートを使う既存のテストが緑であることを確かめる
  - _Requirements: 6.2_
- [ ] 6.2 `tools/edge-case-samples/generate.py` で、エッジケースのサンプルの書き込むセルに `shrinkToFit` を付けて再生成し、既存のテストが緑であることを確かめる
  - _Requirements: 6.2_
- [ ] 6.3 `ReportConversionGoldenTests` を `Excel2Pdf` とセル番地の指定に書き換える(`MaxDigitWidthPx`・`DocumentName` を明示する)。
      スナップショットの差分が design.md 6.1 の2種類だけであることを確かめる
  - _Requirements: 6.1, 6.2_
- [ ] 6.4 layout-fidelity-reviewer のレビューを経て、ゴールデンを更新する
  - _Requirements: 6.2_
- [ ] 6.5 `EdgeCaseSampleTests` をセル番地の指定に書き換える(注文書は `UnsupportedElements = Ignore` も渡す)。置換キーだけの観点は削除し、ignore/error のテストは 4.4・4.5 に置き換える
  - _Requirements: 6.3_

## 7. CLI の作り直し(帳票定義の削除より前。`Excel2Pdf` だけで作れる)

- [ ] 7.1 `tests/Utsushi.Cli.Tests` に `Utsushi.TestSupport` への参照を加える。`CommandLineOptions.Parse` を `check`・`convert` のコマンドと新しいオプションに作り直し、帳票定義のオプションを削除する。引数の解釈のテストを含む
  - _Requirements: 5.4, 5.5, 5.6, 5.8_
- [ ] 7.2 `Program.Run` と `convert` を実装する(`Main` は `Run` を呼ぶだけにする。終了コード0/2/3)。テストを含む
  - _Requirements: 5.3, 5.6, 5.7_
- [ ] 7.3 `check` の出力(未対応要素・文字の収まらない箇所・フォントの代替・書式の近似(参考))と終了コード0/1を実装する。テストを含む
  - _Requirements: 5.1, 5.2_

## 8. 帳票定義の仕組みの削除

外側(公開API)から内側(モデル)へ、置換キーの定義を参照しているもの(Substitution・Layout)を先に外す。

- [ ] 8.1 `Excel2Pdf` の帳票コード付きのコンストラクタ・コンバータを受け取るコンストラクタ・`SetField`・`ClearField`・`ReportCode`・`CheckFit` を削除し、関係するテストを削除・書き換える
  - _Requirements: 1.3, 4.5_
- [ ] 8.2 `ReportPdfConverter` を internal の `ConversionPipeline` にし、帳票定義ありの経路と公開の変換メソッドを削除する。公開コンストラクタを使っていたテストを、エンジンの internal のコンストラクタに移す
  - _Requirements: 1.4_
- [ ] 8.3 Substitution: 置換キー(`ICellSubstitutor.Apply`)・必須キー・未知のキーの検証と、そのテストを削除する
  - _Requirements: 1.3_
- [ ] 8.4 Layout: `ValidateRequiredFieldsAreInPrintRanges`・`FindSubstitutionKey`・`OverflowByCell`・`PrintAreaOverride`・`FitIssue.SubstitutionKey`・`FitCheckResult` を削除する
      (`OverflowBehavior` は Layout の internal の列挙に移す。ゴールデンの `key=` の列が無くなる)。定義の `overflow` で確かめていた Layout のテストを、セルの書式に書き換えるか削除する
  - _Requirements: 1.3_
- [ ] 8.5 ReportDefinition: JSONの読み込み・リポジトリ・置換キーの定義・`ReportDefinition.UnsupportedElements` と旧い `UnsupportedElementPolicy`・
      `ReportModelBuilder` の置換キー・印刷範囲の上書きの検証を削除する。`samples/**/definition.json` と、`tests/Utsushi.ReportDefinition.Tests` の JSON・リポジトリ・`FieldIndexCacheTests` を削除する
  - _Requirements: 1.5_
- [ ] 8.6 Core: 帳票定義・置換キーの例外と `ProcessingStage.ReportDefinition` を削除する(`ReportModelBuilder` のシートの不在は `InvalidOperationException` にする)
  - _Requirements: 1.3_
- [ ] 8.7 `ReportCode` を `DocumentName` に改める(`UtsushiException`・`PagedLayout`・`WorkbookReadOptions`。ゴールデンの1行目が `document=` になる)
  - _Requirements: 1.6_
- [ ] 8.8 `ReportDefinition` を `ConversionSettings` に縮めて名前を変える
  - _Requirements: 1.5_
- [ ] 8.9 code-reviewer・layout-fidelity-reviewer・security-reviewer。`dotnet build --no-incremental -p:UtsushiVs2019Compiler=true` でも確かめる
  - _Requirements: 1.1_

## 9. ドキュメントと steering

- [ ] 9.1 `.kiro/steering/structure.md`(構成・帳票の追加手順)・`tech.md` を改め、design.md 7 のとおり方針を書き換える。product.md・structure.md・tech.md の「移行中」の注記を外す
  - _Requirements: 1.3_
- [ ] 9.2 `CLAUDE.md`(現在の状態・開発コマンド・必ず守る制約。「移行中」の注記を外す)と `README.md` を改める
  - _Requirements: 1.3, 5.8_
- [ ] 9.3 `docs/ライブラリの使い方.md` を改める(既定がエラーに変わったこと・`ReportCode` の改名・`Check()`・エンジン・移行前に `check` で確かめる手順)
  - _Requirements: 1.6, 2.1, 4.1_
- [ ] 9.4 `docs/テンプレート作成ガイド.md` を改める。「再現しない」の表に診断で検知できるもの(未対応要素・書式の近似)を示し、
      表に無い近似(分数・指数の表示形式、右揃えのインデント、図形の塗りつぶし・文字の向き・`fontRef` の書体、画像の反転)の行を加える。登録前のチェックリストを `check` に合わせる
  - _Requirements: 5.1, 7.2_
- [ ] 9.5 `docs/エッジケース検証レポート.md`・`samples/edge-cases/README.md` を改め、`docs/帳票定義スキーマ.md` を削除する。`docs/実装設計失敗事例集.md` の廃止した機能の節に注記を加える
  - _Requirements: 1.5_
- [ ] 9.6 `.claude/agents/` の帳票定義への言及を改める
  - _Requirements: 1.5_
- [ ] 9.7 中核仕様の「廃止予定」の注記を、廃止済みに改める
- [ ] 9.8 doc-reviewer・spec-compliance-reviewer(本仕様・中核仕様と実装の整合)
