# 実装計画: 入口の Excel2Pdf への一本化と診断の強化

対象要件: `.kiro/specs/excel2pdf-consolidation/requirements.md`
対象設計: `.kiro/specs/excel2pdf-consolidation/design.md`

> **状況**: 仕様の作成段階。要件・設計のレビューを経てから実装に着手する。
>
> **順序の方針**:
> - 中核仕様を最初に改訂する(CLAUDE.md「先に仕様側を更新する」)。
> - どの時点で止めても、今より欠落の検知が弱くならないようにする。そのため、新しい診断の仕組み(2〜4)、回帰テストの移行(5)、
>   CLIの作り直し(6)を先に行い、帳票定義の削除(7)はその後にする。
> - 各タスクの終わりで `dotnet build`・`dotnet test` が緑であることを保つ(タスクの途中でビルドを壊したまま次へ進まない)。

- [ ] 1. 中核仕様の改訂
  - [ ] 1.1 `.kiro/specs/excel-report-pdf-conversion/requirements.md` を、requirements.md「中核仕様への影響」のとおり改める。
        帳票定義を削除するまで(タスク7)残る機能の要件(1.4, 2.1〜2.6 など)には、「廃止予定(excel2pdf-consolidation のタスク7で削除)」と注記する
    - _Requirements: 1.3, 1.5, 2.1, 2.7_
  - [ ] 1.2 同 design.md・tasks.md を改める(tasks.md は完了済みのタスクを残し、本仕様で廃止・変更する旨を注記する)
  - [ ] 1.3 spec-compliance-reviewer・doc-reviewer

- [ ] 2. Parsing: 未対応要素の収集(2.1〜2.4は1つの単位として、ビルドとテストを緑にしてから完了とする)
  - [ ] 2.1 `Utsushi.Core` に `UnsupportedElementKind`・`UnsupportedElement` を追加する。`UnsupportedWorkbookElementException` に `Elements`・`IsTruncated` を持たせ、`ElementKind` を削除する
    - _Requirements: 2.2, 2.4_
  - [ ] 2.2 `UnsupportedElementCollector`(記録の単位・件数の上限・`Message`/`Detail` の正規化。上限を渡せる internal のオーバーロード)を追加し、`WorkbookModel` に `UnsupportedElements`・`UnsupportedElementsTruncated` を加える
    - _Requirements: 2.4, 2.5_
  - [ ] 2.3 `OpenXmlWorkbookReader` の扱いの指定を見ている全箇所(23+2)を収集に置き換え、描画オブジェクトの検知をアンカーごとにする。黙って読み飛ばす経路を記録する。
        グループの非対応の子要素の種類を直す。`WorkbookReadOptions.UnsupportedElementBehavior` を削除する。
        帳票定義ありの経路では、`ReportPdfConverter.ReadWorkbook`(`Convert*`・`ComputeLayout*`・`CheckFit*` が共通に通る)で、`unsupportedElements: "error"` のときに一覧から例外を作る(design.md 2.4)
    - _Requirements: 2.2, 2.3, 2.6, 2.7_
  - [ ] 2.4 テスト(design.md 6.3 の Parsing の観点)。既存の Parsing のテストの `Error`・`ElementKind` の検証を、一覧の検証に書き換える
    - _Requirements: 2.4, 2.5, 2.6, 2.7_

- [ ] 3. Rendering: フォントの代替の問い合わせ
  - [ ] 3.1 `FontResolver` の代替の記録を辞書にし、`GetSubstitution` を公開する
    - _Requirements: 3.1, 3.3, 3.4_
  - [ ] 3.2 テスト(厳格モード・代替後・外字の補い)
    - _Requirements: 3.1, 3.3, 3.4_

- [ ] 4. ファサード: エンジン・未対応要素の扱い・診断
  - [ ] 4.1 `Excel2PdfEngine` を追加する(`ReportPdfConverter.CreateDefault` 相当。部品と `FontResolver` を受け取るコンストラクタは internal)。`Excel2Pdf` にエンジンを受け取るコンストラクタを追加する
    - _Requirements: 1.1, 1.2_
  - [ ] 4.2 `Excel2PdfOptions.UnsupportedElements`(既定 `Error`)を追加し、保存の時点で読み込みの直後に判定する。旧い同名の列挙は完全修飾名で参照する。
        既存の `Excel2PdfTests` のうち未対応要素を含むサンプルを既定の設定で開くテストに `Ignore` を明示する
    - _Requirements: 2.1, 2.2, 2.3_
  - [ ] 4.3 `Excel2Pdf.Check()` と `Excel2PdfDiagnostics`・`FontSubstitution` を追加する(レイアウトの結果のフォントを集めて問い合わせる)
    - _Requirements: 3.1, 3.2, 4.1, 4.2, 4.3, 4.4_
  - [ ] 4.4 テスト(design.md 6.3 のファサードの観点。エンジンの共有・破棄を含む)。`edge-order-form` で既定のエラー・`Check()`・`Ignore` を確かめる
    - _Requirements: 1.2, 2.1, 2.2, 2.3, 3.2, 4.2, 6.3_
  - [ ] 4.5 security-reviewer(収集の上限・`Message`/`Detail` に文字列を含めないことと正規化・`src/Utsushi.Parsing/OpenXml/` 全体の、黙って読み飛ばす経路と未対応ループの横断確認)、code-reviewer
    - _Requirements: 2.4, 2.5, 2.6, 2.7_

- [ ] 5. ゴールデンテストの移行(帳票定義の削除より前)
  - [ ] 5.1 `tools/Utsushi.SampleGenerator`・`tools/edge-case-samples/generate.py` で、書き込むセルに `shrinkToFit`・`wrapText` を付け、テンプレートを再生成する
    - _Requirements: 6.2_
  - [ ] 5.2 `ReportConversionGoldenTests` を `Excel2Pdf` とセル番地の指定に書き換える(`MaxDigitWidthPx`・`DocumentName` を明示する)。
        スナップショットの差分が design.md 6.1 の2種類だけであることを確かめ、layout-fidelity-reviewer のレビューを経て更新する。
        同じテンプレートを使う他のテストの結果も確かめる
    - _Requirements: 6.1, 6.2_
  - [ ] 5.3 `EdgeCaseSampleTests` をセル番地の指定に書き換える(置換キーだけの観点は削除し、ignore/error のテストは 4.4 に置き換える)
    - _Requirements: 6.3_

- [ ] 6. CLI の作り直し(帳票定義の削除より前。`Excel2Pdf` だけで作れる)
  - [ ] 6.1 `check`・`convert` コマンドと `CommandLineOptions.Parse`・`Program.Run` を実装し、帳票定義のオプションを削除する
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8_
  - [ ] 6.2 テスト(引数の解釈、`Run` の終了コードと出力)
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7_

- [ ] 7. 帳票定義の仕組みの削除(7.1〜7.5は1つの単位として、ビルドとテストを緑にしてから完了とする)
  - [ ] 7.1 ファサード: `ReportPdfConverter` を internal の `ConversionPipeline` にし、帳票コード付きのコンストラクタ・`SetField`・`ClearField`・`ReportCode`・`CheckFit` を削除する
    - _Requirements: 1.1, 1.3, 1.4, 4.5_
  - [ ] 7.2 ReportDefinition: JSONの読み込み・リポジトリ・置換キーの定義・旧い `UnsupportedElementPolicy` を削除し、`ReportDefinition` を `ConversionSettings` に縮める。`samples/**/definition.json` を削除する
    - _Requirements: 1.5_
  - [ ] 7.3 Substitution・Layout: 置換キー・必須キー・`OverflowByCell`・`PrintAreaOverride`・`FitIssue.SubstitutionKey`・`FitCheckResult` を削除する
    - _Requirements: 1.3_
  - [ ] 7.4 Core・Layout・Parsing: 帳票定義・置換キーの例外を削除し、`ReportCode` を `DocumentName` に改める(`UtsushiException`・`PagedLayout`・`WorkbookReadOptions`。ゴールデンの1行目が変わる)
    - _Requirements: 1.3, 1.6_
  - [ ] 7.5 帳票定義に依存するテストを削除・書き換える(`tests/Utsushi.ReportDefinition.Tests` の JSON・リポジトリのテスト、Substitution の置換キーのテスト、`Excel2PdfTests` の帳票定義ありのテストなど)
    - _Requirements: 1.3, 1.5_
  - [ ] 7.6 code-reviewer・layout-fidelity-reviewer・security-reviewer。`dotnet build --no-incremental -p:UtsushiVs2019Compiler=true` でも確かめる
    - _Requirements: 1.1_

- [ ] 8. ドキュメントと steering
  - [ ] 8.1 `.kiro/steering/structure.md`(構成・帳票の追加手順)・`tech.md`、`CLAUDE.md`(開発コマンド・現在の状態・必ず守る制約)、`README.md` を改め、design.md 7 のとおり方針を書き換える。
        product.md・structure.md・tech.md・CLAUDE.md の「移行中」の注記を外す
    - _Requirements: 1.3, 5.8_
  - [ ] 8.2 `docs/ライブラリの使い方.md`・`docs/テンプレート作成ガイド.md`・`docs/エッジケース検証レポート.md`・`samples/edge-cases/README.md` を改め、`docs/帳票定義スキーマ.md` を削除する。`docs/実装設計失敗事例集.md` の廃止した機能の節に注記を加える。
        `docs/ライブラリの使い方.md` には、既定がエラーに変わったこと・`ReportCode` の改名と、移行前に `check` で確かめる手順を書く
    - _Requirements: 1.6, 2.1, 4.1, 5.1_
  - [ ] 8.3 `.claude/agents/` の帳票定義への言及を改める
    - _Requirements: 1.5_
  - [ ] 8.4 中核仕様の「廃止予定」の注記を、廃止済みに改める
  - [ ] 8.5 doc-reviewer・spec-compliance-reviewer(本仕様・中核仕様と実装の整合)
