# 実装計画: 入口の Excel2Pdf への一本化と診断の強化

対象要件: `.kiro/specs/excel2pdf-consolidation/requirements.md`
対象設計: `.kiro/specs/excel2pdf-consolidation/design.md`

> **状況**: 仕様の作成段階。要件・設計のレビューを経てから実装に着手する。
>
> **順序の方針**: どの時点で止めても、今より欠落の検知が弱くならないようにする。まず新しい診断の仕組みを足す(1〜3)。
> 次に回帰テストを帳票定義に依存しない形に移す(4)。帳票定義を削除するのは、その後にする(5)。

- [ ] 1. Parsing: 未対応要素の収集
  - [ ] 1.1 `Utsushi.Core` に `UnsupportedElementKind`・`UnsupportedElement` を追加し、`UnsupportedWorkbookElementException` に `Elements`・`IsTruncated` を持たせる
    - _Requirements: 2.2, 2.4_
  - [ ] 1.2 `UnsupportedElementCollector`(件数の上限、テスト用に上限を渡せる internal のオーバーロード)を追加し、`WorkbookModel` に `UnsupportedElements`・`UnsupportedElementsTruncated` を加える
    - _Requirements: 2.5_
  - [ ] 1.3 `OpenXmlWorkbookReader` の `Error` 分岐の throw を収集に置き換え、グラフ・図表枠・レガシー図形・外部参照の検知を常に行う。
        `WorkbookReadOptions.UnsupportedElementBehavior` を削除する。
        帳票定義ありの経路(この時点ではまだ残っている)では、ファサードが `unsupportedElements: "error"` のときに一覧から例外を作る
    - _Requirements: 2.2, 2.3, 2.6, 2.7_
  - [ ] 1.4 黙って捨てていた壊れた描画オブジェクトを `MalformedDrawingObject` として記録する。グループの非対応の子要素の種類を `UnsupportedGroupChild` に直す
    - _Requirements: 2.7_
  - [ ] 1.5 テスト(各箇所の種類・シート・アンカー・Detail、上限超過の1件化、件数の上限、読まないシートの要素を含まないこと)。既存の Parsing のテストの `Error` の検証を、一覧の検証に書き換える
    - _Requirements: 2.4, 2.5, 2.6, 2.7_

- [ ] 2. Rendering: フォントの代替の問い合わせ
  - [ ] 2.1 `FontResolver` の代替の記録を辞書にし、`TryGetSubstitution` を公開する
    - _Requirements: 3.1, 3.3, 3.4_
  - [ ] 2.2 テスト(厳格モード・代替後・外字の補い)
    - _Requirements: 3.1, 3.3, 3.4_

- [ ] 3. ファサード: エンジン・未対応要素の扱い・診断
  - [ ] 3.1 `Excel2PdfEngine` を追加する(`ReportPdfConverter.CreateDefault` 相当。部品を差し替えるコンストラクタは internal、`InternalsVisibleTo` でテストに許す)。`Excel2Pdf` にエンジンを受け取るコンストラクタを追加する
    - _Requirements: 1.2_
  - [ ] 3.2 `Excel2PdfOptions.UnsupportedElements`(既定 `Error`)を追加し、保存の時点で読み込みの直後に判定する
    - _Requirements: 2.1, 2.2, 2.3_
  - [ ] 3.3 `Excel2Pdf.Check()` と `Excel2PdfDiagnostics`・`FontSubstitution` を追加する(レイアウトの結果のフォントを集めて問い合わせる)
    - _Requirements: 3.1, 3.2, 4.1, 4.2, 4.3, 4.4_
  - [ ] 3.4 テスト(既定でエラー・全件・Ignore・`Check()` は例外にしない・2回目の変換でも代替が返る・使わないフォントは返らない)。`edge-order-form` で確かめる
    - _Requirements: 2.1, 2.2, 2.3, 3.2, 4.2, 6.3_
  - [ ] 3.5 security-reviewer(収集の上限・メッセージに文字列を含めないこと・`src/Utsushi.Parsing/OpenXml/` 全体の未対応ループの横断確認)
    - _Requirements: 2.4, 2.5, 2.6_

- [ ] 4. ゴールデンテストの移行(帳票定義の削除より前)
  - [ ] 4.1 `tools/Utsushi.SampleGenerator` で、サンプル帳票の差し込みセルに `shrinkToFit`・`wrapText` を付け、`samples/reports/*/template.xlsx` を再生成する
    - _Requirements: 6.2_
  - [ ] 4.2 `ReportConversionGoldenTests` を `Excel2Pdf` とセル番地の指定に書き換える。スナップショットの差分が design.md 6.1 の3種類だけであることを確かめ、layout-fidelity-reviewer のレビューを経て更新する
    - _Requirements: 6.1, 6.2_
  - [ ] 4.3 `EdgeCaseSampleTests` をセル番地の指定に書き換える(置換キーだけの観点は削除し、ignore/error のテストは 3.4 に置き換える)
    - _Requirements: 6.3_

- [ ] 5. 帳票定義の仕組みの削除
  - [ ] 5.1 ファサード: `ReportPdfConverter` を internal の `ConversionPipeline` にし、帳票コード付きのコンストラクタ・`SetField`・`ClearField`・`ReportCode`・`CheckFit`・`FitCheckResult` を削除する
    - _Requirements: 1.3, 1.4, 4.5_
  - [ ] 5.2 ReportDefinition: JSONの読み込み・リポジトリ・置換キーの定義を削除し、`ReportDefinition` を `ConversionSettings` に縮める。`samples/**/definition.json` を削除する
    - _Requirements: 1.5_
  - [ ] 5.3 Substitution・Layout: 置換キー・必須キー・`OverflowByCell`・`PrintAreaOverride`・`FitIssue.SubstitutionKey` を削除する
    - _Requirements: 1.3_
  - [ ] 5.4 Core: 帳票定義・置換キーの例外を削除し、`ReportCode` を `DocumentName` に改める(`PagedLayout` を含む。ゴールデンの1行目が変わる)
    - _Requirements: 1.3_
  - [ ] 5.5 帳票定義に依存するテストを削除・書き換える(`tests/Utsushi.ReportDefinition.Tests` の JSON・リポジトリのテスト、Substitution の置換キーのテスト、`Excel2PdfTests` の帳票定義ありのテストなど)
    - _Requirements: 1.3, 1.5_
  - [ ] 5.6 code-reviewer・layout-fidelity-reviewer・security-reviewer。`dotnet build -p:UtsushiVs2019Compiler=true` でも確かめる
    - _Requirements: 1.1_

- [ ] 6. CLI の作り直し
  - [ ] 6.1 `check`・`convert` コマンドと `CommandLineOptions.Parse`・`Program.Run` を実装する
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7, 5.8_
  - [ ] 6.2 テスト(引数の解釈、`Run` の終了コードと出力)
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 5.7_

- [ ] 7. ドキュメントと steering
  - [ ] 7.1 `.kiro/steering/structure.md`(構成・帳票の追加手順)・`tech.md`、`CLAUDE.md`(開発コマンド・現在の状態)、`README.md` を改める
    - _Requirements: 1.3, 5.8_
  - [ ] 7.2 `docs/ライブラリの使い方.md`・`docs/テンプレート作成ガイド.md`・`docs/エッジケース検証レポート.md`・`samples/edge-cases/README.md` を改め、`docs/帳票定義スキーマ.md` を削除する。`docs/実装設計失敗事例集.md` の廃止した機能の節に注記を加える。
        `docs/ライブラリの使い方.md` には、既定がエラーに変わったことと、移行前に `check` で確かめる手順を書く
    - _Requirements: 2.1, 4.1, 5.1_
  - [ ] 7.3 `.claude/agents/` の帳票定義への言及を改める
    - _Requirements: 1.5_
  - [ ] 7.4 doc-reviewer

- [ ] 8. 中核仕様の改訂
  - [ ] 8.1 `.kiro/specs/excel-report-pdf-conversion/requirements.md` を requirements.md「中核仕様への影響」のとおり改める
  - [ ] 8.2 同 design.md・tasks.md を改める(tasks.md は完了済みのタスクを残し、本仕様で廃止した旨を注記する)
  - [ ] 8.3 spec-compliance-reviewer(本仕様・中核仕様と実装の整合)
