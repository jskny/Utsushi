---
inclusion: always
---

# Utsushi プロジェクト構成方針

> 本リポジトリは現時点でドキュメント/ルール整備のみが行われた段階であり、コードベースは未着手。
> 以降の実装は本方針に従ってディレクトリ・プロジェクトを作成すること。

## レイヤー構成

自社帳票のPDF化パイプラインは以下の一方向の依存関係を持つレイヤーに分割する。

```
Excel解析レイヤー (Parsing)
        ↓
帳票定義レイヤー (ReportDefinition)
        ↓
値差し込み(置換)レイヤー (Substitution)
        ↓
レイアウト計算レイヤー (Layout)
        ↓
PDF描画レイヤー (Rendering)
```

- 上位レイヤー(Rendering側)は下位レイヤー(Parsing側)の実装詳細(例: OpenXml SDKの型)を直接知ってはならない。レイヤー間の受け渡しは各レイヤーが定義する内部モデル(POCO)を介する。
- 帳票ごとの差異(セル位置、置換対象キー、既知の特殊処理)は「帳票定義レイヤー」のデータ(JSON等)に閉じ込め、他レイヤーのコードに帳票固有の分岐を持ち込まない。

## 想定ソリューション構成

```
Utsushi/
├── src/
│   ├── Utsushi.Parsing/            # Excel(OOXML)構造の読み取り。DocumentFormat.OpenXml に依存してよい唯一のプロジェクト
│   ├── Utsushi.ReportDefinition/   # 帳票定義モデルとローダー(JSON等の定義ファイル → 内部モデル)
│   ├── Utsushi.Substitution/       # 指定セルへの文字列置換ロジック
│   ├── Utsushi.Layout/             # 改ページ・印刷範囲・列幅行高・フォントメトリクスに基づくレイアウト計算
│   ├── Utsushi.Rendering/          # SkiaSharp を用いたPDF描画。Layoutレイヤーの計算結果のみを入力とする
│   └── Utsushi.Cli/ (or .Api/)     # エントリーポイント(CLIまたはAPI)。各レイヤーを組み立てて実行する
├── tests/
│   ├── Utsushi.Parsing.Tests/
│   ├── Utsushi.ReportDefinition.Tests/
│   ├── Utsushi.Substitution.Tests/
│   ├── Utsushi.Layout.Tests/
│   ├── Utsushi.Rendering.Tests/
│   └── Utsushi.Golden.Tests/       # 登録済み帳票サンプルによるゴールデン(回帰)テスト
├── samples/
│   └── reports/                    # テスト用の帳票サンプル(実データはマスキング済みのものに限る)
├── docs/                           # 人が読む補足ドキュメント(必要に応じて)
├── .kiro/
│   ├── steering/                   # 本ファイル群。プロジェクト全体に常時適用される方針
│   └── specs/                      # 機能ごとの要件定義書・設計書・タスクリスト
├── .claude/
│   └── agents/                     # コードレビュー・テスト作成・セキュリティレビュー等のサブエージェント定義
└── Utsushi.sln
```

## 命名規則

- 名前空間・プロジェクト名は `Utsushi.<レイヤー名>` とする。
- 帳票定義ファイルは `reports/<帳票コード>/definition.json` のように帳票コード単位でディレクトリを分ける。
- テストのゴールデンファイルは `tests/Utsushi.Golden.Tests/Fixtures/<帳票コード>/` に配置する。

## 帳票の追加手順(標準ワークフロー)

1. `samples/reports/` に対象Excelファイル(個人情報等はマスキング)を追加する。
2. `Utsushi.ReportDefinition` に帳票定義(印刷範囲・置換対象セル・既知の特殊処理)を追加する。
3. `tests/Utsushi.Golden.Tests` にレビュー済みの期待PDF(またはページ記述ログ)を追加する。
4. 新規要件が既存のレイアウト計算で表現できない場合のみ、`.kiro/specs/` に新しい要件を追記し、`Utsushi.Layout` を拡張する。

## ドキュメントとサブエージェントの関係

- 新機能・大きな変更に着手する前に、`.kiro/specs/<機能名>/` に `requirements.md` → `design.md` → `tasks.md` の順で仕様を作成・更新する(spec駆動開発)。
- 実装完了後は `code-reviewer`、テスト未整備の場合は `test-writer`、外部入力を扱う変更(Excel読み込み・ファイルパス処理等)には `security-reviewer`、仕様との整合性確認には `spec-compliance-reviewer` を用いる。各サブエージェントの詳細は `.claude/agents/` を参照。
