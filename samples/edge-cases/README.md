# エッジケース検証用の帳票サンプル

`samples/reports/` の合成サンプルは、Utsushi が対応している機能を狙って配置した「お手本どおりの」テンプレートである。
ここにあるのは、帳票の担当者が普段の Excel 操作で作りがちなテンプレートを再現したサンプルで、
極端な差し込み値と組み合わせてライブラリの挙動を確かめるために使う。検証の結果は `docs/エッジケース検証レポート.md` にまとめている。

- `template.xlsx` は `tools/edge-case-samples/generate.py`(Python + openpyxl)の生成物。`definition.json` は手書き。
- 自動テストは `tests/Utsushi.Golden.Tests/EdgeCaseSampleTests.cs`(`dotnet test` で一緒に実行される)。
- 登録済み帳票(`samples/reports/`)とは別のディレクトリに置いている。CLI で試すときは `--definitions samples/edge-cases` を指定する。

## サンプルの一覧

| 帳票コード | シート | 主に確かめること |
|---|---|---|
| `edge-cover-letter` | 送付状 | 差し込み値の極端な値。長い宛名(縮小)・部署(切り取り)・担当者名(はみ出し)・3行ぶんの住所(折り返し、縦位置は Excel 既定の「下」)・5行ぶんの備考、横位置「標準」のままの金額セル、英字フォント(Calibri)のセル・テンプレートに存在しないセル・列全体の書式だけのセルへの差し込み |
| `edge-sales-list` | 売上一覧 | 「すべての列を1ページに印刷」(横1×縦自動)・横向き・150行の明細の自動改ページ・見出し行の繰り返し・非表示の列(社外秘の原価)・非表示の行・高さ0の行・手動改ページ(縮小印刷では無視される)・数式(キャッシュ値)・曜日付きの日付と和暦の書式・SUBTOTAL |
| `edge-schedule` | 工程表 | 行と列の両方向の改ページ・印刷タイトル行と列・ページの方向「左から右」・行と列の手動改ページ・拡大縮小90%・ページ中央・先頭ページ番号(5)・先頭ページのみ別指定のヘッダー/フッター・90度回転した日付見出し・改ページをまたぐ結合セル |
| `edge-order-form` | 注文書 | 表紙・注文書・非表示のマスタの3シート構成(変換対象は2枚目)、テーマの色、文字列の表示形式 `@" 御中"`、入力規則(プルダウン)、リッチテキスト、数式のキャッシュ値・エラー値(`#DIV/0!`)・差し込みセルを参照する数式、条件付き書式、縦書き・均等割り付け・選択範囲内で中央、網かけ・グラデーション・斜線・各種線種、各種の表示形式(`[h]:mm`・分数・指数・`#,##0,"千円"`・`[Red]` 付きの負数など)、コメント、ハイパーリンク、PNG/JPEG 画像、グラフ |

すべてのブックは日本語版 Excel の新規ブックと同じく標準フォントを游ゴシック 11pt(標準の行の高さ 18.75pt)にしている。
このため帳票定義の `maxDigitWidthPx` は 8 にしている。

## CLI で試す

游ゴシックの無い環境では `--allow-font-fallback` を付ける(同梱の BIZ UDPゴシックで代替する。見た目は Excel と一致しない)。

```bash
dotnet run --project src/Utsushi.Cli -- \
  --report edge-cover-letter --definitions samples/edge-cases \
  --input samples/edge-cases/edge-cover-letter/template.xlsx --output /tmp/edge-cover-letter.pdf \
  --allow-font-fallback \
  --set PostalCode="〒100-0001" --set Address="東京都千代田区千代田1-1-1" --set CompanyName="株式会社サンプル商事" \
  --set SendDate="2026年4月1日" --set DocumentNo="0012" --set Sender="株式会社Utsushi" \
  --set Item1="ご請求書" --set Item1Copies="1"
```

## テンプレートを作り直す

```bash
pip install -r tools/edge-case-samples/requirements.txt
python3 tools/edge-case-samples/generate.py samples/edge-cases
```

生成結果は毎回同じバイト列になる(ZIP の日時と文書プロパティの日時を固定している)。

openpyxl は Excel と次の点が異なるため、生成後にシートの XML を書き換えて Excel で保存したブックと同じ状態にしている。

- 数式の計算結果(キャッシュ値)を保存しない → 計算結果を書き込む(Utsushi は数式を再計算せず、保存時の計算結果を出力するため)。
- 高さ0の行を書き出さない → `ht="0"` を書き込む。

ただし Excel 実機で保存したブックではないため、Excel が書き出す要素(テーマの細部・`x14` 拡張・プリンター設定など)を
すべて再現しているわけではない。実帳票の登録時は、従来どおり Excel で保存したテンプレートで確かめる。
