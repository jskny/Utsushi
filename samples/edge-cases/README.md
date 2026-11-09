# エッジケース検証用の帳票サンプル

`samples/reports/` の合成サンプルは、Utsushi が対応している機能を狙って配置した「お手本どおりの」テンプレートである。
ここにあるのは、帳票の担当者が普段の Excel 操作で作りがちなテンプレートを再現したサンプルで、
極端な差し込み値と組み合わせてライブラリの挙動を確かめるために使う。検証の結果は `docs/エッジケース検証レポート.md` にまとめている。

- `template.xlsx` は `tools/edge-case-samples/generate.py`(Python + openpyxl)の生成物。
- 自動テストは `tests/Utsushi.Golden.Tests/EdgeCaseSampleTests.cs`(`dotnet test` で一緒に実行される)。書き込むセルの番地は同テストの
  `CoverLetterValues` などにまとめている。
- 書き込むセルのはみ出しの扱いは、テンプレートのセルの書式で表している(注文書の宛先 A3・売上一覧の部署名 J2 は「縮小して全体を表示する」)。

## サンプルの一覧

| サンプル名 | シート | 主に確かめること |
|---|---|---|
| `edge-cover-letter` | 送付状 | 差し込み値の極端な値。長い宛名(縮小)・部署・担当者名(はみ出し)・3行ぶんの住所(折り返し、縦位置は Excel 既定の「下」)・5行ぶんの備考、横位置「標準」のままの金額セル、英字フォント(Calibri)のセル・テンプレートに存在しないセル・列全体の書式だけのセルへの差し込み、行全体を塗りつぶした帯(2行目) |
| `edge-sales-list` | 売上一覧 | 「すべての列を1ページに印刷」(横1×縦自動)・横向き・150行の明細の自動改ページ・見出し行の繰り返し・非表示の列(社外秘の原価)・非表示の行・高さ0の行・手動改ページ(縮小印刷では無視される)・数式(キャッシュ値)・曜日付きの日付と和暦の書式・SUBTOTAL |
| `edge-schedule` | 工程表 | 行と列の両方向の改ページ・印刷タイトル行と列・ページの方向「左から右」・行と列の手動改ページ・拡大縮小90%・ページ中央・先頭ページ番号(5)・先頭ページのみ別指定のヘッダー/フッター・90度回転した日付見出し・改ページをまたぐ結合セル |
| `edge-order-form` | 注文書 | 表紙・注文書・非表示のマスタの3シート構成(変換対象は2枚目)、テーマの色、文字列の表示形式 `@" 御中"`、入力規則(プルダウン)、リッチテキスト、数式のキャッシュ値・エラー値(`#DIV/0!`)・差し込みセルを参照する数式、条件付き書式、縦書き・均等割り付け・選択範囲内で中央、網かけ・グラデーション・斜線・各種線種、各種の表示形式(`[h]:mm`・分数・指数・`#,##0,"千円"`・`[Red]` 付きの負数など)、コメント、ハイパーリンク、PNG/JPEG 画像、グラフ(未対応要素のため、既定では変換がエラーになる) |

すべてのブックは日本語版 Excel の新規ブックと同じく標準フォントを游ゴシック 11pt(標準の行の高さ 18.75pt)にしている。
列幅の換算に使う最大数字幅は、標準フォントからの見積もりと同じ 8 になる(テストでは `MaxDigitWidthPx = 8` を明示している)。

## CLI で試す

游ゴシックの無い環境では `--allow-font-fallback` を付ける(同梱の BIZ UDPゴシックで代替する。見た目は Excel と一致しない)。

```bash
# 診断する(注文書はグラフ・コメントを未対応要素として知らせる)
dotnet run --project src/Utsushi.Cli -- check samples/edge-cases/edge-order-form/template.xlsx --allow-font-fallback

# 送付状に値を書き込んで変換する
dotnet run --project src/Utsushi.Cli -- \
  convert samples/edge-cases/edge-cover-letter/template.xlsx /tmp/edge-cover-letter.pdf --allow-font-fallback \
  --text A3="〒100-0001" --text A4="東京都千代田区千代田1-1-1" --text A7="株式会社サンプル商事" \
  --text H3="2026年4月1日" --text H4="0012" --text F5="株式会社Utsushi" \
  --text B15="ご請求書" --text G15="1"

# 注文書を、グラフ・コメントを除いて変換する
dotnet run --project src/Utsushi.Cli -- \
  convert samples/edge-cases/edge-order-form/template.xlsx /tmp/edge-order-form.pdf --allow-font-fallback --ignore-unsupported
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
