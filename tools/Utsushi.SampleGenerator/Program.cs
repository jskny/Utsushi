using System;
using System.Globalization;
using System.IO;
using static Utsushi.SampleGenerator.SampleStyles;

namespace Utsushi.SampleGenerator;

/// <summary>
/// テスト用の帳票サンプル(.xlsx)を生成する。
/// </summary>
/// <remarks>
/// <para>
/// 実運用の帳票テンプレートは社内データであり、このリポジトリには含められない。
/// そこで、実帳票と同等の構造(結合セル・罫線・数値書式・印刷範囲・印刷タイトル・
/// 手動/自動改ページ)を持つ<b>合成サンプル</b>をここで生成し、
/// 開発とゴールデンテストの対象とする。
/// </para>
/// <para>
/// 生成するのは <c>template.xlsx</c> のみで、<c>definition.json</c>(帳票定義)は
/// 人が編集するデータとしてリポジトリに直接コミットしている。
/// </para>
/// <para>
/// 実行方法:
/// <code>dotnet run --project tools/Utsushi.SampleGenerator -- samples/reports</code>
/// </para>
/// </remarks>
internal static class Program
{
    private static int Main(string[] args)
    {
        var outputRoot = args.Length > 0 ? args[0] : Path.Combine("samples", "reports");
        Directory.CreateDirectory(outputRoot);

        GenerateInvoice(Path.Combine(outputRoot, "invoice"));
        GenerateDeliveryNote(Path.Combine(outputRoot, "delivery-note"));
        GenerateReceipt(Path.Combine(outputRoot, "receipt"));

        Console.WriteLine($"帳票サンプルを生成しました: {Path.GetFullPath(outputRoot)}");
        return 0;
    }

    /// <summary>
    /// 請求書サンプル: 1ページに収まる帳票。結合セル・罫線・数値書式・印刷範囲を含む。
    /// </summary>
    private static void GenerateInvoice(string directory)
    {
        Directory.CreateDirectory(directory);

        var builder = new SpreadsheetBuilder("請求書")
        {
            PrintArea = "$A$1:$F$34",
        };

        // 列幅(文字数単位)
        builder.SetColumnWidth(1, 4.0);    // A: 行番号
        builder.SetColumnWidth(2, 28.0);   // B: 品名
        builder.SetColumnWidth(3, 8.0);    // C: 数量
        builder.SetColumnWidth(4, 6.0);    // D: 単位
        builder.SetColumnWidth(5, 12.0);   // E: 単価
        builder.SetColumnWidth(6, 14.0);   // F: 金額

        builder.SetRowHeight(1, 30.0);
        builder.SetRowHeight(2, 8.0);
        builder.SetRowHeight(8, 24.0);

        // 表題
        builder.Merge("A1:F1");
        builder.SetText(1, 1, "請求書", Style.Title);

        // 宛先・発行情報
        builder.Merge("A3:C3");
        builder.SetText(3, 1, "株式会社サンプル商事 御中", Style.CustomerName);

        builder.SetText(3, 5, "請求番号", Style.LabelRight);
        builder.SetText(3, 6, "INV-0000", Style.Value);

        builder.SetText(4, 5, "発行日", Style.LabelRight);
        builder.SetNumber(4, 6, ToSerial(new DateTime(2026, 4, 1)), Style.DateValue);

        builder.SetText(5, 5, "支払期限", Style.LabelRight);
        builder.SetNumber(5, 6, ToSerial(new DateTime(2026, 4, 30)), Style.DateValue);

        builder.Merge("A5:C5");
        builder.SetText(5, 1, "件名: サンプル案件", Style.SectionLabel);

        // 合計金額の強調表示
        builder.Merge("A8:B8");
        builder.SetText(8, 1, "ご請求金額", Style.SectionLabel);
        builder.Merge("C8:F8");
        builder.SetNumber(8, 3, 0, Style.TotalCurrency);
        // 結合範囲の右端の罫線はF8自身のRightに保持されるため、アンカー(C8)だけでなく
        // 範囲内の全セルに同じ箱罫線スタイルを設定する(D8:F8)。
        builder.SetStyleOnly(8, 4, Style.TotalCurrency);
        builder.SetStyleOnly(8, 5, Style.TotalCurrency);
        builder.SetStyleOnly(8, 6, Style.TotalCurrency);

        // 明細表
        const int headerRow = 11;
        builder.SetText(headerRow, 1, "No.", Style.TableHeader);
        builder.SetText(headerRow, 2, "品名", Style.TableHeader);
        builder.SetText(headerRow, 3, "数量", Style.TableHeader);
        builder.SetText(headerRow, 4, "単位", Style.TableHeader);
        builder.SetText(headerRow, 5, "単価", Style.TableHeader);
        builder.SetText(headerRow, 6, "金額", Style.TableHeader);

        var items = new (string Name, double Quantity, string Unit, double UnitPrice)[]
        {
            ("サンプル設計費", 1, "式", 320000),
            ("サンプル実装費", 24, "人日", 48000),
            ("サンプル試験費", 8, "人日", 42000),
            ("ドキュメント作成", 1, "式", 150000),
        };

        var subtotal = 0.0;
        for (var i = 0; i < 12; i++)
        {
            var row = headerRow + 1 + i;
            builder.SetStyleOnly(row, 1, Style.TableTextCentered);
            builder.SetStyleOnly(row, 2, Style.TableText);
            builder.SetStyleOnly(row, 3, Style.TableNumber);
            builder.SetStyleOnly(row, 4, Style.TableTextCentered);
            builder.SetStyleOnly(row, 5, Style.TableCurrency);
            builder.SetStyleOnly(row, 6, Style.TableCurrency);

            if (i >= items.Length)
            {
                continue;
            }

            var item = items[i];
            var amount = item.Quantity * item.UnitPrice;
            subtotal += amount;

            builder.SetNumber(row, 1, i + 1, Style.TableTextCentered);
            builder.SetText(row, 2, item.Name, Style.TableText);
            builder.SetNumber(row, 3, item.Quantity, Style.TableNumber);
            builder.SetText(row, 4, item.Unit, Style.TableTextCentered);
            builder.SetNumber(row, 5, item.UnitPrice, Style.TableCurrency);
            builder.SetNumber(row, 6, amount, Style.TableCurrency);
        }

        // 小計・消費税・合計
        var tax = Math.Floor(subtotal * 0.1);
        var totals = new (string Label, double Value)[]
        {
            ("小計", subtotal),
            ("消費税(10%)", tax),
            ("合計", subtotal + tax),
        };

        for (var i = 0; i < totals.Length; i++)
        {
            var row = 25 + i;
            builder.Merge($"D{row}:E{row}");
            builder.SetText(row, 4, totals[i].Label, Style.TableTextCentered);
            builder.SetStyleOnly(row, 5, Style.TableTextCentered);
            builder.SetNumber(row, 6, totals[i].Value, Style.TableCurrency);
        }

        // 備考
        builder.Merge("A29:F33");
        builder.SetText(
            29, 1,
            "備考:\n・お振込手数料は貴社にてご負担をお願いいたします。\n・本請求書はテスト用の合成サンプルです。",
            Style.Note);

        builder.Save(Path.Combine(directory, "template.xlsx"), Create());
    }

    /// <summary>
    /// 納品書サンプル: 複数ページにわたる帳票。印刷タイトル行の繰り返しと自動改ページを検証する。
    /// </summary>
    private static void GenerateDeliveryNote(string directory)
    {
        Directory.CreateDirectory(directory);

        var builder = new SpreadsheetBuilder("納品書")
        {
            // 1〜7行目(表題・宛先・明細見出し)を各ページの先頭に繰り返す。
            PrintTitleRows = "$1:$7",
            PrintArea = "$A$1:$E$90",

            // 複数ページになる帳票なので、フッターにページ番号を入れる。
            OddHeader = "&R&A",
            OddFooter = "&L&D&C&P / &N ページ&R&\"MS PGothic,Bold\"サンプル",
        };

        builder.SetColumnWidth(1, 5.0);
        builder.SetColumnWidth(2, 34.0);
        builder.SetColumnWidth(3, 9.0);
        builder.SetColumnWidth(4, 7.0);
        builder.SetColumnWidth(5, 14.0);

        builder.SetRowHeight(1, 28.0);

        builder.Merge("A1:E1");
        builder.SetText(1, 1, "納品書", Style.Title);

        builder.Merge("A3:C3");
        builder.SetText(3, 1, "株式会社サンプル商事 御中", Style.CustomerName);

        builder.SetText(3, 4, "納品番号", Style.LabelRight);
        builder.SetText(3, 5, "DN-0000", Style.Value);

        builder.SetText(4, 4, "納品日", Style.LabelRight);
        builder.SetNumber(4, 5, ToSerial(new DateTime(2026, 4, 15)), Style.DateValue);

        const int headerRow = 7;
        builder.SetText(headerRow, 1, "No.", Style.TableHeader);
        builder.SetText(headerRow, 2, "品名", Style.TableHeader);
        builder.SetText(headerRow, 3, "数量", Style.TableHeader);
        builder.SetText(headerRow, 4, "単位", Style.TableHeader);
        builder.SetText(headerRow, 5, "備考", Style.TableHeader);

        // 1ページに収まらない行数を意図的に入れ、自動改ページと印刷タイトルの繰り返しを発生させる。
        for (var i = 0; i < 80; i++)
        {
            var row = headerRow + 1 + i;
            builder.SetNumber(row, 1, i + 1, Style.TableTextCentered);
            builder.SetText(row, 2, $"サンプル品目 {(i + 1).ToString("000", CultureInfo.InvariantCulture)}", Style.TableText);
            builder.SetNumber(row, 3, ((i % 7) + 1) * 3, Style.TableNumber);
            builder.SetText(row, 4, "個", Style.TableTextCentered);
            builder.SetStyleOnly(row, 5, Style.TableText);
        }

        builder.Save(Path.Combine(directory, "template.xlsx"), Create());
    }

    /// <summary>
    /// 領収書サンプル: 1シートに「本紙」と「控え」を別々の印刷範囲として持つ帳票。
    /// 複数の印刷範囲(要件3.6)がそれぞれ独立したページになることを検証する。
    /// </summary>
    private static void GenerateReceipt(string directory)
    {
        Directory.CreateDirectory(directory);

        var builder = new SpreadsheetBuilder("領収書")
        {
            // 本紙(1〜13行目)と控え(16〜28行目)を別々の印刷範囲にする。
            // 範囲の間にある14〜15行目(区切りの注記)は印刷されない。
            PrintArea = "$A$1:$D$13,$A$16:$D$28",
            OddFooter = "&C- &P -",
        };

        builder.SetColumnWidth(1, 14.0);
        builder.SetColumnWidth(2, 24.0);
        builder.SetColumnWidth(3, 14.0);
        builder.SetColumnWidth(4, 16.0);

        // 本紙と控えは同じ構成なので、開始行だけ変えて2回書く。
        WriteReceiptBlock(builder, startRow: 1, title: "領 収 書");
        WriteReceiptBlock(builder, startRow: 16, title: "領 収 書(控)");

        // 印刷範囲の外に置く注記。PDFに出てはいけない。
        builder.SetText(14, 1, "※この行は印刷範囲外のため出力されない", Style.Note);

        builder.Save(Path.Combine(directory, "template.xlsx"), Create());
    }

    private static void WriteReceiptBlock(SpreadsheetBuilder builder, int startRow, string title)
    {
        builder.SetRowHeight(startRow, 30.0);

        builder.Merge($"A{startRow}:D{startRow}");
        builder.SetText(startRow, 1, title, Style.Title);

        builder.Merge($"A{startRow + 2}:B{startRow + 2}");
        builder.SetText(startRow + 2, 1, "株式会社サンプル商事 様", Style.CustomerName);

        builder.SetText(startRow + 2, 3, "発行日", Style.LabelRight);
        builder.SetNumber(startRow + 2, 4, ToSerial(new DateTime(2026, 4, 20)), Style.DateValue);

        builder.SetText(startRow + 4, 1, "金額", Style.SectionLabel);
        builder.Merge($"B{startRow + 4}:D{startRow + 4}");
        builder.SetNumber(startRow + 4, 2, 0, Style.TotalCurrency);
        // 結合範囲の右端の罫線はD列自身のRightに保持されるため、アンカー(B)だけでなく
        // 範囲内の全セルに同じ箱罫線スタイルを設定する。
        builder.SetStyleOnly(startRow + 4, 3, Style.TotalCurrency);
        builder.SetStyleOnly(startRow + 4, 4, Style.TotalCurrency);

        builder.SetText(startRow + 6, 1, "但し", Style.SectionLabel);
        builder.Merge($"B{startRow + 6}:D{startRow + 6}");
        builder.SetText(startRow + 6, 2, "サンプル代金として", Style.TableText);
        builder.SetStyleOnly(startRow + 6, 3, Style.TableText);
        builder.SetStyleOnly(startRow + 6, 4, Style.TableText);

        builder.SetText(startRow + 8, 1, "上記正に領収いたしました。", Style.Value);

        builder.Merge($"C{startRow + 10}:D{startRow + 11}");
        builder.SetText(startRow + 10, 3, "株式会社Utsushi\n東京都サンプル区1-2-3", Style.Note);
    }

    /// <summary>Excel のシリアル値(1900年日付システム)へ変換する。</summary>
    private static double ToSerial(DateTime date)
    {
        var epoch = new DateTime(1899, 12, 30);
        return (date - epoch).TotalDays;
    }
}
