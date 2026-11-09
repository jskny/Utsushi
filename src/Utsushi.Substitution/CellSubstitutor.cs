using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution
{
    /// <summary>
    /// セル番地を指定した値の書き込みの既定実装(要件2.7、14.2・14.3)。
    /// </summary>
    /// <remarks>
    /// <para>書式(フォント・配置・罫線・数値書式)には一切触れない(要件2.2)。はみ出し時の挙動はセルの書式に従う。</para>
    /// </remarks>
    public sealed class CellSubstitutor : ICellSubstitutor
    {
        /// <summary>
        /// 差し込み値・セル上書き値1つあたりの最大文字数(UTF-16のコード単位数)。
        /// </summary>
        /// <remarks>
        /// Excel のセルに入れられる文字数の上限(32,767文字)に合わせている。これを超える値は元のExcelでも
        /// 表現できない帳票にとって異常な入力であり、上限が無いと巨大な文字列の折り返し計算・字形の検証・描画に
        /// 際限なくCPUとメモリを費やしてしまう(サービスから呼ばれた場合のサービス拒否の防止)。
        /// </remarks>
        internal const int MaxValueLength = 32767;

        /// <summary>
        /// 1回の変換で受け付けるセル番地の直接指定(セル上書き)の最大件数。
        /// </summary>
        /// <remarks>
        /// 帳票の差し込み箇所は多くても数百件程度であり、これを大きく超える件数は呼び出し元の不具合か
        /// 悪意のある入力である。値1つあたりの上限(<see cref="MaxValueLength"/>)だけでは件数×長さで
        /// メモリと処理時間を際限なく消費できてしまうため、件数にも上限を設ける。
        /// </remarks>
        internal const int MaxCellOverrideCount = 10000;

        /// <inheritdoc />
        public ReportModel ApplyNumericCellOverrides(ReportModel report, IReadOnlyDictionary<string, double> numericOverrides)
        {
            if (report is null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            if (numericOverrides is null)
            {
                throw new ArgumentNullException(nameof(numericOverrides));
            }

            if (numericOverrides.Count == 0)
            {
                return report;
            }

            // 件数の上限(要件2.16)は、文字列の直接指定(ApplyCellOverrides)と合わせて数える。
            var total = report.OverriddenCells.Count + numericOverrides.Count;
            if (total > MaxCellOverrideCount)
            {
                throw new InvalidSubstitutionValueException(
                    "cellOverrides",
                    $"セル番地の直接指定が文字列・数値の合計で {total} 件あり、上限({MaxCellOverrideCount} 件)を超えています。",
                    report.Definition.ReportCode,
                    report.Sheet.Name);
            }

            var asText = new Dictionary<string, string>(numericOverrides.Count, StringComparer.Ordinal);
            foreach (var (addressText, value) in numericOverrides)
            {
                if (double.IsNaN(value) || double.IsInfinity(value))
                {
                    throw new InvalidSubstitutionValueException(
                        addressText ?? string.Empty,
                        $"セル番地 '{addressText}' に指定した数値 {value.ToString(CultureInfo.InvariantCulture)} はセルに表示できません。",
                        report.Definition.ReportCode,
                        report.Sheet.Name,
                        CellAddress.TryParse(addressText, out var invalidAt) ? invalidAt : null);
                }

                asText[addressText!] = value.ToString("R", CultureInfo.InvariantCulture);
            }

            // 番地の検証(A1形式・同じセルの重複・結合範囲の内側)、差し込みセルの記録、overflow 指定の解除は
            // 文字列の直接指定と同じ(要件14.2)。そのうえで、置き換えたセルを数値のセルにする。
            var overridden = ApplyCellOverrides(report, asText);

            var sheet = overridden.Sheet;
            var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);
            foreach (var (addressText, value) in numericOverrides)
            {
                var address = CellAddress.Parse(addressText);
                var cell = cells[address];
                var format = cell.Style.NumberFormat;
                var formatted = Parsing.NumberFormatting.FormatWithFallback(value, format);
                cells[address] = cell.WithNumber(asText[addressText], formatted.Text, formatted.Color, formatted.Fallback);
            }

            return overridden with { Sheet = sheet with { Cells = cells } };
        }

        /// <inheritdoc />
        public ReportModel ApplyCellOverrides(ReportModel report, IReadOnlyDictionary<string, string> cellOverrides)
        {
            if (report is null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            if (cellOverrides is null)
            {
                throw new ArgumentNullException(nameof(cellOverrides));
            }

            if (cellOverrides.Count == 0)
            {
                return report;
            }

            var definition = report.Definition;
            var sheet = report.Sheet;

            if (cellOverrides.Count > MaxCellOverrideCount)
            {
                throw new InvalidSubstitutionValueException(
                    "cellOverrides",
                    $"セル番地の直接指定が {cellOverrides.Count} 件あり、上限({MaxCellOverrideCount} 件)を超えています。",
                    definition.ReportCode,
                    sheet.Name);
            }

            // まず全入力を検証してから一括で書き換える(要件2.8, 2.9)。
            var parsed = ParseAndValidateAddresses(definition, sheet, cellOverrides);

            var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);
            var substitutedCells = new HashSet<CellAddress>(report.SubstitutedCells);
            var overriddenCells = new HashSet<CellAddress>(report.OverriddenCells);

            foreach (var (address, replacement) in parsed)
            {
                overriddenCells.Add(address);
                var existing = sheet.GetCell(address);
                cells[address] = existing is null
                    ? new CellModel(replacement, CellValueKind.Text, sheet.GetEffectiveStyle(address), replacement)
                    : existing.WithText(replacement);
                MarkSubstituted(substitutedCells, address, replacement);
            }

            var updatedSheet = sheet with { Cells = cells };
            return report with { Sheet = updatedSheet, SubstitutedCells = substitutedCells, OverriddenCells = overriddenCells };
        }

        /// <summary>
        /// 差し込み済みセルの集合を更新する。空文字は「テンプレートの値を消す」指定であり、
        /// 何も出力されないことが意図どおりのため、以降の欠落検出(要件2.13, 2.14)の対象から外す。
        /// </summary>
        private static void MarkSubstituted(HashSet<CellAddress> substitutedCells, CellAddress address, string replacement)
        {
            if (replacement.Length == 0)
            {
                substitutedCells.Remove(address);
                return;
            }

            substitutedCells.Add(address);
        }

        /// <summary>
        /// 書き込む値を検証し、改行コードを LF に統一して返す(要件2.10, 2.11)。
        /// </summary>
        /// <remarks>
        /// 改行(LF/CR)以外の制御文字(タブを含む)・不可視の書式文字(Unicodeの Cf)はフォントに字形が無く、PDF上で豆腐や空白になる。
        /// 対になっていないサロゲートも同様に文字として描画できない。いずれも呼び出し元の
        /// データ不備であるため、描画を試みずにエラーとする。
        /// </remarks>
        private static string NormalizeValue(
            string? value,
            string description,
            string target,
            ReportDefinition definition,
            string sheetName,
            CellAddress? cell)
        {
            if (value is null)
            {
                throw new InvalidSubstitutionValueException(
                    target,
                    $"{description} の値が null です。空欄にする場合は空文字を渡してください。",
                    definition.ReportCode,
                    sheetName,
                    cell);
            }

            if (value.Length > MaxValueLength)
            {
                throw new InvalidSubstitutionValueException(
                    target,
                    $"{description} の値が {value.Length} 文字あり、上限({MaxValueLength} 文字。Excelのセルに入れられる文字数)を超えています。",
                    definition.ReportCode,
                    sheetName,
                    cell);
            }

            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (c == '\n' || c == '\r' || c == '\u2028' || c == '\u2029')
                {
                    continue;
                }

                if (char.IsHighSurrogate(c) && i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    i++;
                    continue;
                }

                string? problem = null;
                if (char.IsControl(c))
                {
                    problem = $"制御文字 U+{(int)c:X4}";
                }
                else if (char.IsSurrogate(c))
                {
                    problem = $"対になっていないサロゲート U+{(int)c:X4}";
                }
                else if (char.GetUnicodeCategory(c) == UnicodeCategory.Format)
                {
                    // ゼロ幅スペース(U+200B)・BOM(U+FEFF)など、Webや他システムからの貼り付けで
                    // 混入しやすい不可視の書式文字。字形を持たず豆腐として描画される。
                    problem = $"不可視の書式文字 U+{(int)c:X4}";
                }

                if (problem is not null)
                {
                    throw new InvalidSubstitutionValueException(
                        target,
                        $"{description} の値の {i + 1} 文字目に{problem}が含まれています。"
                            + "PDFに文字として描画できないため、呼び出し元で除去または置き換えてください。",
                        definition.ReportCode,
                        sheetName,
                        cell);
                }
            }

            // CRLF・CR単独、およびUnicodeの行区切り(U+2028)・段落区切り(U+2029)をLFに統一する。
            return value.IndexOf('\r') < 0 && value.IndexOf('\u2028') < 0 && value.IndexOf('\u2029') < 0
                ? value
                : value.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\u2028', '\n').Replace('\u2029', '\n');
        }

        /// <summary>
        /// セル番地文字列をパースし、結合セル範囲内の非アンカー位置を指定していないかを検証する(要件2.8, 2.9)。
        /// </summary>
        private static List<(CellAddress Address, string Replacement)> ParseAndValidateAddresses(
            ReportDefinition definition, SheetModel sheet, IReadOnlyDictionary<string, string> cellOverrides)
        {
            var parsed = new List<(CellAddress, string)>(cellOverrides.Count);
            var addressTextByCell = new Dictionary<CellAddress, string>(cellOverrides.Count);

            foreach (var (addressText, replacement) in cellOverrides)
            {
                if (!CellAddress.TryParse(addressText, out var address))
                {
                    throw new InvalidCellOverrideAddressException(
                        addressText,
                        $"セル番地 '{addressText}' はA1形式として解釈できません。",
                        definition.ReportCode,
                        sheet.Name);
                }

                // "B1"・"b1"・"$B$1" は辞書のキーとしては別物だが同じセルを指す。
                // そのまま適用すると列挙順で後の値だけが黙って採用されるため、正規化した番地で重複を検出する。
                if (addressTextByCell.TryGetValue(address, out var otherText))
                {
                    throw new InvalidCellOverrideAddressException(
                        addressText,
                        $"セル番地 '{addressText}' と '{otherText}' が同じセル {address} を指しています。"
                            + "1つのセルへの直接指定は1つにしてください。",
                        definition.ReportCode,
                        sheet.Name);
                }

                addressTextByCell.Add(address, addressText);

                if (sheet.IsNonAnchorMergedCell(address, out var merged))
                {
                    throw new NonAnchorMergedCellOverrideException(
                        address,
                        merged!.Anchor,
                        $"セル '{address}' は結合セル範囲の先頭(アンカー: '{merged.Anchor}')ではないため、"
                            + "直接指定して上書きすることはできません。アンカーのセル番地を指定してください。",
                        definition.ReportCode,
                        sheet.Name);
                }

                var normalized = NormalizeValue(
                    replacement, $"セル {address} の直接指定", addressText, definition, sheet.Name, address);

                parsed.Add((address, normalized));
            }

            return parsed;
        }
    }
}
