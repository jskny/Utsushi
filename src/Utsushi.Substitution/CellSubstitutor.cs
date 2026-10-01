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
    /// 帳票定義に基づくセル値置換の既定実装。
    /// </summary>
    /// <remarks>
    /// <para>書式(フォント・配置・罫線・数値書式)には一切触れない(要件2.2)。</para>
    /// <para>
    /// はみ出し時の挙動(shrink/clip/wrap/overflow)はここでは計算せず、
    /// 帳票定義の指定値を <see cref="ReportModel.OverflowByCell"/> に載せて
    /// Layout レイヤーへ引き渡すのみとする(要件2.5、design.md「Substitution レイヤー」)。
    /// </para>
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
        public ReportModel Apply(ReportModel report, IReadOnlyDictionary<string, string> values)
        {
            if (report is null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            if (values is null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            var definition = report.Definition;
            var sheet = report.Sheet;

            ValidateKnownKeys(definition, sheet.Name, values);
            var normalizedValues = NormalizeFieldValues(definition, sheet.Name, values);
            ValidateRequiredValues(definition, sheet.Name, normalizedValues);

            var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);
            var overflowByCell = new Dictionary<CellAddress, OverflowBehavior>(report.OverflowByCell);
            var substitutedCells = new HashSet<CellAddress>(report.SubstitutedCells);

            foreach (var field in definition.SubstitutionFields)
            {
                // 任意フィールドで値が渡されていない場合は、テンプレートの既存値をそのまま残す。
                if (!normalizedValues.TryGetValue(field.Key, out var replacement))
                {
                    continue;
                }

                var existing = sheet.GetCell(field.Cell);
                cells[field.Cell] = existing is null
                    ? new CellModel(replacement, CellValueKind.Text, CellStyle.Default, replacement)
                    : existing.WithText(replacement);
                MarkSubstituted(substitutedCells, field.Cell, replacement);

                // 帳票定義がはみ出し挙動を明示している場合だけ Layout レイヤーへ伝える(要件2.5)。
                // 未指定(null)のときは記録せず、Excel 側のセル書式を Layout がそのまま使う。
                if (field.Overflow is { } overflow)
                {
                    overflowByCell[field.Cell] = overflow;
                }
            }

            var updatedSheet = sheet with { Cells = cells };
            return report with { Sheet = updatedSheet, OverflowByCell = overflowByCell, SubstitutedCells = substitutedCells };
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

            // Apply と同様、まず全入力を検証してから一括で書き換える(要件2.8, 2.9)。
            var parsed = ParseAndValidateAddresses(definition, sheet, cellOverrides);

            var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);
            var overflowByCell = new Dictionary<CellAddress, OverflowBehavior>(report.OverflowByCell);
            var substitutedCells = new HashSet<CellAddress>(report.SubstitutedCells);

            foreach (var (address, replacement) in parsed)
            {
                var existing = sheet.GetCell(address);
                cells[address] = existing is null
                    ? new CellModel(replacement, CellValueKind.Text, CellStyle.Default, replacement)
                    : existing.WithText(replacement);
                MarkSubstituted(substitutedCells, address, replacement);

                // この経路は帳票定義のoverflow指定を経由しないため、名前付きキー方式(Apply)が
                // 同じセルに残した指定があれば取り除き、常にExcel側のセル書式に従わせる(design.md)。
                overflowByCell.Remove(address);
            }

            var updatedSheet = sheet with { Cells = cells };
            return report with { Sheet = updatedSheet, OverflowByCell = overflowByCell, SubstitutedCells = substitutedCells };
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

        /// <summary>置換キー経由の値をすべて検証・正規化する(要件2.10, 2.11)。</summary>
        private static Dictionary<string, string> NormalizeFieldValues(
            ReportDefinition definition, string sheetName, IReadOnlyDictionary<string, string> values)
        {
            var normalized = new Dictionary<string, string>(values.Count, StringComparer.Ordinal);

            foreach (var (key, value) in values)
            {
                CellAddress? cell = definition.TryGetField(key, out var field) ? field.Cell : null;
                normalized[key] = NormalizeValue(value, $"置換キー '{key}'", key, definition, sheetName, cell);
            }

            return normalized;
        }

        /// <summary>
        /// 置換値を検証し、改行コードを LF に統一して返す(要件2.10, 2.11)。
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

                // 必須キーのセルを直接指定で空にすると、要件2.12の検証をすり抜けてしまう。
                var requiredField = definition.RequiredFields.FirstOrDefault(f => f.Cell == address);
                if (requiredField is not null && string.IsNullOrWhiteSpace(normalized))
                {
                    throw new RequiredSubstitutionValueMissingException(
                        requiredField.Key,
                        $"必須の置換キー '{requiredField.Key}'(セル {address})を、セル番地の直接指定で空にすることはできません。",
                        definition.ReportCode,
                        sheet.Name,
                        address);
                }

                parsed.Add((address, normalized));
            }

            return parsed;
        }

        /// <summary>帳票定義に存在しない置換キーが渡されていないかを検証する(要件2.3)。</summary>
        private static void ValidateKnownKeys(
            ReportDefinition definition, string sheetName, IReadOnlyDictionary<string, string> values)
        {
            foreach (var key in values.Keys)
            {
                if (definition.TryGetField(key, out _))
                {
                    continue;
                }

                throw new SubstitutionKeyNotFoundException(
                    key,
                    $"置換キー '{key}' は帳票 '{definition.ReportCode}' の帳票定義に存在しません。",
                    definition.ReportCode,
                    sheetName);
            }
        }

        /// <summary>
        /// 必須の置換キーに値が渡されているかを検証する(要件2.4)。空文字・空白のみの値も
        /// 「値が無い」ものとして扱う(要件2.12)。
        /// </summary>
        private static void ValidateRequiredValues(
            ReportDefinition definition, string sheetName, IReadOnlyDictionary<string, string> values)
        {
            foreach (var field in definition.RequiredFields)
            {
                var found = values.TryGetValue(field.Key, out var value);
                if (found && !string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                throw new RequiredSubstitutionValueMissingException(
                    field.Key,
                    found
                        ? $"必須の置換キー '{field.Key}'(セル {field.Cell})の値が空です。"
                        : $"必須の置換キー '{field.Key}'(セル {field.Cell})に対応する値が渡されていません。",
                    definition.ReportCode,
                    sheetName,
                    field.Cell);
            }
        }
    }
}
