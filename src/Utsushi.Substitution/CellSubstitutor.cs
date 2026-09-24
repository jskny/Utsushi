using System;
using System.Collections.Generic;
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
        /// 改行(LF/CR)以外の制御文字(タブを含む)はフォントに字形が無く、PDF上で豆腐や空白になる。
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

            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (c == '\n' || c == '\r')
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

            return value.IndexOf('\r') < 0
                ? value
                : value.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        /// <summary>
        /// セル番地文字列をパースし、結合セル範囲内の非アンカー位置を指定していないかを検証する(要件2.8, 2.9)。
        /// </summary>
        private static List<(CellAddress Address, string Replacement)> ParseAndValidateAddresses(
            ReportDefinition definition, SheetModel sheet, IReadOnlyDictionary<string, string> cellOverrides)
        {
            var parsed = new List<(CellAddress, string)>(cellOverrides.Count);

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

                parsed.Add((address, NormalizeValue(
                    replacement, $"セル {address} の直接指定", addressText, definition, sheet.Name, address)));
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
