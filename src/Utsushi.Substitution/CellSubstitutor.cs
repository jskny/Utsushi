using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution;

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
        ValidateRequiredValues(definition, sheet.Name, values);

        var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);
        var overflowByCell = new Dictionary<CellAddress, OverflowBehavior>(report.OverflowByCell);

        foreach (var field in definition.SubstitutionFields)
        {
            // 任意フィールドで値が渡されていない場合は、テンプレートの既存値をそのまま残す。
            if (!values.TryGetValue(field.Key, out var replacement))
            {
                continue;
            }

            var existing = sheet.GetCell(field.Cell);
            cells[field.Cell] = existing is null
                ? new CellModel(replacement, CellValueKind.Text, CellStyle.Default, replacement)
                : existing.WithText(replacement);

            // 帳票定義がはみ出し挙動を明示している場合だけ Layout レイヤーへ伝える(要件2.5)。
            // 未指定(null)のときは記録せず、Excel 側のセル書式を Layout がそのまま使う。
            if (field.Overflow is { } overflow)
            {
                overflowByCell[field.Cell] = overflow;
            }
        }

        var updatedSheet = sheet with { Cells = cells };
        return report with { Sheet = updatedSheet, OverflowByCell = overflowByCell };
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
        var cells = new Dictionary<CellAddress, CellModel>(sheet.Cells);

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

            var existing = sheet.GetCell(address);
            cells[address] = existing is null
                ? new CellModel(replacement, CellValueKind.Text, CellStyle.Default, replacement)
                : existing.WithText(replacement);
        }

        var updatedSheet = sheet with { Cells = cells };
        return report with { Sheet = updatedSheet };
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

    /// <summary>必須の置換キーに値が渡されているかを検証する(要件2.4)。</summary>
    private static void ValidateRequiredValues(
        ReportDefinition definition, string sheetName, IReadOnlyDictionary<string, string> values)
    {
        foreach (var field in definition.RequiredFields)
        {
            if (values.ContainsKey(field.Key))
            {
                continue;
            }

            throw new RequiredSubstitutionValueMissingException(
                field.Key,
                $"必須の置換キー '{field.Key}'(セル {field.Cell})に対応する値が渡されていません。",
                definition.ReportCode,
                sheetName,
                field.Cell);
        }
    }
}
