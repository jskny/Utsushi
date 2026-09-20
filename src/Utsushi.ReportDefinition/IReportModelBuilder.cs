using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.ReportDefinitions
{
    /// <summary>
    /// <see cref="WorkbookModel"/> と帳票定義を突合して <see cref="ReportModel"/> を構築する。
    /// </summary>
    public interface IReportModelBuilder
    {
        /// <summary>
        /// 帳票定義が期待するシート構成・セル番地が実在するかを検証し、<see cref="ReportModel"/> を構築する。
        /// </summary>
        /// <exception cref="ReportStructureMismatchException">
        /// シート名やセル番地が一致しない場合(要件1.4)。
        /// </exception>
        ReportModel Build(WorkbookModel workbook, ReportDefinition definition);
    }

    /// <inheritdoc />
    public sealed class ReportModelBuilder : IReportModelBuilder
    {
        /// <inheritdoc />
        public ReportModel Build(WorkbookModel workbook, ReportDefinition definition)
        {
            if (workbook is null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            if (definition is null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            var sheet = workbook.FindSheet(definition.SheetName);
            if (sheet is null)
            {
                var available = string.Join(", ", workbook.Sheets.Select(s => "'" + s.Name + "'"));
                throw new ReportStructureMismatchException(
                    $"帳票定義が期待するシート '{definition.SheetName}' がブックに存在しません。"
                    + $"ブック内のシート: {(available.Length == 0 ? "(なし)" : available)}",
                    definition.ReportCode,
                    definition.SheetName);
            }

            ValidateSubstitutionCells(sheet, definition);
            ValidatePrintArea(sheet, definition);

            return ReportModel.Create(definition, sheet, workbook.DefaultFont);
        }

        /// <summary>
        /// 置換対象セルがシートの使用範囲内に存在するかを検証する(要件1.4)。
        /// </summary>
        /// <remarks>
        /// 値が空のセルは <c>Cells</c> に現れないことがあるため、セルの存在ではなく
        /// 「使用範囲に含まれるか」「結合セルの左上か」を判定基準とする。
        /// </remarks>
        private static void ValidateSubstitutionCells(SheetModel sheet, ReportDefinition definition)
        {
            var usedRange = sheet.GetUsedRange();

            foreach (var field in definition.SubstitutionFields)
            {
                if (usedRange is null || !usedRange.Value.Contains(field.Cell))
                {
                    throw new ReportStructureMismatchException(
                        $"置換キー '{field.Key}' の対象セル {field.Cell} が、"
                        + $"シート '{sheet.Name}' の使用範囲{(usedRange is null ? "(空)" : " " + usedRange.Value)}の外にあります。",
                        definition.ReportCode,
                        sheet.Name,
                        field.Cell);
                }

                var merged = sheet.FindMergedRange(field.Cell);
                if (merged is not null && merged.Anchor != field.Cell)
                {
                    throw new ReportStructureMismatchException(
                        $"置換キー '{field.Key}' の対象セル {field.Cell} は結合範囲 {merged.Range} の内側にあります。"
                        + $"結合範囲の左上セル {merged.Anchor} を指定してください。",
                        definition.ReportCode,
                        sheet.Name,
                        field.Cell);
                }
            }
        }

        /// <summary>
        /// 帳票定義の印刷範囲上書きが使用範囲と矛盾しないことを確認する。
        /// </summary>
        private static void ValidatePrintArea(SheetModel sheet, ReportDefinition definition)
        {
            if (definition.PrintAreaOverride is not { } printArea)
            {
                return;
            }

            var usedRange = sheet.GetUsedRange();
            if (usedRange is null)
            {
                throw new ReportStructureMismatchException(
                    $"帳票定義が印刷範囲 {printArea} を指定していますが、シート '{sheet.Name}' にセルがありません。",
                    definition.ReportCode,
                    sheet.Name);
            }
        }
    }
}
