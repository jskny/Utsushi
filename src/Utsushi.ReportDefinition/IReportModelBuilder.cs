using System;
using System.Linq;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.ReportDefinitions
{
    /// <summary>
    /// <see cref="WorkbookModel"/> と変換の設定から <see cref="ReportModel"/> を構築する。
    /// </summary>
    public interface IReportModelBuilder
    {
        /// <summary>変換の設定が指すシートを取り出し、<see cref="ReportModel"/> を構築する。</summary>
        /// <exception cref="InvalidOperationException">
        /// 設定のシートがブックに無い場合(読み込みは対象のシートだけを返すため、通常は起きない)。
        /// </exception>
        ReportModel Build(WorkbookModel workbook, ConversionSettings definition);
    }

    /// <inheritdoc />
    public sealed class ReportModelBuilder : IReportModelBuilder
    {
        /// <inheritdoc />
        public ReportModel Build(WorkbookModel workbook, ConversionSettings definition)
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
                // 読み込み(Parsing)は変換対象のシートだけを返し、無ければ InvalidExcelFileException にする。
                // ここに来るのは、契約に反する IWorkbookReader が差し替えられた場合だけである。
                var available = string.Join(", ", workbook.Sheets.Select(s => "'" + s.Name + "'"));
                throw new InvalidOperationException(
                    $"変換するシート '{definition.SheetName}' がブックに存在しません。"
                    + $"ブック内のシート: {(available.Length == 0 ? "(なし)" : available)}");
            }

            return ReportModel.Create(definition, sheet, workbook.DefaultFont);
        }
    }
}
