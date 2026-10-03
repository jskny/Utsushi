using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Rendering;

namespace Utsushi
{
    /// <summary>
    /// Excel ファイルのセルに値を書き込み、PDF として保存する簡易API(要件14)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="ReportPdfConverter"/> の上に置いた薄い入口で、変換の処理・検証・例外は同じものを使う。
    /// </para>
    /// <code>
    /// using var pdf = new Utsushi.Excel2Pdf("input.xlsx");
    /// pdf.SetText("C2", "Hello World");
    /// pdf.SetValue("D5", 123.5);
    /// pdf.Save("output.pdf");
    /// </code>
    /// <para>
    /// 帳票コードを指定しないコンストラクタは帳票定義なしの変換(要件12。アクティブシートを変換する)、
    /// 帳票コードと帳票定義のルートを指定するコンストラクタは帳票定義ありの変換(要件1〜11)を使う。
    /// </para>
    /// <para>
    /// コンストラクタはファイルを開かない。<see cref="Save(string)"/> のたびに入力ファイルを読み直すため、
    /// 値を変えて何度でも保存できる。複数スレッドから同じインスタンスを同時に使わないこと。
    /// </para>
    /// </remarks>
    public sealed class Excel2Pdf : IDisposable
    {
        /// <summary>Excel の1900年日付システムで表せる最も古い日付。</summary>
        private static readonly DateTime MinimumDate = new(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

        private readonly ReportPdfConverter _converter;
        private readonly bool _ownsConverter;
        private readonly Excel2PdfOptions _options;

        /// <summary>正規化したセル番地 → 設定した値(<see cref="string"/> または <see cref="double"/>)。</summary>
        private readonly Dictionary<CellAddress, object> _cells = new();

        /// <summary>置換キー → 値(帳票定義ありのときだけ。<see cref="SetField"/>)。</summary>
        private readonly Dictionary<string, string> _fields = new(StringComparer.Ordinal);

        private bool _disposed;

        /// <summary>帳票定義を使わずに変換する(要件12)。ブックのアクティブシートを変換する。</summary>
        /// <param name="xlsxPath">入力の Excel ファイルのパス。</param>
        /// <param name="options">フォント・PDF出力などの設定。null の場合は既定(フォントは厳格モード)。</param>
        public Excel2Pdf(string xlsxPath, Excel2PdfOptions? options = null)
            : this(xlsxPath, reportCode: null, options, converter: null, reportDefinitionRoot: null)
        {
        }

        /// <summary>登録済みの帳票定義を使って変換する(要件1〜11)。</summary>
        /// <param name="xlsxPath">入力の Excel ファイル(テンプレート)のパス。</param>
        /// <param name="reportCode">帳票コード(<c>definition.json</c> の <c>reportCode</c>)。</param>
        /// <param name="reportDefinitionRoot">帳票定義のルートディレクトリ(例: <c>samples/reports</c>)。</param>
        /// <param name="options">フォント・PDF出力などの設定。null の場合は既定(フォントは厳格モード)。</param>
        public Excel2Pdf(string xlsxPath, string reportCode, string reportDefinitionRoot, Excel2PdfOptions? options = null)
            : this(
                xlsxPath,
                reportCode ?? throw new ArgumentNullException(nameof(reportCode)),
                options,
                converter: null,
                reportDefinitionRoot ?? throw new ArgumentNullException(nameof(reportDefinitionRoot)))
        {
        }

        /// <summary>
        /// 呼び出し元が作ったコンバータを使う。大量に発行する場合に、フォントの解決結果を使い回すために使う。
        /// コンバータは <see cref="Dispose"/> で解放しない(要件14.9)。
        /// </summary>
        /// <param name="xlsxPath">入力の Excel ファイルのパス。</param>
        /// <param name="converter">変換に使うコンバータ。</param>
        /// <param name="reportCode">
        /// 帳票コード。null の場合は帳票定義なしの変換。帳票定義ありで使う場合は、帳票定義のルートを指定して作ったコンバータ
        /// (<see cref="ReportPdfConverter.CreateDefault(string, FontResolverOptions?, PdfRenderOptions?)"/>)を渡す。
        /// </param>
        /// <param name="options">
        /// 帳票定義なしの変換の文書名・最大数字幅。フォント・PDF出力の設定(<see cref="Excel2PdfOptions.Fonts"/>・
        /// <see cref="Excel2PdfOptions.Render"/>)はコンバータのものを使うため、指定するとエラー。
        /// </param>
        public Excel2Pdf(string xlsxPath, ReportPdfConverter converter, string? reportCode = null, Excel2PdfOptions? options = null)
            : this(
                xlsxPath,
                reportCode,
                options,
                converter ?? throw new ArgumentNullException(nameof(converter)),
                reportDefinitionRoot: null)
        {
        }

        private Excel2Pdf(
            string xlsxPath,
            string? reportCode,
            Excel2PdfOptions? options,
            ReportPdfConverter? converter,
            string? reportDefinitionRoot)
        {
            if (xlsxPath is null)
            {
                throw new ArgumentNullException(nameof(xlsxPath));
            }

            if (string.IsNullOrWhiteSpace(xlsxPath))
            {
                throw new ArgumentException("入力の Excel ファイルのパスが空です。", nameof(xlsxPath));
            }

            _options = options ?? new Excel2PdfOptions();
            ReportPdfConverter.ValidateMaxDigitWidth(_options.MaxDigitWidthPx);
            if (reportCode is not null && _options.MaxDigitWidthPx is not null)
            {
                throw new ArgumentException(
                    "帳票定義ありの変換では、最大数字幅は帳票定義の maxDigitWidthPx を使います。"
                        + $"{nameof(Excel2PdfOptions.MaxDigitWidthPx)} は指定できません。",
                    nameof(options));
            }

            if (reportCode is not null && string.IsNullOrWhiteSpace(reportCode))
            {
                throw new ArgumentException("帳票コードが空です。帳票定義なしで変換する場合は null を渡してください。", nameof(reportCode));
            }

            XlsxPath = xlsxPath;
            ReportCode = reportCode;

            if (converter is not null)
            {
                if (_options.Fonts is not null || _options.Render is not null)
                {
                    throw new ArgumentException(
                        $"コンバータを渡す場合、フォント・PDF出力の設定はコンバータのものを使います。"
                            + $"{nameof(Excel2PdfOptions.Fonts)}・{nameof(Excel2PdfOptions.Render)} は指定できません。",
                        nameof(options));
                }

                _converter = converter;
                _ownsConverter = false;
            }
            else
            {
                _converter = reportDefinitionRoot is null
                    ? ReportPdfConverter.CreateDefault(_options.Fonts, _options.Render)
                    : ReportPdfConverter.CreateDefault(reportDefinitionRoot, _options.Fonts, _options.Render);
                _ownsConverter = true;
            }
        }

        /// <summary>入力の Excel ファイルのパス。</summary>
        public string XlsxPath { get; }

        /// <summary>帳票コード。帳票定義なしの変換では null。</summary>
        public string? ReportCode { get; }

        /// <summary>
        /// セルに文字列を設定する(要件14.2)。同じセルに何度設定しても、最後に設定した値を使う。
        /// セルの数値書式は効かない(数値として表示したい場合は <see cref="SetValue(string, double)"/>)。
        /// </summary>
        /// <param name="cell">セル番地(A1形式。例: <c>"C2"</c>)。結合セルは範囲の左上のセル。</param>
        /// <param name="text">文字列。改行(LF/CRLF)は、折り返し表示のセルでは改行、それ以外では1行につながる。</param>
        /// <returns>このインスタンス(続けて設定できる)。</returns>
        /// <exception cref="InvalidCellOverrideAddressException">セル番地をA1形式として解釈できない場合。</exception>
        /// <remarks>文字列の内容の検証(制御文字・長さなど。要件2.10, 2.16)は保存の時点で行う。</remarks>
        public Excel2Pdf SetText(string cell, string text)
        {
            ThrowIfDisposed();
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            _cells[ParseCell(cell)] = text;
            return this;
        }

        /// <summary>
        /// セルに数値を設定する(要件14.3)。セルの数値書式(通貨・桁区切り・日付など)と色の指定で表示される。
        /// </summary>
        /// <param name="cell">セル番地(A1形式)。</param>
        /// <param name="value">数値。NaN・無限大は不可。</param>
        /// <returns>このインスタンス。</returns>
        /// <exception cref="InvalidCellOverrideAddressException">セル番地をA1形式として解釈できない場合。</exception>
        /// <exception cref="InvalidSubstitutionValueException">数値が NaN・無限大の場合(要件14.4)。</exception>
        public Excel2Pdf SetValue(string cell, double value)
        {
            ThrowIfDisposed();
            var address = ParseCell(cell);
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new InvalidSubstitutionValueException(
                    cell,
                    $"セル {address} に設定した数値 {value.ToString(CultureInfo.InvariantCulture)} はセルに表示できません。",
                    DocumentLabel,
                    sheetName: null,
                    address);
            }

            _cells[address] = value;
            return this;
        }

        /// <summary>
        /// セルに数値(金額など)を設定する。<see cref="SetValue(string, double)"/> と同じ。Excel のセルの値と同じく
        /// <see cref="double"/> で扱うため、有効数字が15桁を超える部分は丸められる。
        /// </summary>
        public Excel2Pdf SetValue(string cell, decimal value) => SetValue(cell, (double)value);

        /// <summary>
        /// セルに整数を設定する。<see cref="SetValue(string, double)"/> と同じ。<see cref="double"/> で扱うため、
        /// 絶対値が 2^53 を超える整数は丸められる。
        /// </summary>
        public Excel2Pdf SetValue(string cell, long value) => SetValue(cell, (double)value);

        /// <summary>
        /// セルに時間(経過時間・時刻)を設定する。日数をシリアル値として設定し、セルの時刻の書式(<c>h:mm</c>・<c>[h]:mm</c> など)で表示される。
        /// </summary>
        /// <param name="cell">セル番地(A1形式)。</param>
        /// <param name="value">時間。負の値は不可。</param>
        /// <exception cref="InvalidSubstitutionValueException">負の時間の場合(Excel は負の時刻を表示できない)。</exception>
        public Excel2Pdf SetValue(string cell, TimeSpan value)
        {
            ThrowIfDisposed();
            var address = ParseCell(cell);
            if (value < TimeSpan.Zero)
            {
                throw new InvalidSubstitutionValueException(
                    cell,
                    $"セル {address} に設定した時間 {value} は負のため、Excel の時刻として表せません。",
                    DocumentLabel,
                    sheetName: null,
                    address);
            }

            _cells[address] = value.TotalDays;
            return this;
        }

        /// <summary>
        /// セルに日時を設定する(要件14.3)。Excel のシリアル値(1900年日付システム)として設定し、セルの日付・時刻の書式で表示される。
        /// </summary>
        /// <remarks>
        /// 表示形式が「標準」のセルではシリアル値(例: 46113)のまま表示される(Excel で日付を入力したときのような表示形式の
        /// 自動設定は行わない)。テンプレートのセルに日付の表示形式を設定しておくこと。<see cref="DateTime.Kind"/> は考慮せず、
        /// 年月日・時刻をそのまま使う。時刻だけを設定する場合は <see cref="SetValue(string, TimeSpan)"/> を使う。
        /// </remarks>
        /// <param name="cell">セル番地(A1形式)。</param>
        /// <param name="value">日時。1900年1月1日より前は不可。</param>
        /// <exception cref="InvalidSubstitutionValueException">1900年1月1日より前の日時の場合(要件14.4)。</exception>
        public Excel2Pdf SetValue(string cell, DateTime value)
        {
            ThrowIfDisposed();
            var address = ParseCell(cell);
            if (value < MinimumDate)
            {
                throw new InvalidSubstitutionValueException(
                    cell,
                    $"セル {address} に設定した日時 {value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} は、"
                        + "Excel の日付(1900年1月1日以降)として表せません。",
                    DocumentLabel,
                    sheetName: null,
                    address);
            }

            _cells[address] = ToExcelSerial(value);
            return this;
        }

        /// <summary>
        /// 置換キーに値を設定する(帳票定義ありのときだけ。要件14.7)。置換キーによる置換(要件2.1〜2.6)として扱い、
        /// 帳票定義の必須キー・はみ出し時の挙動(<c>overflow</c>)が効く。
        /// </summary>
        /// <param name="key">置換キー(<c>substitutionFields[].key</c>)。</param>
        /// <param name="text">文字列。</param>
        /// <exception cref="InvalidOperationException">帳票定義なしで作った場合。</exception>
        public Excel2Pdf SetField(string key, string text)
        {
            ThrowIfDisposed();
            if (key is null)
            {
                throw new ArgumentNullException(nameof(key));
            }

            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            if (ReportCode is null)
            {
                throw new InvalidOperationException(
                    $"置換キーは帳票定義ありの変換でだけ使えます。帳票定義なしでは {nameof(SetText)} でセル番地を指定してください。");
            }

            _fields[key] = text;
            return this;
        }

        /// <summary>
        /// PDF をファイルへ保存する(要件14.8)。失敗時に不完全なファイルを残さない。保存のたびに入力の Excel ファイルを読み直す。
        /// </summary>
        /// <param name="pdfPath">保存先のパス。</param>
        /// <exception cref="UtsushiException">入力ファイル・帳票定義・設定した値・レイアウト・描画・書き込みのいずれかで失敗した場合。</exception>
        /// <exception cref="NotSupportedException">
        /// 渡したコンバータの <see cref="Substitution.ICellSubstitutor"/> が数値の直接指定に対応しておらず、<see cref="SetValue(string, double)"/>
        /// などで数値を設定した場合(既定のコンバータでは起きない)。
        /// </exception>
        public void Save(string pdfPath)
        {
            ThrowIfDisposed();
            ReportPdfConverter.ValidateOutputPath(pdfPath);

            var layout = ComputeLayout(checkFit: false);
            _converter.RenderToFile(layout, pdfPath);
        }

        /// <summary>PDF をストリームへ書き出す。</summary>
        /// <param name="output">出力先。</param>
        public void Save(Stream output)
        {
            ThrowIfDisposed();
            if (output is null)
            {
                throw new ArgumentNullException(nameof(output));
            }

            var layout = ComputeLayout(checkFit: false);
            _converter.RenderToStream(layout, output);
        }

        /// <summary>
        /// PDF を出力せずに、文字がセルに収まらない箇所を返す(要件13)。設定した値で確認する。
        /// </summary>
        public FitCheckResult CheckFit()
        {
            ThrowIfDisposed();
            return ReportPdfConverter.ToFitCheckResult(ComputeLayout(checkFit: true));
        }

        /// <summary>自身が作ったコンバータ(フォント解決のリソース)を解放する。渡されたコンバータは解放しない。</summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_ownsConverter)
            {
                _converter.Dispose();
            }
        }

        /// <summary>例外・PDFのタイトルに使う名前(帳票コード、無ければ文書名)。</summary>
        private string DocumentLabel => ReportCode ?? _options.DocumentName ?? Path.GetFileNameWithoutExtension(XlsxPath);

        internal PagedLayout ComputeLayout(bool checkFit)
        {
            var texts = new Dictionary<string, string>();
            var numbers = new Dictionary<string, double>();
            foreach (var (address, value) in _cells)
            {
                if (value is double number)
                {
                    numbers[address.ToString()] = number;
                }
                else
                {
                    texts[address.ToString()] = (string)value;
                }
            }

            using var input = ReportPdfConverter.OpenInputFile(XlsxPath, DocumentLabel);
            return ReportCode is not null
                ? _converter.ComputeLayoutCore(ReportCode, input, _fields, texts, numbers, checkFit)
                : _converter.ComputeLayoutWithoutDefinitionCore(
                    input, texts, DocumentLabel, _options.MaxDigitWidthPx, numbers, checkFit);
        }

        private CellAddress ParseCell(string cell)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            if (!CellAddress.TryParse(cell, out var address))
            {
                throw new InvalidCellOverrideAddressException(
                    cell, $"セル番地 '{cell}' はA1形式として解釈できません。", DocumentLabel);
            }

            return address;
        }

        /// <summary>
        /// 日時を Excel のシリアル値(1900年日付システム)にする。Excel は1900年を閏年として扱う(存在しない1900年2月29日が60)ため、
        /// 1900年3月1日より前は OLE オートメーション日付より1小さい。
        /// </summary>
        internal static double ToExcelSerial(DateTime value)
        {
            var serial = value.ToOADate();
            return serial < 61 ? serial - 1 : serial;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(Excel2Pdf));
            }
        }
    }

    /// <summary><see cref="Excel2Pdf"/> の設定。</summary>
    public sealed class Excel2PdfOptions
    {
        /// <summary>フォント解決の設定。null の場合は厳格モード(<see cref="FontResolverOptions.Strict"/>)。</summary>
        public FontResolverOptions? Fonts { get; init; }

        /// <summary>PDF出力の設定。null の場合は既定(<see cref="PdfRenderOptions.Default"/>)。</summary>
        public PdfRenderOptions? Render { get; init; }

        /// <summary>
        /// 帳票定義なしの変換で、PDFのタイトル・ヘッダー/フッターのファイル名・エラー情報に使う名前(要件12.5)。
        /// null の場合は入力ファイル名から拡張子を除いたもの。
        /// </summary>
        public string? DocumentName { get; init; }

        /// <summary>帳票定義なしの変換で、列幅の換算に使う最大数字幅(ピクセル)。null の場合はブックの標準フォントから見積もる。</summary>
        public double? MaxDigitWidthPx { get; init; }
    }
}
