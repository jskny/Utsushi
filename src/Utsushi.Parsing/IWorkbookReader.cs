using System.IO;
using Utsushi.Parsing.Model;

namespace Utsushi.Parsing;

/// <summary>
/// <c>.xlsx</c> を読み取り、内部モデル <see cref="WorkbookModel"/> を構築する。
/// </summary>
/// <remarks>
/// 実装は OOXML SDK の型を <see cref="WorkbookModel"/> の外に漏らしてはならない(design.md「Parsing レイヤー」)。
/// </remarks>
public interface IWorkbookReader
{
    /// <summary>ストリームからワークブックを読み取る。</summary>
    /// <param name="xlsxStream">読み取り可能な .xlsx のストリーム。</param>
    /// <param name="options">読み取りオプション。null の場合は既定値を使う。</param>
    WorkbookModel Read(Stream xlsxStream, WorkbookReadOptions? options = null);
}

/// <summary>
/// ワークブック読み取りのオプション。帳票定義由来の設定を Parsing レイヤーへ渡すために使う。
/// </summary>
/// <param name="UnsupportedElementBehavior">サポート外要素(図形・グラフ・外部参照)を検出したときの挙動。要件1.5。</param>
/// <param name="ReportCode">エラーに含める帳票コード(判明している場合)。要件6.4。</param>
/// <param name="SheetNameFilter">読み取るシートを限定する場合のシート名。null ならすべてのシートを読む。</param>
public sealed record WorkbookReadOptions(
    UnsupportedElementBehavior UnsupportedElementBehavior = UnsupportedElementBehavior.Ignore,
    string? ReportCode = null,
    string? SheetNameFilter = null)
{
    public static WorkbookReadOptions Default { get; } = new();
}

/// <summary>サポート外要素を検出したときの挙動(要件1.5)。</summary>
public enum UnsupportedElementBehavior
{
    /// <summary>無視して読み取りを続行する。</summary>
    Ignore = 0,

    /// <summary>エラーとして変換を中止する。</summary>
    Error,
}
