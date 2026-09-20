using System.IO;
using Utsushi.Layout.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// <see cref="PagedLayout"/> をPDFとして描画する。
    /// </summary>
    public interface IPdfRenderer
    {
        /// <summary>レイアウト結果をPDFとして <paramref name="output"/> へ書き出す。</summary>
        /// <remarks>
        /// 途中でエラーが発生した場合、<paramref name="output"/> には何も書き込まない(要件5.4)。
        /// </remarks>
        void Render(PagedLayout layout, Stream output);
    }
}
