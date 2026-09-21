using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Utsushi.SampleGenerator
{
/// <summary>
/// 単色矩形のPNGバイト列を生成する。
/// </summary>
/// <remarks>
/// サンプル帳票のロゴ(要件9)は本物の画像素材を持ち込む必要がなく、
/// 単色矩形で十分なプレースホルダになる。外部の画像処理ライブラリ(<c>System.Drawing.Common</c>等)
/// を追加で持ち込まずに済ませるため、PNG(RFC 2083)の最小構成(IHDR/IDAT/IEND、
/// フィルタなしのRGBスキャンライン)を直接組み立てる。
/// </remarks>
internal static class PlaceholderPng
{
    private static readonly byte[] Signature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    private static readonly uint[] CrcTable = BuildCrcTable();

    /// <summary>指定色で塗りつぶした<paramref name="width"/>x<paramref name="height"/>のPNGを作る。</summary>
    public static byte[] CreateSolidColor(int width, int height, byte r, byte g, byte b)
    {
        var scanlineStride = 1 + (width * 3); // 各行の先頭1バイトはフィルタ種別(0=None)
        var raw = new byte[height * scanlineStride];
        for (var y = 0; y < height; y++)
        {
            var rowStart = y * scanlineStride;
            raw[rowStart] = 0;
            for (var x = 0; x < width; x++)
            {
                var pixelStart = rowStart + 1 + (x * 3);
                raw[pixelStart] = r;
                raw[pixelStart + 1] = g;
                raw[pixelStart + 2] = b;
            }
        }

        using var stream = new MemoryStream();
        stream.Write(Signature, 0, Signature.Length);
        WriteChunk(stream, "IHDR", BuildIhdr(width, height));
        WriteChunk(stream, "IDAT", ZlibCompress(raw));
        WriteChunk(stream, "IEND", Array.Empty<byte>());
        return stream.ToArray();
    }

    private static byte[] BuildIhdr(int width, int height)
    {
        var data = new byte[13];
        WriteUInt32BigEndian(data, 0, (uint)width);
        WriteUInt32BigEndian(data, 4, (uint)height);
        data[8] = 8; // ビット深度
        data[9] = 2; // カラータイプ: truecolor(RGB)
        data[10] = 0; // 圧縮方式(deflateのみ)
        data[11] = 0; // フィルタ方式
        data[12] = 0; // インタレースなし
        return data;
    }

    /// <summary>PNGのIDATが要求するzlib形式(RFC 1950: ヘッダ2byte + deflate + Adler-32)へ圧縮する。</summary>
    private static byte[] ZlibCompress(byte[] data)
    {
        using var compressed = new MemoryStream();
        compressed.WriteByte(0x78);
        compressed.WriteByte(0x9C);

        using (var deflate = new DeflateStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
        {
            deflate.Write(data, 0, data.Length);
        }

        var adler = new byte[4];
        WriteUInt32BigEndian(adler, 0, Adler32(data));
        compressed.Write(adler, 0, adler.Length);
        return compressed.ToArray();
    }

    private static uint Adler32(byte[] data)
    {
        const uint Modulo = 65521;
        uint a = 1, b = 0;
        foreach (var value in data)
        {
            a = (a + value) % Modulo;
            b = (b + a) % Modulo;
        }

        return (b << 16) | a;
    }

    private static void WriteChunk(Stream stream, string type, byte[] data)
    {
        var length = new byte[4];
        WriteUInt32BigEndian(length, 0, (uint)data.Length);
        stream.Write(length, 0, length.Length);

        var typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        var crcInput = new byte[typeBytes.Length + data.Length];
        Buffer.BlockCopy(typeBytes, 0, crcInput, 0, typeBytes.Length);
        Buffer.BlockCopy(data, 0, crcInput, typeBytes.Length, data.Length);

        var crc = new byte[4];
        WriteUInt32BigEndian(crc, 0, Crc32(crcInput));
        stream.Write(crc, 0, crc.Length);
    }

    private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }

    private static uint[] BuildCrcTable()
    {
        var table = new uint[256];
        for (uint n = 0; n < 256; n++)
        {
            var c = n;
            for (var k = 0; k < 8; k++)
            {
                c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
            }

            table[n] = c;
        }

        return table;
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var value in data)
        {
            crc = CrcTable[(crc ^ value) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFFu;
    }
}
}
