using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Utsushi.Parsing.OpenXml;

/// <summary>
/// 数値書式1セクション分(例: <c>"¥"#,##0.00</c>)を解釈して数値を整形する。
/// </summary>
/// <remarks>
/// 対応するのは自社帳票で使う範囲、すなわち
/// 桁区切り(<c>,</c>)・小数桁(<c>0</c>/<c>#</c>/<c>?</c>)・百分率(<c>%</c>)・
/// リテラル(<c>"..."</c>、<c>\x</c>)・色指定(<c>[Red]</c> 等、表示上は無視)である。
/// 指数表記・分数表記は対象外で、その場合は <see cref="Format"/> が General 相当の文字列を返す。
/// </remarks>
internal sealed class NumericSection
{
    private readonly List<Token> _tokens;
    private readonly int _integerDigits;
    private readonly int _decimalDigits;
    private readonly bool _useThousandsSeparator;
    private readonly int _percentCount;
    private readonly double _scaleDivisor;
    private readonly bool _supported;

    private NumericSection(
        List<Token> tokens,
        int integerDigits,
        int decimalDigits,
        bool useThousandsSeparator,
        int percentCount,
        double scaleDivisor,
        bool supported)
    {
        _tokens = tokens;
        _integerDigits = integerDigits;
        _decimalDigits = decimalDigits;
        _useThousandsSeparator = useThousandsSeparator;
        _percentCount = percentCount;
        _scaleDivisor = scaleDivisor;
        _supported = supported;
    }

    public static NumericSection Parse(string section)
    {
        var tokens = new List<Token>();
        var integerDigits = 0;
        var decimalDigits = 0;
        var useThousands = false;
        var percentCount = 0;
        var scaleDivisor = 1.0;
        var supported = true;

        var seenDecimalPoint = false;
        var placeholderSeen = false;
        var trailingCommas = 0;

        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];

            switch (c)
            {
                case '"':
                {
                    var end = section.IndexOf('"', i + 1);
                    var literal = end < 0 ? section.Substring(i + 1) : section.Substring(i + 1, end - i - 1);
                    tokens.Add(Token.Literal(literal));
                    i = end < 0 ? section.Length : end;
                    break;
                }

                case '\\' when i + 1 < section.Length:
                    tokens.Add(Token.Literal(section[i + 1].ToString()));
                    i++;
                    break;

                case '[':
                {
                    // [Red] や [$¥-411] 等。色・ロケール指定は表示に反映しない。
                    var end = section.IndexOf(']', i);
                    var content = end < 0 ? string.Empty : section.Substring(i + 1, end - i - 1);
                    if (content.StartsWith("$", StringComparison.Ordinal))
                    {
                        // [$記号-ロケールID] 形式。記号部分のみリテラルとして出力する。
                        var symbol = content.Substring(1);
                        var dash = symbol.IndexOf('-');
                        if (dash >= 0)
                        {
                            symbol = symbol.Substring(0, dash);
                        }

                        if (symbol.Length > 0)
                        {
                            tokens.Add(Token.Literal(symbol));
                        }
                    }

                    i = end < 0 ? section.Length : end;
                    break;
                }

                case '_':
                    // _x は「文字xの幅の空白」。等幅近似として空白1つに置き換える。
                    tokens.Add(Token.Literal(" "));
                    i++;
                    break;

                case '*':
                    // *x は「xで埋める」。帳票では桁揃え目的のため出力しない。
                    i++;
                    break;

                case '0':
                case '#':
                case '?':
                    placeholderSeen = true;
                    trailingCommas = 0;
                    if (seenDecimalPoint)
                    {
                        decimalDigits++;
                        if (c == '0')
                        {
                            tokens.Add(Token.RequiredDecimalDigit());
                        }
                    }
                    else
                    {
                        if (c == '0')
                        {
                            integerDigits++;
                        }

                        tokens.Add(Token.IntegerPlaceholder());
                    }

                    break;

                case '.':
                    seenDecimalPoint = true;
                    tokens.Add(Token.DecimalPoint());
                    break;

                case ',':
                    if (placeholderSeen && !seenDecimalPoint)
                    {
                        // 数字プレースホルダの直後の "," は桁区切り。
                        // 末尾に続く "," は1000で割る指示(例: #,##0, は千単位表示)。
                        useThousands = true;
                        trailingCommas++;
                    }
                    else
                    {
                        tokens.Add(Token.Literal(","));
                    }

                    break;

                case '%':
                    percentCount++;
                    tokens.Add(Token.Literal("%"));
                    break;

                case 'E':
                case 'e':
                    // 指数表記は対象外。
                    supported = false;
                    break;

                case '/':
                    // 分数表記は対象外。
                    supported = false;
                    break;

                default:
                    tokens.Add(Token.Literal(c.ToString()));
                    break;
            }
        }

        // 末尾の桁区切りカンマは千単位スケーリングを意味する。
        for (var i = 0; i < trailingCommas; i++)
        {
            scaleDivisor *= 1000.0;
        }

        if (trailingCommas > 0)
        {
            // 末尾のカンマは区切り記号としては使わない。
            useThousands = section.Contains("#,#") || section.Contains("0,0") || section.Contains("#,0");
        }

        if (!placeholderSeen)
        {
            supported = false;
        }

        return new NumericSection(
            tokens, integerDigits, decimalDigits, useThousands, percentCount, scaleDivisor, supported);
    }

    public string Format(double value)
    {
        if (!_supported)
        {
            return value.ToString("0.##########", CultureInfo.InvariantCulture);
        }

        var scaled = value;
        for (var i = 0; i < _percentCount; i++)
        {
            scaled *= 100.0;
        }

        scaled /= _scaleDivisor;

        var numberText = BuildNumberText(scaled);
        var sb = new StringBuilder();
        var emitted = false;

        foreach (var token in _tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.Literal:
                    sb.Append(token.Text);
                    break;

                case TokenKind.IntegerPlaceholder:
                case TokenKind.DecimalPoint:
                case TokenKind.RequiredDecimalDigit:
                    if (!emitted)
                    {
                        sb.Append(numberText);
                        emitted = true;
                    }

                    break;
            }
        }

        if (!emitted)
        {
            sb.Append(numberText);
        }

        return sb.ToString();
    }

    private string BuildNumberText(double value)
    {
        var pattern = new StringBuilder();
        pattern.Append(_useThousandsSeparator ? "#,##" : string.Empty);
        pattern.Append(_integerDigits > 0 ? new string('0', _integerDigits) : "0");

        if (_decimalDigits > 0)
        {
            pattern.Append('.').Append(new string('0', _decimalDigits));
        }

        var text = value.ToString(pattern.ToString(), CultureInfo.InvariantCulture);

        // 整数部プレースホルダが 1 つも無い書式(例: ".00")では先頭の 0 を落とす。
        if (_integerDigits == 0 && text.StartsWith("0.", StringComparison.Ordinal))
        {
            text = text.Substring(1);
        }

        return text;
    }

    private enum TokenKind
    {
        Literal,
        IntegerPlaceholder,
        DecimalPoint,
        RequiredDecimalDigit,
    }

    private readonly struct Token
    {
        private Token(TokenKind kind, string text)
        {
            Kind = kind;
            Text = text;
        }

        public TokenKind Kind { get; }

        public string Text { get; }

        public static Token Literal(string text) => new(TokenKind.Literal, text);

        public static Token IntegerPlaceholder() => new(TokenKind.IntegerPlaceholder, string.Empty);

        public static Token DecimalPoint() => new(TokenKind.DecimalPoint, string.Empty);

        public static Token RequiredDecimalDigit() => new(TokenKind.RequiredDecimalDigit, string.Empty);
    }
}
