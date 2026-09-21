using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.ReportDefinitions
{
    /// <summary>
    /// 帳票定義JSONを <see cref="ReportDefinition"/> へ変換する。スキーマ検証もここで行う(要件6.3)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// JSON Schema ライブラリは導入せず、必須項目・型・値域を明示的に検証する。
    /// 帳票定義のスキーマは小さく固定的であり、依存を1つ減らせるため(`.kiro/steering/tech.md`
    /// 「依存ライブラリ追加時のルール」)。スキーマ本体は docs/帳票定義スキーマ.md に記述する。
    /// </para>
    /// <para>検証に失敗した場合は、どのプロパティが問題かを示す <see cref="ReportDefinitionSchemaException"/> を送出する。</para>
    /// </remarks>
    public static class ReportDefinitionJsonReader
    {
        /// <summary>帳票定義スキーマのバージョン。JSON 側の <c>schemaVersion</c> と一致する必要がある。</summary>
        public const int SupportedSchemaVersion = 1;

        /// <summary>JSON文字列を帳票定義へ変換する。</summary>
        /// <param name="json">帳票定義JSON。</param>
        /// <param name="definitionPath">エラーメッセージに含める定義ファイルのパス(任意)。</param>
        public static ReportDefinition Read(string json, string? definitionPath = null)
        {
            if (json is null)
            {
                throw new ArgumentNullException(nameof(json));
            }

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(json, new JsonDocumentOptions
                {
                    CommentHandling = JsonCommentHandling.Skip,
                    AllowTrailingCommas = true,
                });
            }
            catch (JsonException ex)
            {
                throw new ReportDefinitionSchemaException(
                    $"帳票定義がJSONとして解釈できません: {ex.Message}",
                    definitionPath: definitionPath,
                    innerException: ex);
            }

            using (document)
            {
                var root = document.RootElement;
                if (root.ValueKind != JsonValueKind.Object)
                {
                    throw Schema("帳票定義のルートはオブジェクトである必要があります。", null, definitionPath, "$");
                }

                ValidateSchemaVersion(root, definitionPath);

                var reportCode = RequireString(root, "reportCode", null, definitionPath);
                var sheetName = RequireString(root, "sheetName", reportCode, definitionPath);
                var fields = ReadSubstitutionFields(root, reportCode, definitionPath);
                var toleranceMm = ReadTolerance(root, reportCode, definitionPath);
                var unsupported = ReadUnsupportedPolicy(root, reportCode, definitionPath);
                var maxDigitWidth = ReadMaxDigitWidth(root, reportCode, definitionPath);
                var printAreaOverride = ReadPrintAreaOverride(root, reportCode, definitionPath);

                return new ReportDefinition(
                    reportCode, sheetName, fields, toleranceMm, unsupported, maxDigitWidth, printAreaOverride);
            }
        }

        private static void ValidateSchemaVersion(JsonElement root, string? definitionPath)
        {
            if (!root.TryGetProperty("schemaVersion", out var element))
            {
                throw Schema(
                    $"必須プロパティ 'schemaVersion' がありません(対応バージョン: {SupportedSchemaVersion})。",
                    null, definitionPath, "schemaVersion");
            }

            if (element.ValueKind != JsonValueKind.Number || !element.TryGetInt32(out var version))
            {
                throw Schema("'schemaVersion' は整数である必要があります。", null, definitionPath, "schemaVersion");
            }

            if (version != SupportedSchemaVersion)
            {
                throw Schema(
                    $"帳票定義のスキーマバージョン {version} には対応していません(対応バージョン: {SupportedSchemaVersion})。",
                    null, definitionPath, "schemaVersion");
            }
        }

        private static IReadOnlyList<SubstitutionFieldDefinition> ReadSubstitutionFields(
            JsonElement root, string reportCode, string? definitionPath)
        {
            if (!root.TryGetProperty("substitutionFields", out var array))
            {
                throw Schema("必須プロパティ 'substitutionFields' がありません。", reportCode, definitionPath, "substitutionFields");
            }

            if (array.ValueKind != JsonValueKind.Array)
            {
                throw Schema("'substitutionFields' は配列である必要があります。", reportCode, definitionPath, "substitutionFields");
            }

            var fields = new List<SubstitutionFieldDefinition>();
            var seenKeys = new HashSet<string>(StringComparer.Ordinal);
            var index = 0;

            foreach (var element in array.EnumerateArray())
            {
                var path = $"substitutionFields[{index}]";
                if (element.ValueKind != JsonValueKind.Object)
                {
                    throw Schema($"{path} はオブジェクトである必要があります。", reportCode, definitionPath, path);
                }

                var key = RequireString(element, "key", reportCode, definitionPath, path);
                if (!seenKeys.Add(key))
                {
                    throw Schema($"置換キー '{key}' が重複しています。", reportCode, definitionPath, $"{path}.key");
                }

                var cellText = RequireString(element, "cell", reportCode, definitionPath, path);
                if (!CellAddress.TryParse(cellText, out var cell))
                {
                    throw Schema(
                        $"'{cellText}' はセル番地として解釈できません(A1形式で指定してください)。",
                        reportCode, definitionPath, $"{path}.cell");
                }

                var required = ReadOptionalBool(element, "required", reportCode, definitionPath, path) ?? false;
                var overflow = ReadOverflow(element, reportCode, definitionPath, path);

                fields.Add(new SubstitutionFieldDefinition(key, cell, required, overflow));
                index++;
            }

            return fields;
        }

        /// <summary>
        /// はみ出し挙動を読む。未指定は <c>null</c>(= Excel 側のセル書式に従う)を返す。
        /// </summary>
        private static OverflowBehavior? ReadOverflow(
            JsonElement element, string reportCode, string? definitionPath, string path)
        {
            if (!element.TryGetProperty("overflow", out var value))
            {
                return null;
            }

            if (value.ValueKind != JsonValueKind.String)
            {
                throw Schema("'overflow' は文字列である必要があります。", reportCode, definitionPath, $"{path}.overflow");
            }

            return value.GetString() switch
            {
                "shrink" => OverflowBehavior.Shrink,
                "clip" => OverflowBehavior.Clip,
                "wrap" => OverflowBehavior.Wrap,
                "overflow" => OverflowBehavior.Overflow,
                var other => throw Schema(
                    $"'overflow' の値 '{other}' は不正です(shrink / clip / wrap / overflow のいずれか)。",
                    reportCode, definitionPath, $"{path}.overflow"),
            };
        }

        private static double ReadTolerance(JsonElement root, string reportCode, string? definitionPath) =>
            ReadOptionalRangedDouble(
                root, "toleranceMm", reportCode, definitionPath,
                ReportDefinition.DefaultToleranceMm, minValue: 0, minInclusive: true, rangeDescription: "0以上");

        private static double ReadMaxDigitWidth(JsonElement root, string reportCode, string? definitionPath) =>
            ReadOptionalRangedDouble(
                root, "maxDigitWidthPx", reportCode, definitionPath,
                ReportDefinition.DefaultMaxDigitWidthPx, minValue: 0, minInclusive: false, rangeDescription: "正の数");

        /// <summary>下限付きの任意の数値プロパティを読む。未指定時は <paramref name="defaultValue"/> を返す。</summary>
        private static double ReadOptionalRangedDouble(
            JsonElement root, string propertyName, string reportCode, string? definitionPath,
            double defaultValue, double minValue, bool minInclusive, string rangeDescription)
        {
            if (!root.TryGetProperty(propertyName, out var value))
            {
                return defaultValue;
            }

            if (value.ValueKind != JsonValueKind.Number || !value.TryGetDouble(out var number))
            {
                throw Schema($"'{propertyName}' は数値である必要があります。", reportCode, definitionPath, propertyName);
            }

            if (minInclusive ? number < minValue : number <= minValue)
            {
                throw Schema(
                    $"'{propertyName}' は{rangeDescription}である必要があります。", reportCode, definitionPath, propertyName);
            }

            return number;
        }

        private static UnsupportedElementPolicy ReadUnsupportedPolicy(
            JsonElement root, string reportCode, string? definitionPath)
        {
            if (!root.TryGetProperty("unsupportedElements", out var value))
            {
                return UnsupportedElementPolicy.Ignore;
            }

            if (value.ValueKind != JsonValueKind.String)
            {
                throw Schema(
                    "'unsupportedElements' は文字列である必要があります。", reportCode, definitionPath, "unsupportedElements");
            }

            return value.GetString() switch
            {
                "ignore" => UnsupportedElementPolicy.Ignore,
                "error" => UnsupportedElementPolicy.Error,
                var other => throw Schema(
                    $"'unsupportedElements' の値 '{other}' は不正です(ignore / error のいずれか)。",
                    reportCode, definitionPath, "unsupportedElements"),
            };
        }

        private static CellRange? ReadPrintAreaOverride(JsonElement root, string reportCode, string? definitionPath)
        {
            if (!root.TryGetProperty("printArea", out var value) || value.ValueKind == JsonValueKind.Null)
            {
                return null;
            }

            if (value.ValueKind != JsonValueKind.String)
            {
                throw Schema("'printArea' は文字列である必要があります。", reportCode, definitionPath, "printArea");
            }

            var text = value.GetString();
            if (!CellRange.TryParse(text, out var range))
            {
                throw Schema(
                    $"'{text}' はセル範囲として解釈できません(A1:H40 形式で指定してください)。",
                    reportCode, definitionPath, "printArea");
            }

            return range;
        }

        private static string RequireString(
            JsonElement parent, string propertyName, string? reportCode, string? definitionPath, string? parentPath = null)
        {
            var path = parentPath is null ? propertyName : $"{parentPath}.{propertyName}";

            if (!parent.TryGetProperty(propertyName, out var value))
            {
                throw Schema($"必須プロパティ '{path}' がありません。", reportCode, definitionPath, path);
            }

            if (value.ValueKind != JsonValueKind.String)
            {
                throw Schema($"'{path}' は文字列である必要があります。", reportCode, definitionPath, path);
            }

            var text = value.GetString();
            if (string.IsNullOrWhiteSpace(text))
            {
                throw Schema($"'{path}' が空文字です。", reportCode, definitionPath, path);
            }

            return text!;
        }

        private static bool? ReadOptionalBool(
            JsonElement parent, string propertyName, string reportCode, string? definitionPath, string parentPath)
        {
            if (!parent.TryGetProperty(propertyName, out var value))
            {
                return null;
            }

            return value.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                _ => throw Schema(
                    $"'{parentPath}.{propertyName}' は真偽値である必要があります。",
                    reportCode, definitionPath, $"{parentPath}.{propertyName}"),
            };
        }

        private static ReportDefinitionSchemaException Schema(
            string message, string? reportCode, string? definitionPath, string propertyPath) =>
            new(message, reportCode, definitionPath, propertyPath);
    }
}
