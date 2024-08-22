using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Text.Json;

namespace SheetReader.Utilities
{
    internal static class Extensions
    {
        public static bool EqualsIgnoreCase(this string? thisString, string? text, bool trim = false)
        {
            if (trim)
            {
                thisString = thisString.Nullify();
                text = text.Nullify();
            }

            if (thisString == null)
                return text == null;

            if (text == null)
                return false;

            if (thisString.Length != text.Length)
                return false;

            return string.Compare(thisString, text, StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static string? Nullify(this string? text)
        {
            if (text == null)
                return null;

            if (string.IsNullOrWhiteSpace(text))
                return null;

            var t = text.Trim();
            return t.Length == 0 ? null : t;
        }

        [return: NotNullIfNotNull(nameof(path))]
        public static string? NormPath(string path) => path?.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        public static bool FileEnsureDirectory(string path, bool throwOnError = true)
        {
            ArgumentNullException.ThrowIfNull(path);
            if (!Path.IsPathRooted(path))
            {
                path = Path.GetFullPath(path);
            }

            var dir = Path.GetDirectoryName(path);
            if (dir == null)
                return false;

            if (Directory.Exists(dir))
                return true;

            try
            {
                Directory.CreateDirectory(dir);
                return true;
            }
            catch
            {
                if (throwOnError)
                    throw;

                return false;
            }
        }

        public static bool? GetNullableBoolean(this JsonElement element, string propertyName)
        {
            ArgumentNullException.ThrowIfNull(propertyName);
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            if (!element.TryGetProperty(propertyName, out var property))
                return null;

            return property.ValueKind switch
            {
                JsonValueKind.False => false,
                JsonValueKind.True => true,
                _ => null,
            };
        }

        public static string? GetNullifiedString(this JsonElement element, string propertyName)
        {
            ArgumentNullException.ThrowIfNull(propertyName);
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            if (!element.TryGetProperty(propertyName, out var property))
                return null;

            if (property.ValueKind == JsonValueKind.String)
                return property.GetString();

            return property.GetRawText();
        }

        public static int? GetNullableInt32(this JsonElement element, string propertyName)
        {
            ArgumentNullException.ThrowIfNull(propertyName);
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            if (!element.TryGetProperty(propertyName, out var property))
                return null;

            if (property.ValueKind == JsonValueKind.Number)
            {
                if (property.TryGetInt32(out var i))
                    return i;

                return null;
            }

            var text = property.GetRawText();
            if (int.TryParse(text, out var value))
                return value;

            return null;
        }

        public static object? ToObject(this JsonElement element, JsonBookOptions options)
        {
            element.TryConvertToObject(options, out var value);
            return value;
        }

        public static bool TryConvertToObject(this JsonElement element, JsonBookOptions options, out object? value)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Null:
                    value = null;
                    return true;

                case JsonValueKind.Object:
                    var dic = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    foreach (var child in element.EnumerateObject())
                    {
                        if (!child.Value.TryConvertToObject(options, out var childValue))
                        {
                            value = null;
                            return false;
                        }

                        dic[child.Name] = childValue;
                    }

                    if (dic.Count == 0)
                    {
                        value = null;
                        return true;
                    }

                    value = dic;
                    return true;

                case JsonValueKind.Array:
                    var objects = new object?[element.GetArrayLength()];
                    var i = 0;
                    foreach (var child in element.EnumerateArray())
                    {
                        if (!child.TryConvertToObject(options, out var childValue2))
                        {
                            value = null;
                            return false;
                        }
                        objects[i++] = childValue2;
                    }
                    value = objects;
                    return true;

                case JsonValueKind.String:
                    var str = element.ToString();
                    if (options.HasFlag(JsonBookOptions.ParseDates) && DateTime.TryParseExact(str, ["o", "r", "s"], null, DateTimeStyles.None, out var dt))
                    {
                        value = dt;
                        return true;
                    }
                    value = str;
                    return true;

                case JsonValueKind.Number:
                    if (element.TryGetInt32(out var i2))
                    {
                        value = i2;
                        return true;
                    }

                    if (element.TryGetInt64(out var i64))
                    {
                        value = i64;
                        return true;
                    }

                    if (element.TryGetDecimal(out var dec))
                    {
                        value = dec;
                        return true;
                    }

                    if (element.TryGetDouble(out var dbl))
                    {
                        value = dbl;
                        return true;
                    }
                    break;

                case JsonValueKind.True:
                    value = true;
                    return true;

                case JsonValueKind.False:
                    value = false;
                    return true;
            }

            value = null;
            return false;
        }

        public static void WriteCsv(TextWriter writer, string cell, bool addSeparator = true, bool forExcel = true)
        {
            ArgumentNullException.ThrowIfNull(writer);

            var max = 32758;
            if (forExcel && cell != null && cell.Length > max)
            {
                cell = cell[..max];
            }

            if (cell != null && cell.IndexOfAny(['\t', '\r', '\n', '"']) >= 0)
            {
                writer.Write('"');
                writer.Write(cell.Replace("\"", "\"\""));
                writer.Write('"');
            }
            else
            {
                writer.Write(cell);
            }

            if (addSeparator)
            {
                writer.Write('\t');
            }
        }
    }
}
