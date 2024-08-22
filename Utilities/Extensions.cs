using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

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
    }
}
