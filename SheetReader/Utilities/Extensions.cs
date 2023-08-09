using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SheetReader.Utilities
{
    internal static class Extensions
    {
        [return: NotNullIfNotNull(nameof(exception))]
        public static string? GetAllMessagesWithDots(this Exception? exception) => exception.GetAllMessages(".");

        [return: NotNullIfNotNull(nameof(exception))]
        public static string? GetAllMessages(this Exception? exception) => exception.GetAllMessages(Environment.NewLine);

        [return: NotNullIfNotNull(nameof(exception))]
        public static string? GetAllMessages(this Exception? exception, string? separator)
        {
            if (exception == null)
                return null;

            separator ??= Environment.NewLine;
            return string.Join(separator, EnumerateAllExceptionsMessages(exception).Distinct(StringComparer.OrdinalIgnoreCase));
        }

        public static IEnumerable<Exception> EnumerateAllExceptions(this Exception? exception)
        {
            if (exception == null)
                yield break;

            if (exception is AggregateException agg)
            {
                foreach (var ae in agg.InnerExceptions)
                {
                    foreach (var child in EnumerateAllExceptions(ae))
                    {
                        yield return child;
                    }
                }
            }
            else
            {
                if (exception is not TargetInvocationException) // useless
                    yield return exception;

                foreach (var child in EnumerateAllExceptions(exception.InnerException))
                {
                    yield return child;
                }
            }
        }

        public static IEnumerable<string> EnumerateAllExceptionsMessages(this Exception? exception)
        {
            foreach (var e in EnumerateAllExceptions(exception))
            {
                var typeName = GetExceptionTypeName(e);
                string message;
                if (string.IsNullOrWhiteSpace(typeName))
                {
                    message = e.Message;
                }
                else
                {
                    message = typeName + ": " + e.Message;
                }
                var normalized = norm(message);
                if (normalized != null)
                    yield return normalized;

                static string? norm(string? msg)
                {
                    msg = msg.Nullify();
                    if (msg == null)
                        return null;

                    if (!msg.EndsWith("."))
                    {
                        msg += ".";
                    }
                    return msg;
                }
            }
        }

        private static string? GetExceptionTypeName(Exception exception)
        {
            if (exception == null)
                return null;

            var type = exception.GetType();
            if (type == null || string.IsNullOrWhiteSpace(type.FullName))
                return null;

            if (type.FullName.StartsWith("System.") ||
                type.FullName.StartsWith("Microsoft."))
                return null;

            return type.FullName;
        }

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
