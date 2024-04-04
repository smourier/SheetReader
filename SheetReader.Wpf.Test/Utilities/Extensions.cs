using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SheetReader.Wpf.Test.Utilities
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
                    if (string.IsNullOrWhiteSpace(msg))
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

        public static void RaiseMenuItemClickOnKeyGesture(this ItemsControl control, KeyEventArgs args, bool throwOnError = true)
        {
            ArgumentNullException.ThrowIfNull(args);
            if (control == null)
                return;

            var kgc = new KeyGestureConverter();
            foreach (var item in control.Items.OfType<MenuItem>())
            {
                if (!string.IsNullOrWhiteSpace(item.InputGestureText))
                {
                    KeyGesture? gesture = null;
                    if (throwOnError)
                    {
                        gesture = kgc.ConvertFrom(item.InputGestureText) as KeyGesture;
                    }
                    else
                    {
                        try
                        {
                            gesture = kgc.ConvertFrom(item.InputGestureText) as KeyGesture;
                        }
                        catch
                        {
                        }
                    }

                    if (gesture != null && gesture.Matches(null, args))
                    {
                        item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                        args.Handled = true;
                        return;
                    }
                }

                RaiseMenuItemClickOnKeyGesture(item, args, throwOnError);
                if (args.Handled)
                    return;
            }
        }
    }
}
