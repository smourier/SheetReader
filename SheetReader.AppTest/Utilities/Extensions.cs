using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SheetReader.AppTest.Utilities
{
    internal static class Extensions
    {
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
