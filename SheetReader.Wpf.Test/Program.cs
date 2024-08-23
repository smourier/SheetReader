using System;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using SheetReader.Wpf.Test.Utilities;

namespace SheetReader.Wpf.Test
{
    internal static class Program
    {
        private static bool _errorShown;

        [STAThread]
        static void Main()
        {
            var app = new App();
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;
            app.InitializeComponent();
            app.Run();
        }

        public static void ShowError(Window? window, Exception error, bool shutdownOnClose)
        {
            if (error == null || _errorShown)
                return;

            _errorShown = true;
            var shutdown = false;

            if (window == null)
            {
                MessageBox.Show(
                    "An error has occured: " + error.GetAllMessagesWithDots(),
                    Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyDescriptionAttribute>()?.Description,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            else
            {
                MessageBox.Show(
                    window,
                    "An error has occured: " + error.GetAllMessagesWithDots(),
                    Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyDescriptionAttribute>()?.Description,
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            if (shutdownOnClose || shutdown)
            {
                Application.Current.Shutdown();
            }

            _errorShown = false;
        }

        private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e) => ShowError(null, e.Exception, false);
        private static void Current_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            ShowError(null, e.Exception, false);
        }
    }
}