using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace SheetReader.Wpf.Test
{
    public partial class ProgressWindow : Window
    {
        private bool _mustBeClosed;

        public ProgressWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        public void DoProgress(string text) => Dispatcher.Invoke(() =>
        {
            lbl.Content = text ?? string.Empty;
        });

        public void DoFinished() => Dispatcher.Invoke(() =>
        {
            _mustBeClosed = true;
            pg.IsIndeterminate = false;
            pg.Value = pg.Maximum;
            DialogResult = true;
            try
            {
                Close();
            }
            catch
            {
                // in case the confirmation MessageBox is shown, this will fail here
            }
        });

        private void OnClosing(object sender, CancelEventArgs e)
        {
            if (_mustBeClosed)
                return;

            // since we're multi-threaded, _mustBeClosed may have changed
            e.Cancel = MessageBox.Show(this,
                "Are you sure you want to cancel the current operation?",
                Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title,
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes;

            if (!e.Cancel)
            {
                DialogResult = false;
            }

            if (_mustBeClosed)
            {
                e.Cancel = false;
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
        }
    }
}
