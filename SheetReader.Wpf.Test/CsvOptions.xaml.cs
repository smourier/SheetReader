using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace SheetReader.Wpf.Test
{
    public partial class CsvOptions : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private bool _firstRowDefinesColumns;
        private bool _csvWriteColumns;

        public CsvOptions()
        {
            InitializeComponent();
            DataContext = this;

#if DEBUG
            CsvWriteColumns = true;
#endif
        }

        public bool Open { get; set; }

        public bool FirstRowDefinesColumns
        {
            get => _firstRowDefinesColumns;
            set
            {
                if (_firstRowDefinesColumns == value)
                    return;

                _firstRowDefinesColumns = value;
                OnPropertyChanged();
            }
        }

        public bool CsvWriteColumns
        {
            get => _csvWriteColumns;
            set
            {
                if (_csvWriteColumns == value)
                    return;

                _csvWriteColumns = value;
                OnPropertyChanged();
            }
        }

        private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
        }

        private void OnCancelClick(object sender, RoutedEventArgs e) => Close();
        private void OnOKClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void OnOKOpenClick(object sender, RoutedEventArgs e)
        {
            Open = true;
            DialogResult = true;
            Close();
        }
    }
}
