using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace SheetReader.Wpf.Test
{
    public partial class JsonOptions : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private bool _asObjects;
        private bool _indented;
        private bool _firstRowDefinesColumns;
        private bool _noDefaultCellValues;

        public JsonOptions()
        {
            InitializeComponent();
            DataContext = this;

#if DEBUG
            Indented = true;
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

        public bool Indented
        {
            get => _indented;
            set
            {
                if (_indented == value)
                    return;

                _indented = value;
                OnPropertyChanged();
            }
        }

        public bool AsObjects
        {
            get => _asObjects;
            set
            {
                if (_asObjects == value)
                    return;

                _asObjects = value;
                OnPropertyChanged();
            }
        }

        public bool NoDefaultCellValues
        {
            get => _noDefaultCellValues;
            set
            {
                if (_noDefaultCellValues == value)
                    return;

                _noDefaultCellValues = value;
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
