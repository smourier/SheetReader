using System;
using System.Collections;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace SheetReader.Wpf.Test
{
    public partial class OpenUrl : Window, INotifyPropertyChanged, INotifyDataErrorInfo, IDataErrorInfo
    {
        private string? _url;

        public event PropertyChangedEventHandler? PropertyChanged;
        public event EventHandler<DataErrorsChangedEventArgs>? ErrorsChanged;

        public OpenUrl()
        {
            InitializeComponent();
            DataContext = this;
        }

        public string? Url
        {
            get => _url;
            set
            {
                if (_url == value)
                    return;

                _url = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValid));
                OnErrorsChanged();
            }
        }

        public bool IsValid => !((INotifyDataErrorInfo)this).HasErrors;
        string IDataErrorInfo.Error => ((IDataErrorInfo)this)[null!];
        bool INotifyDataErrorInfo.HasErrors => ((IDataErrorInfo)this).Error != null;
        string IDataErrorInfo.this[string columnName] => ((INotifyDataErrorInfo)this).GetErrors(columnName).OfType<string>().FirstOrDefault()!;
        IEnumerable INotifyDataErrorInfo.GetErrors(string? propertyName)
        {
            if (propertyName == null || propertyName == nameof(Url))
            {
                if (!Uri.TryCreate(Url, UriKind.Absolute, out var uri))
                {
                    if (string.IsNullOrWhiteSpace(Url))
                        yield return $"Url cannot be empty";
                    else
                        yield return $"Url '{Url}' is invalid";
                }
                else if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
                    yield return $"Url must be http or https, {uri.Scheme} is unsupported";
            }
        }

        private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        private void OnErrorsChanged([CallerMemberName] string? name = null) => ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(name));

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
    }
}
