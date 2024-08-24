using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using SheetReader.Wpf.Test.Utilities;

namespace SheetReader.Wpf.Test
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private string? _fileName;
        private BookDocument? _book;
        private static readonly HttpClient _httpClient = new();

        public MainWindow()
        {
            InitializeComponent();
            tc.ItemsSource = Sheets;
            DataContext = this;

            Task.Run(Settings.Current.CleanRecentFiles);
        }

        public bool HasFile => FileName != null;
        public bool HasNotTempFile => FileName != null && !FileName.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase);
        public ObservableCollection<BookDocumentSheet> Sheets { get; } = [];
        public string? FileName
        {
            get => _fileName;
            set
            {
                if (_fileName == value)
                    return;

                _fileName = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileName)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasFile)));
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            MainMenu.RaiseMenuItemClickOnKeyGesture(e);

            if (e.Key == Key.T && Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                var lastRecent = Settings.Current.RecentFilesPaths?.FirstOrDefault();
                if (lastRecent != null)
                {
                    LoadDocument(lastRecent);
                }
            }
        }

        private void OpenWithDefaultEditor_Click(object sender, RoutedEventArgs e) => OpenWithDefaultEditor();
        private void OpenFromUrl_Click(object sender, RoutedEventArgs e) => OpenFromUrl();
        private void ExportAsCsv_Click(object sender, RoutedEventArgs e) => Export(false);
        private void ExportAsJson_Click(object sender, RoutedEventArgs e) => Export(true);
        private void ClearRecentFiles_Click(object sender, RoutedEventArgs e) => Settings.Current.ClearRecentFiles();
        private void Exit_Click(object sender, RoutedEventArgs e) => Close();
        private void About_Click(object sender, RoutedEventArgs e) => MessageBox.Show(Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title + " - " + (IntPtr.Size == 4 ? "32" : "64") + "-bit" + Environment.NewLine + "Copyright (C) 2021-" + DateTime.Now.Year + " Simon Mourier. All rights reserved.", Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title, MessageBoxButton.OK, MessageBoxImage.Information);
        private void Open_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = ".xlsx",
                RestoreDirectory = true,
                Filter = "Sheet Files|*.csv;*.xlsx;*.json|All files|*"
            };
            if (ofd.ShowDialog() != true)
                return;

            var dlg = new FileLoadOptions() { Owner = this };
            if (dlg.ShowDialog() != true)
                return;

            var options = LoadOptions.None;
            if (dlg.FirstRowDefinesColumns)
            {
                options |= LoadOptions.FirstRowDefinesColumns;
            }

            LoadDocument(ofd.FileName, options, null, null);
        }

        private void Export(bool json)
        {
            if (!HasFile)
                return;

            string filter;
            string defaultExt;
            if (json)
            {
                filter = "*.json";
                defaultExt = ".json";
            }
            else
            {
                filter = "*.csv";
                defaultExt = ".csv";
            }

            var sfd = new SaveFileDialog
            {
                CheckPathExists = true,
                DefaultExt = defaultExt,
                RestoreDirectory = true,
                Filter = "Sheet Files|" + filter + "|All files|*"
            };

            if (sfd.ShowDialog() != true)
                return;

            bool openResult;
            var options = ExportOptions.None;
            if (json)
            {
                var dlg = new JsonOptions { Owner = this };
                if (!dlg.ShowDialog().GetValueOrDefault())
                    return;

                openResult = dlg.Open;
                if (dlg.AsObjects)
                {
                    options |= ExportOptions.JsonRowsAsObject;
                }

                if (dlg.FirstRowDefinesColumns)
                {
                    options |= ExportOptions.FirstRowDefinesColumns;
                }

                if (dlg.NoDefaultCellValues)
                {
                    options |= ExportOptions.JsonNoDefaultCellValues;
                }

                if (dlg.Indented)
                {
                    options |= ExportOptions.JsonIndented;
                }
            }
            else
            {
                var dlg = new CsvOptions { Owner = this };
                if (!dlg.ShowDialog().GetValueOrDefault())
                    return;

                openResult = dlg.Open;
                if (dlg.CsvWriteColumns)
                {
                    options |= ExportOptions.CsvWriteColumns;
                }

                if (dlg.FirstRowDefinesColumns)
                {
                    options |= ExportOptions.FirstRowDefinesColumns;
                }
            }

            _book!.Export(sfd.FileName, options);
            if (openResult)
            {
                Extensions.OpenFile(sfd.FileName);
            }
            else
            {
                MessageBox.Show("Data was successfully exported to '" + sfd.FileName + "'", Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title, MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void OpenFromUrl()
        {
            var dlg = new OpenUrl() { Owner = this };
            if (dlg.ShowDialog() != true)
                return;

            var options = LoadOptions.None;
            if (dlg.FirstRowDefinesColumns)
            {
                options |= LoadOptions.FirstRowDefinesColumns;
            }

            var uri = new Uri(dlg.Url!);
            _ = LoadDocumentFromUrl(uri, options);
        }

        private void SheetControl_SelectionChanged(object sender, RoutedEventArgs e)
        {
            var sc = (SheetControl)sender;
            selection.Text = "Selection:" + sc.Selection.ToString();
        }

        private void OpenWithDefaultEditor()
        {
            if (FileName == null)
                return;

            Extensions.OpenFile(FileName);
        }

        private void OnFileOpened(object sender, RoutedEventArgs e)
        {
            const int fixedRecentItemsCount = 2;
            while (RecentFilesMenuItem.Items.Count > fixedRecentItemsCount)
            {
                RecentFilesMenuItem.Items.RemoveAt(0);
            }

            var recents = Settings.Current.RecentFilesPaths;
            if (recents != null)
            {
                foreach (var recent in recents)
                {
                    var item = new MenuItem { Header = recent.FilePath };
                    RecentFilesMenuItem.Items.Insert(RecentFilesMenuItem.Items.Count - fixedRecentItemsCount, item);
                    item.Click += (s, e) => LoadDocument(recent);
                }
            }

            RecentFilesMenuItem.IsEnabled = RecentFilesMenuItem.Items.Count > fixedRecentItemsCount;
        }

        private void LoadDocument(RecentFile recent)
        {
            var options = recent.LoadOptions;
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                var dlg = new FileLoadOptions() { Owner = this, FirstRowDefinesColumns = options.HasFlag(LoadOptions.FirstRowDefinesColumns) };
                if (dlg.ShowDialog() != true)
                    return;

                if (dlg.FirstRowDefinesColumns)
                {
                    options |= LoadOptions.FirstRowDefinesColumns;
                }
                else
                {
                    options &= ~LoadOptions.FirstRowDefinesColumns;
                }
            }

            if (!Uri.TryCreate(recent.FilePath, UriKind.Absolute, out var uri) || uri.Scheme == Uri.UriSchemeFile)
            {
                LoadDocument(recent.FilePath, options, null, null);
            }
            else
            {
                _ = LoadDocumentFromUrl(uri, options);
            }
        }

        private async Task LoadDocumentFromUrl(Uri uri, LoadOptions options)
        {
            var fileName = uri.Segments.Last();
            var filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), fileName);
            IOUtilities.FileEnsureDirectory(filePath);

            var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using var response = await _httpClient.SendAsync(request).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            string? downloadedFileName = null;
            if (BookFormat.GetFromFileExtension(Path.GetExtension(filePath)) == null)
            {
                var ext = IOUtilities.GetExtensionFromContentType(response.Content.Headers.ContentType?.MediaType);
                downloadedFileName = response.Content.Headers.ContentDisposition?.FileNameStar ?? response.Content.Headers.ContentDisposition?.FileName;
                if (downloadedFileName == null)
                {
                    downloadedFileName = fileName + ext;
                }
                else if (string.IsNullOrEmpty(Path.GetExtension(downloadedFileName)))
                {
                    downloadedFileName += ext;
                }
            }

            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
            {
                using var file = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                await stream.CopyToAsync(file).ConfigureAwait(false);
            }

            if (downloadedFileName != null)
            {
                var downloadedFilePath = Path.Combine(Path.GetDirectoryName(filePath)!, downloadedFileName);
                IOUtilities.FileOverwrite(filePath, downloadedFilePath);
                IOUtilities.FileDelete(filePath, true, false);
                filePath = downloadedFilePath;
            }

            _ = Dispatcher.BeginInvoke(() => LoadDocument(filePath, options, uri.ToString(), fileName));
        }

        private void LoadDocument(string? filePath, LoadOptions options, string? url, string? fileName)
        {
            Sheets.Clear();
            if (filePath != null)
            {
                bool loadBook()
                {
                    var format = BookFormat.GetFromFileExtension(Path.GetExtension(filePath)) ?? new CsvBookFormat();
                    format.LoadOptions = options;
                    if (fileName != null)
                    {
                        format.Name = Path.GetFileNameWithoutExtension(fileName);
                    }

                    _book = new StyledBookDocument();

                    if (new FileInfo(filePath).Length > 2_000_000)
                    {
                        var dlg = new ProgressWindow() { Owner = this };
                        var keepRunning = true;
                        Task.Run(() =>
                        {
                            string? currentSheet = null;
                            var currentRow = 0;
                            _book.StateChanged += (s, e) =>
                            {
                                switch (e.Type)
                                {
                                    case StateChangedType.SheetAdded:
                                        currentSheet = e.Sheet.Name;
                                        currentRow = 0;
                                        break;

                                    case StateChangedType.RowAdded:
                                        if (e.Type == StateChangedType.RowAdded && (currentRow % 1000) == 0)
                                        {
                                            dlg.DoProgress($"Sheet '{currentSheet}' rows {currentRow}");
                                        }
                                        currentRow++;
                                        break;
                                }

                                if (!keepRunning)
                                {
                                    e.Cancel = true;
                                }
                            };

                            _book.Load(filePath, format);
                            dlg.DoFinished();
                        });

                        if (dlg.ShowDialog() == false)
                        {
                            keepRunning = false;
                            Sheets.Clear();
                            return false;
                        }
                    }
                    else
                    {
                        _book.Load(filePath, format);
                    }

                    foreach (var sheet in _book.Sheets)
                    {
                        Sheets.Add(sheet);
                    }

                    if (Sheets.Count > 1)
                    {
                        // this is so dumb, is there a better way
                        tc.SelectedIndex = 1;
                        Dispatcher.BeginInvoke(() => { tc.SelectedIndex = 0; });
                    }
                    return true;
                }

                if (Debugger.IsAttached)
                {
                    if (!loadBook())
                        return;
                }
                else
                {
                    try
                    {
                        if (!loadBook())
                            return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.GetAllMessagesWithDots(), "Error");
                        return;
                    }
                }

                // not sure why, but with only 1 tab, binding doesn't work...
                if (Sheets.Count == 1)
                {
                    var first = Sheets[0];
                    Sheets.Add(first);
                    Sheets.RemoveAt(1);
                }

                Title = "Sheet Reader - " + Path.GetFileName(filePath);
                FileName = filePath;

                if (url != null)
                {
                    Settings.Current.AddRecentFile(url, options);
                }
                else
                {
                    Settings.Current.AddRecentFile(filePath, options);
                }

                if (Sheets.Count > 0)
                {
                    // not sure there's a better way to focus on grid
                    var tabItem = (TabItem)tc.ItemContainerGenerator.ContainerFromItem(Sheets[0]);
                    Keyboard.Focus(tabItem);
                    tabItem.MoveFocus(new TraversalRequest(FocusNavigationDirection.Down));
                    tabItem.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                }

                if (url != null)
                {
                    IOUtilities.FileDelete(filePath, true, false);
                }
            }
            else
            {
                Title = "Sheet Reader";
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) &&
                e.Data.GetData(DataFormats.FileDrop) is string[] files &&
                files.Any(f => Book.IsSupportedFileExtension(Path.GetExtension(f))))
            {
                var dlg = new FileLoadOptions() { Owner = this };
                if (dlg.ShowDialog() != true)
                    return;

                var options = LoadOptions.None;
                if (dlg.FirstRowDefinesColumns)
                {
                    options |= LoadOptions.FirstRowDefinesColumns;
                }

                foreach (var file in files.Where(f => Book.IsSupportedFileExtension(Path.GetExtension(f))))
                {
                    LoadDocument(file, options, null, null);
                }
            }
        }

        private void SheetControl_MouseMove(object sender, MouseEventArgs e)
        {
            var ctl = (SheetControl)sender;
            var result = ctl.HitTest(e.GetPosition(ctl));
            if (result.RowCol != null)
            {
                if (result.IsOverRowHeader)
                {
                    status.Text = $"Row: {result.RowCol.RowIndex + 1}";
                    return;
                }

                if (result.IsOverColumnHeader)
                {
                    status.Text = $"Column: {Row.GetExcelColumnName(result.RowCol.ColumnIndex)} ({result.RowCol.ColumnIndex})";
                    return;
                }

                var cell = result.Cell;
                if (cell != null)
                {
                    status.Text = $"Cell: {result.RowCol.ExcelReference}: {ctl.Sheet.FormatValue(cell.Value)}";
                    return;
                }
            }
            status.Text = string.Empty;
        }

        private sealed class StyledBookDocument : BookDocument
        {
            protected override BookDocumentSheet CreateSheet(Sheet sheet) => new StyledBookDocumentSheet(this, sheet);
        }

        private sealed class StyledBookDocumentSheet(BookDocument book, Sheet sheet) : BookDocumentSheet(book, sheet)
        {
            protected override IDictionary<int, Column> CreateColumns() => new ConcurrentDictionary<int, Column>();
            protected override IDictionary<int, BookDocumentRow> CreateRows() => new ConcurrentDictionary<int, BookDocumentRow>();
            protected override BookDocumentRow CreateRow(BookDocument book, Row row) => new StyledBookDocumentRow(book, this, row);
        }

        private sealed class StyledBookDocumentRow(BookDocument book, BookDocumentSheet sheet, Row row) : BookDocumentRow(book, sheet, row)
        {
            protected override BookDocumentCell CreateCell(Cell cell)
            {
                if (!cell.IsError)
                {
                    if (cell.Value is IDictionary)
                        return new BookDocumentStyledCell(cell) { Style = DictionaryStyle.Instance };

                    if (cell.Value is Array)
                        return new BookDocumentStyledCell(cell) { Style = ArrayStyle.Instance };

                    // sample style : green + aligned right when detecting number
                    if (IsNumber(cell.Value) || IsParsableNumber(cell.Value))
                        return new BookDocumentStyledCell(cell) { Style = NumberStyle.Instance };

                    // sample style : orange + aligned right when detecting datetime
                    if (IsDateTime(cell.Value) || IsParsableDateTime(cell.Value))
                        return new BookDocumentStyledCell(cell) { Style = DateTimeStyle.Instance };
                }

                return base.CreateCell(cell);
            }

            private static bool IsNumber(object? value) =>
                value is int || value is sbyte || value is short || value is long ||
                value is uint || value is byte || value is ushort || value is ulong ||
                value is float || value is double || value is decimal;

            private static bool IsParsableNumber(object? value)
            {
                if (value == null) return false;
                var svalue = string.Format("{0}", value);
                if (string.IsNullOrEmpty(svalue)) return false;
                if (long.TryParse(svalue, out _)) return true;
                if (ulong.TryParse(svalue, out _)) return true;
                if (double.TryParse(svalue, out _)) return true;
                if (decimal.TryParse(svalue, out _)) return true;
                return false;
            }

            private static bool IsDateTime(object? value) => value is DateTime || value is DateTimeOffset;
            private static bool IsParsableDateTime(object? value)
            {
                if (value == null) return false;
                var svalue = string.Format("{0}", value);
                if (string.IsNullOrEmpty(svalue)) return false;
                if (DateTime.TryParse(svalue, null, out _)) return true;
                if (DateTimeOffset.TryParse(svalue, out _)) return true;
                return false;
            }
        }

        private sealed class NumberStyle : BookDocumentCellStyle
        {
            public static readonly NumberStyle Instance = new();
            public override Brush? Foreground { get => Brushes.Green; }
            public override TextAlignment? TextAlignment { get => System.Windows.TextAlignment.Right; }
            public override string ToString() => "Number";
        }

        private sealed class DateTimeStyle : BookDocumentCellStyle
        {
            public static readonly DateTimeStyle Instance = new();
            public override Brush? Foreground { get => Brushes.Orange; }
            public override TextAlignment? TextAlignment { get => System.Windows.TextAlignment.Right; }
            public override string ToString() => "DateTime";
        }

        private sealed class ArrayStyle : BookDocumentCellStyle
        {
            public static readonly ArrayStyle Instance = new();
            public override Brush? Background { get => Brushes.LightBlue; }
            public override TextAlignment? TextAlignment { get => System.Windows.TextAlignment.Center; }
            public override string ToString() => "Array";
        }

        private sealed class DictionaryStyle : BookDocumentCellStyle
        {
            public static readonly DictionaryStyle Instance = new();
            public override Brush? Background { get => Brushes.LightCoral; }
            public override TextAlignment? TextAlignment { get => System.Windows.TextAlignment.Center; }
            public override string ToString() => "Dictionary";
        }
    }
}
