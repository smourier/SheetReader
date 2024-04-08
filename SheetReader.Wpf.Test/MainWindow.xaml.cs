using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
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
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            tc.ItemsSource = Sheets;

            Task.Run(() => Settings.Current.CleanRecentFiles());
        }

        public string? FileName { get; set; }
        public ObservableCollection<BookDocumentSheet> Sheets { get; } = [];

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            MainMenu.RaiseMenuItemClickOnKeyGesture(e);

            if (e.Key == Key.T && Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                var lastRecent = Settings.Current.RecentFilesPaths?.FirstOrDefault();
                if (lastRecent != null)
                {
                    LoadDocument(lastRecent.FilePath);
                }
            }
        }

        private void OpenWithExcel_Click(object sender, RoutedEventArgs e) => OpenWithExcel();
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
                Filter = "Sheet Files|*.csv;*.xlsx;|All files|*"
            };
            if (ofd.ShowDialog() != true)
                return;

            LoadDocument(ofd.FileName);
        }

        private void SheetControl_SelectionChanged(object sender, RoutedEventArgs e)
        {
            var sc = (SheetControl)sender;
            selection.Text = "Selection:" + sc.Selection.ToString();
        }

        private void OpenWithExcel()
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
                    item.Click += (s, e) => LoadDocument(recent.FilePath);
                }
            }

            RecentFilesMenuItem.IsEnabled = RecentFilesMenuItem.Items.Count > fixedRecentItemsCount;
        }

        private void LoadDocument(string? filePath)
        {
            Sheets.Clear();
            if (filePath != null)
            {
                try
                {
                    var format = BookFormat.GetFromFileExtension(Path.GetExtension(filePath)) ?? new CsvBookFormat();
                    var book = new ConcurrentBookDocument();
                    book.Load(filePath, format);
                    foreach (var sheet in book.Sheets)
                    {
                        Sheets.Add(sheet);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetAllMessagesWithDots(), "Error");
                    return;
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
                Settings.Current.AddRecentFile(filePath);
            }
            else
            {
                Title = "Sheet Reader";
            }
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                foreach (var file in files.Where(f => Book.IsSupportedFileExtension(Path.GetExtension(f))))
                {
                    LoadDocument(file);
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
                    status.Text = $"Cell: {result.RowCol.ExcelReference}: {cell.Value}";
                    return;
                }
            }
            status.Text = string.Empty;
        }

        private sealed class ConcurrentBookDocument : BookDocument
        {
            protected override BookDocumentSheet CreateSheet(Sheet sheet) => new ConcurrentBookDocumentSheet(sheet);
            public override bool IsThreadSafe => true;
        }

        private sealed class ConcurrentBookDocumentSheet(Sheet sheet) : BookDocumentSheet(sheet)
        {
            protected override IDictionary<int, Column> CreateColumns() => new ConcurrentDictionary<int, Column>();
            protected override IDictionary<int, BookDocumentRow> CreateRows() => new ConcurrentDictionary<int, BookDocumentRow>();
            protected override BookDocumentRow CreateRow(Row row) => new ConcurentBookDocumentRow(row);
        }

        private sealed class ConcurentBookDocumentRow(Row row) : BookDocumentRow(row)
        {
            protected override IDictionary<int, BookDocumentCell> CreateCells() => new ConcurrentDictionary<int, BookDocumentCell>();

            protected override BookDocumentCell CreateCell(Cell cell)
            {
                if (!cell.IsError)
                {
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
    }
}
