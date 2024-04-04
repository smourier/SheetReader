using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        public ObservableCollection<SheetData> Sheets { get; } = [];

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            MainMenu.RaiseMenuItemClickOnKeyGesture(e);

            if (e.Key == Key.T && Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
            {
                var lastRecent = Settings.Current.RecentFilesPaths?.FirstOrDefault();
                if (lastRecent != null)
                {
                    Load(lastRecent.FilePath);
                }
            }
        }

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

            Load(ofd.FileName);
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
                    item.Click += (s, e) => Load(recent.FilePath);
                }
            }

            RecentFilesMenuItem.IsEnabled = RecentFilesMenuItem.Items.Count > fixedRecentItemsCount;
        }

        private void Load(string? fileName)
        {
            if (fileName == null || fileName.EqualsIgnoreCase(FileName))
                return;

            Sheets.Clear();
            var book = new Book();
            try
            {
                foreach (var sheet in book.EnumerateSheets(fileName))
                {
                    // we need to load data for WPF data binding
                    var sheetData = new SheetData { Name = sheet.Name, IsHidden = !sheet.IsVisible };
                    foreach (var row in sheet.EnumerateRows())
                    {
                        var rowData = new RowData(row.Index, row.EnumerateCells().ToList());
                        sheetData.Rows.Add(rowData);
                    }

                    sheetData.Columns = sheet.EnumerateColumns().ToList();
                    if (sheetData.Columns.Count > 0)
                    {
                        sheetData.LastColumnIndex = sheetData.Columns.Max(x => x.Index);
                    }

                    if (sheetData.Rows.Count > 0)
                    {
                        sheetData.LastRowIndex = sheetData.Rows.Max(x => x.RowIndex);
                    }
                    Sheets.Add(sheetData);
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
                var dummy = new SheetData();
                Sheets.Add(dummy);
                Sheets.Remove(dummy);
            }
            Title = "Sheet Reader - " + Path.GetFileName(fileName);

            FileName = fileName;
            Settings.Current.AddRecentFile(fileName);
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                foreach (var file in files.Where(f => Book.IsSupportedFileExtension(Path.GetExtension(f))))
                {
                    Load(file);
                }
            }
        }
    }
}
