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
                    var book = new BookDocument();
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
    }
}
