using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using SheetReader.AppTest.Utilities;

namespace SheetReader.AppTest
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            MainMenu.RaiseMenuItemClickOnKeyGesture(e);
        }

        private void Exit_Click(object sender, RoutedEventArgs e) => Close();
        private void About_Click(object sender, RoutedEventArgs e) => MessageBox.Show(Assembly.GetEntryAssembly()!.GetCustomAttribute<AssemblyTitleAttribute>()!.Title + " - " + (IntPtr.Size == 4 ? "32" : "64") + "-bit" + Environment.NewLine + "Copyright (C) 2021-" + DateTime.Now.Year + " Simon Mourier. All rights reserved.", Assembly.GetEntryAssembly().GetCustomAttribute<AssemblyTitleAttribute>()!.Title, MessageBoxButton.OK, MessageBoxImage.Information);
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

            tc.Items.Clear();

            var book = new Book();
            foreach (var sheet in book.EnumerateSheets(ofd.FileName))
            {
                var data = new SheetData { Name = sheet.Name, IsHidden = !sheet.IsVisible };
                foreach (var row in sheet.EnumerateRows())
                {
                    data.Rows.Add(row.EnumerateCells().ToList());
                }
                tc.Items.Add(data);
            }
        }
    }
}
