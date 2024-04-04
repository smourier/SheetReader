using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Markup;

#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("SheetReader.Wpf")]
[assembly: AssemblyDescription("Sheet Reader Wpf")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: AssemblyProduct("Sheet Reader")]
[assembly: AssemblyCopyright("Copyright (C) 2021-2024 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: Guid("490b9e09-43d3-4a44-b945-d9a2542d63a3")]
[assembly: ThemeInfo(ResourceDictionaryLocation.None, ResourceDictionaryLocation.SourceAssembly)]
[assembly: XmlnsPrefix("http://schemas.simonmourier.com/xaml/sheetreader", "smx")]
[assembly: XmlnsDefinition("http://schemas.simonmourier.com/xaml/sheetreader", "SheetReader.Wpf")]

