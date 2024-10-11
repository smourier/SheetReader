using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("SheetReader.Wpf.Test")]
[assembly: AssemblyDescription("Sheet Reader Wpf Test")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: AssemblyProduct("Sheet Reader")]
[assembly: AssemblyCopyright("Copyright (C) 2021-2024 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: Guid("a51e10cf-4a86-4102-a90e-b2dda115e7a7")]
[assembly: SupportedOSPlatform("windows")]
