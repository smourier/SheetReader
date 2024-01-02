using System.Reflection;
using System.Runtime.InteropServices;

#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("SheetReader.AppTest")]
[assembly: AssemblyDescription("Sheet Reader Test")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: AssemblyProduct("Sheet Reader")]
[assembly: AssemblyCopyright("Copyright (C) 2021-2024 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: Guid("c902c69f-baa5-46b8-8c1b-d8d99d103b66")]
