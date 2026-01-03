using System.Reflection;
using System.Runtime.InteropServices;

#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("SheetReader")]
[assembly: AssemblyDescription("Sheet Reader")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: AssemblyProduct("Sheet Reader")]
[assembly: AssemblyCopyright("Copyright (C) 2021-2026 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: ComVisible(false)]
[assembly: Guid("f602601a-5246-4783-a83b-2acd7b5ace00")]
