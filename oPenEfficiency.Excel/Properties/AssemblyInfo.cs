using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("openEfficiency for Excel")]
[assembly: AssemblyDescription("Excel-specific productivity features for openEfficiency")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("openEfficiency")]
[assembly: AssemblyProduct("openEfficiency for Excel")]
[assembly: AssemblyCopyright("Copyright © 2026 FabPei")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// COM visibility is opted-in per type (e.g. the ribbon). Keep the assembly hidden by default.
[assembly: ComVisible(false)]

[assembly: Guid("c1d2e3f4-2222-4333-8444-555566667777")]

// Assembly version and file version are stamped per build by Directory.Build.props
// (StampBuildIdentity -> BuildStampInfo.cs). Do not set them here again: a constant
// version lets the Office VSTO shadow copy keep serving an old DLL after a rebuild.
[assembly: InternalsVisibleTo("oPenEfficiency.Excel.Tests")]
