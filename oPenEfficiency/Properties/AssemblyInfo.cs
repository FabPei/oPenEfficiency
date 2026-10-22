using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;

// Allgemeine Informationen über eine Assembly werden über folgende 
// Attribute gesteuert. Ändern Sie diese Attributwerte, um die Informationen zu ändern,
// die einer Assembly zugeordnet sind.
[assembly: AssemblyTitle("openEfficiency")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("FabPei")]
[assembly: AssemblyProduct("openEfficiency")]
[assembly: AssemblyCopyright("Copyright © 2026 FabPei")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Durch Festlegen von "ComVisible" auf "false" werden die Typen in dieser Assembly unsichtbar 
// für COM-Komponenten. Wenn Sie auf einen Typ in dieser Assembly von 
// COM aus zugreifen müssen, legen Sie das ComVisible-Attribut für diesen Typ auf "true" fest.
[assembly: ComVisible(false)]

// Die folgende GUID bestimmt die ID der Typbibliothek, wenn dieses Projekt für COM verfügbar gemacht wird
[assembly: Guid("4829b17d-9126-4577-a673-b3f0e91bede4")]

// Versionsinformationen für eine Assembly bestehen aus den folgenden vier Werten:
//
//      Hauptversion
//      Nebenversion 
//      Buildnummer
//      Revision
//
// Sie können alle Werte angeben oder die standardmäßigen Build- und Revisionsnummern 
// übernehmen, indem Sie "*" eingeben:
// Assembly version and file version are stamped per build by Directory.Build.props
// (StampBuildIdentity -> BuildStampInfo.cs). Do not set them here again: a constant
// version lets the Office VSTO shadow copy keep serving an old DLL after a rebuild.
[assembly: InternalsVisibleTo("oPenEfficiency.Command.Tests")]

