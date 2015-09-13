using System;
using System.Reflection;
using System.Runtime.InteropServices;

// Following the best practices from: http://blogs.msdn.com/b/jjameson/archive/2009/04/03/shared-assembly-info-in-visual-studio-projects.aspx

// General Information about an assembly is controlled through the following set of attributes. Change these attribute
// values to modify the information associated with an assembly.
[assembly: AssemblyCompany("Blockworks Software")]
[assembly: AssemblyProduct("Leave Management")]
[assembly: AssemblyCopyright("Copyright 2015")]
[assembly: AssemblyTrademark("")]
[assembly: CLSCompliant(true)]

// Make it easy to distinguish Debug and Release builds; for example, through the file properties window.
#if DEBUG
[assembly: AssemblyConfiguration("Debug")]
[assembly: AssemblyDescription("Flavor=Debug")] // a.k.a. "Comments"
#else
[assembly: AssemblyConfiguration("Release")]
[assembly: AssemblyDescription("Flavor=Release")] // a.k.a. "Comments"
#endif

// Setting ComVisible to false makes the types in this assembly not visible to COM components. If you need to access a
// type in this assembly from COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// Note that the assembly version does not get incremented for every build to avoid problems with assembly binding (or
// requiring a policy or <bindingRedirect> in the config file).
[assembly: AssemblyVersion("0.9.0.0")]

// By default, the "Product version" shown in the file properties window is the same as the value specified for
// AssemblyFileVersionAttribute. Set AssemblyInformationalVersionAttribute to be the same as AssemblyVersionAttribute so
// that the "Product version" in the file properties window matches the version displayed in the GAC shell extension.
[assembly: AssemblyInformationalVersion("0.9.0.0")] // a.k.a. "Product version"

// The AssemblyFileVersionAttribute is incremented with every build in order to distinguish one build from another.
// AssemblyFileVersion is specified in AssemblyVersionInfo.cs so that it can be easily incremented by the automated
// build process.
[assembly: AssemblyFileVersion("0.9.486.0")]
