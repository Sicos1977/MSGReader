// name       : VersionInfo.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.29
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Reflection;

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//

[assembly: AssemblyVersion("1.7.0.0")]

// ReSharper disable CheckNamespace

namespace Itenso.Rtf.Converter.Html
// ReSharper restore CheckNamespace
{
    public sealed class VersionInfo
    {
        /// <value>Provides easy access to the assemblies version as a string.</value>
        public static readonly string AssemblyVersion = typeof(VersionInfo).Assembly.GetName().Version.ToString();
    }
}