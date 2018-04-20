// name       : ApplicationInfo.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Globalization;
using System.Reflection;

namespace Itenso.Sys.Application
{
    public static class ApplicationInfo
    {
        public static Version Version
        {
            get
            {
                var assembly = Assembly.GetEntryAssembly();
                if (assembly == null)
                    return null;
                return assembly.GetName().Version;
            }
        } // Version

        public static string VersionName
        {
            get
            {
                var version = Version;
                if (version == null)
                    return null;
                return version.ToString();
            }
        } // VersionName

        public static string ShortVersionName
        {
            get
            {
                var version = Version;
                if (version == null)
                    return null;
                return string.Concat(
                    version.Major.ToString(CultureInfo.InvariantCulture),
                    ".",
                    version.Minor.ToString(CultureInfo.InvariantCulture));
            }
        } // ShortVersionName

        public static string Title
        {
            get
            {
                string title = null;

                var assembly = Assembly.GetEntryAssembly();
                var attributes =
                    assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false)
                        as AssemblyTitleAttribute[];

                if (attributes != null && attributes.Length == 1)
                    title = attributes[0].Title;
                return title;
            }
        } // Title

        public static string Copyright
        {
            get
            {
                string copyright = null;

                var assembly = Assembly.GetEntryAssembly();
                var attributes =
                    assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false)
                        as AssemblyCopyrightAttribute[];

                if (attributes != null && attributes.Length == 1)
                    copyright = attributes[0].Copyright;
                return copyright;
            }
        } // Copyright

        public static string Caption
        {
            get { return string.Concat(Title, " ", VersionName); }
        } // Caption

        public static string ShortCaption
        {
            get { return string.Concat(Title, " ", ShortVersionName); }
        } // ShortCaption
    }
}