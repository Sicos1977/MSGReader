// name       : RtfHtmlSpecialCharCollection.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2010.12.01
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections.Generic;
using System.Text;

namespace Itenso.Rtf.Converter.Html
{
    public class RtfHtmlSpecialCharCollection : Dictionary<RtfVisualSpecialCharKind, string>
    {
        public RtfHtmlSpecialCharCollection()
        {
        } // RtfHtmlSpecialCharCollection

        public RtfHtmlSpecialCharCollection(string settings)
        {
            LoadSettings(settings);
        } // RtfHtmlSpecialCharCollection

        public void LoadSettings(string settings)
        {
            Clear();
            if (string.IsNullOrEmpty(settings))
                return;

            var settingItems = settings.Split(',');
            foreach (var settingItem in settingItems)
            {
                var tokens = settingItem.Split('=');
                if (tokens.Length != 2)
                    continue;

                var charKind = (RtfVisualSpecialCharKind) Enum.Parse(typeof(RtfVisualSpecialCharKind), tokens[0]);
                Add(charKind, tokens[1]);
            }
        } // LoadSettings

        public string GetSettings()
        {
            if (Count == 0)
                return string.Empty;

            var sb = new StringBuilder();
            foreach (var charKind in Keys)
            {
                if (sb.Length > 0)
                    sb.Append(',');
                sb.Append(Enum.GetName(typeof(RtfVisualSpecialCharKind), charKind));
                sb.Append('=');
                sb.Append(this[charKind]);
            }

            return sb.ToString();
        } // GetSettings
    }
}