// name       : RtfDocumentInfo.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Model
{
    public sealed class RtfDocumentInfo : IRtfDocumentInfo
    {
        // members

        public int? Id { get; set; } // Id

        public int? Version { get; set; } // Version

        public int? Revision { get; set; } // Revision

        public string Title { get; set; } // Title

        public string Subject { get; set; } // Subject

        public string Author { get; set; } // Author

        public string Manager { get; set; } // Manager

        public string Company { get; set; } // Company

        public string Operator { get; set; } // Operator

        public string Category { get; set; } // Category

        public string Keywords { get; set; } // Keywords

        public string Comment { get; set; } // Comment

        public string DocumentComment { get; set; } // DocumentComment

        public string HyperLinkbase { get; set; } // HyperLinkbase

        public DateTime? CreationTime { get; set; } // CreationTime

        public DateTime? RevisionTime { get; set; } // RevisionTime

        public DateTime? PrintTime { get; set; } // PrintTime

        public DateTime? BackupTime { get; set; } // BackupTime

        public int? NumberOfPages { get; set; } // NumberOfPages

        public int? NumberOfWords { get; set; } // NumberOfWords

        public int? NumberOfCharacters { get; set; } // NumberOfCharacters

        public int? EditingTimeInMinutes { get; set; } // EditingTimeInMinutes

        public void Reset()
        {
            Id = null;
            Version = null;
            Revision = null;
            Title = null;
            Subject = null;
            Author = null;
            Manager = null;
            Company = null;
            Operator = null;
            Category = null;
            Keywords = null;
            Comment = null;
            DocumentComment = null;
            HyperLinkbase = null;
            CreationTime = null;
            RevisionTime = null;
            PrintTime = null;
            BackupTime = null;
            NumberOfPages = null;
            NumberOfWords = null;
            NumberOfCharacters = null;
            EditingTimeInMinutes = null;
        } // Reset

        public override string ToString()
        {
            return "RTFDocInfo";
        } // ToString
    }
}