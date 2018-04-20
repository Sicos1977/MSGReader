// name       : RtfDocumentInfoBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfDocumentInfoBuilder : RtfElementVisitorBase
    {
        // Members
        private readonly RtfDocumentInfo _info;
        private readonly RtfTextBuilder _textBuilder = new RtfTextBuilder();
        private readonly RtfTimestampBuilder _timestampBuilder = new RtfTimestampBuilder();

        public RtfDocumentInfoBuilder(RtfDocumentInfo info) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (info == null)
                throw new ArgumentNullException(nameof(info));
            _info = info;
        } // RtfDocumentInfoBuilder

        public void Reset()
        {
            _info.Reset();
        } // Reset

        protected override void DoVisitGroup(IRtfGroup group)
        {
            switch (group.Destination)
            {
                case RtfSpec.TagInfo:
                    VisitGroupChildren(group);
                    break;
                case RtfSpec.TagInfoTitle:
                    _info.Title = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoSubject:
                    _info.Subject = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoAuthor:
                    _info.Author = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoManager:
                    _info.Manager = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoCompany:
                    _info.Company = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoOperator:
                    _info.Operator = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoCategory:
                    _info.Category = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoKeywords:
                    _info.Keywords = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoComment:
                    _info.Comment = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoDocumentComment:
                    _info.DocumentComment = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoHyperLinkBase:
                    _info.HyperLinkbase = ExtractGroupText(group);
                    break;
                case RtfSpec.TagInfoCreationTime:
                    _info.CreationTime = ExtractTimestamp(group);
                    break;
                case RtfSpec.TagInfoRevisionTime:
                    _info.RevisionTime = ExtractTimestamp(group);
                    break;
                case RtfSpec.TagInfoPrintTime:
                    _info.PrintTime = ExtractTimestamp(group);
                    break;
                case RtfSpec.TagInfoBackupTime:
                    _info.BackupTime = ExtractTimestamp(group);
                    break;
            }
        } // DoVisitGroup

        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagInfoVersion:
                    _info.Version = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoRevision:
                    _info.Revision = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoNumberOfPages:
                    _info.NumberOfPages = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoNumberOfWords:
                    _info.NumberOfWords = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoNumberOfChars:
                    _info.NumberOfCharacters = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoId:
                    _info.Id = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoEditingTimeMinutes:
                    _info.EditingTimeInMinutes = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag

        private string ExtractGroupText(IRtfGroup group)
        {
            _textBuilder.Reset();
            _textBuilder.VisitGroup(group);
            return _textBuilder.CombinedText;
        } // ExtractGroupText

        private DateTime ExtractTimestamp(IRtfGroup group)
        {
            _timestampBuilder.Reset();
            _timestampBuilder.VisitGroup(group);
            return _timestampBuilder.CreateTimestamp();
        } // ExtractTimestamp
    } // class RtfDocumentInfoBuilder
}