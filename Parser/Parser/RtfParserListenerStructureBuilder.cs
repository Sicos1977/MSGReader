// name       : RtfParserListenerStructureBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Parser
{
    public sealed class RtfParserListenerStructureBuilder : RtfParserListenerBase
    {
        // Members
        private readonly Stack _openGroupStack = new Stack();
        private RtfGroup _curGroup;
        private RtfGroup _structureRoot;

        public IRtfGroup StructureRoot => _structureRoot;
        

        protected override void DoParseBegin()
        {
            _openGroupStack.Clear();
            _curGroup = null;
            _structureRoot = null;
        } // DoParseBegin

        protected override void DoGroupBegin()
        {
            var newGroup = new RtfGroup();
            if (_curGroup != null)
            {
                _openGroupStack.Push(_curGroup);
                _curGroup.WritableContents.Add(newGroup);
            }
            _curGroup = newGroup;
        } // DoGroupBegin

        protected override void DoTagFound(IRtfTag tag)
        {
            if (_curGroup == null)
                throw new RtfStructureException(Strings.MissingGroupForNewTag);
            _curGroup.WritableContents.Add(tag);
        } // DoTagFound

        protected override void DoTextFound(IRtfText text)
        {
            _curGroup?.WritableContents.Add(text);
        } // DoTextFound

        protected override void DoGroupEnd()
        {
            if (_openGroupStack.Count > 0)
            {
                _curGroup = (RtfGroup) _openGroupStack.Pop();
            }
            else
            {
                if (_structureRoot != null)
                    throw new RtfStructureException(Strings.MultipleRootLevelGroups);
                _structureRoot = _curGroup;
                _curGroup = null;
            }
        } // DoGroupEnd

        protected override void DoParseEnd()
        {
            if (_openGroupStack.Count > 0)
                throw new RtfBraceNestingException(Strings.UnclosedGroups);
        } // DoParseEnd
    } // class RtfParserListenerStructureBuilder
}