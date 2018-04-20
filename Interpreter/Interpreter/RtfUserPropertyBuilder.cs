// -- FILE ------------------------------------------------------------------
// name       : RtfUserPropertyBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public sealed class RtfUserPropertyBuilder : RtfElementVisitorBase
    {
        // ----------------------------------------------------------------------
        // members
        private readonly RtfDocumentPropertyCollection collectedProperties;
        private readonly RtfTextBuilder textBuilder = new RtfTextBuilder();
        private string linkValue;
        private string propertyName;
        private int propertyTypeCode;
        private string staticValue;

        // ----------------------------------------------------------------------
        public RtfUserPropertyBuilder(RtfDocumentPropertyCollection collectedProperties) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (collectedProperties == null)
                throw new ArgumentNullException("collectedProperties");
            this.collectedProperties = collectedProperties;
        } // RtfUserPropertyBuilder

        // ----------------------------------------------------------------------
        public IRtfDocumentProperty CreateProperty()
        {
            return new RtfDocumentProperty(propertyTypeCode, propertyName, staticValue, linkValue);
        } // CreateProperty

        // ----------------------------------------------------------------------
        public void Reset()
        {
            propertyTypeCode = 0;
            propertyName = null;
            staticValue = null;
            linkValue = null;
        } // Reset

        // ----------------------------------------------------------------------
        protected override void DoVisitGroup(IRtfGroup group)
        {
            switch (group.Destination)
            {
                case RtfSpec.TagUserProperties:
                    VisitGroupChildren(group);
                    break;
                case null:
                    Reset();
                    VisitGroupChildren(group);
                    collectedProperties.Add(CreateProperty());
                    break;
                case RtfSpec.TagUserPropertyName:
                    textBuilder.Reset();
                    textBuilder.VisitGroup(group);
                    propertyName = textBuilder.CombinedText;
                    break;
                case RtfSpec.TagUserPropertyValue:
                    textBuilder.Reset();
                    textBuilder.VisitGroup(group);
                    staticValue = textBuilder.CombinedText;
                    break;
                case RtfSpec.TagUserPropertyLink:
                    textBuilder.Reset();
                    textBuilder.VisitGroup(group);
                    linkValue = textBuilder.CombinedText;
                    break;
            }
        } // DoVisitGroup

        // ----------------------------------------------------------------------
        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagUserPropertyType:
                    propertyTypeCode = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag
    } // class RtfUserPropertyBuilder
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------