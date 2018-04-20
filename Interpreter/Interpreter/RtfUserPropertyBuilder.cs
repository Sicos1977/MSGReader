// name       : RtfUserPropertyBuilder.cs
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
    public sealed class RtfUserPropertyBuilder : RtfElementVisitorBase
    {
        // Members
        private readonly RtfDocumentPropertyCollection _collectedProperties;
        private readonly RtfTextBuilder _textBuilder = new RtfTextBuilder();
        private string _linkValue;
        private string _propertyName;
        private int _propertyTypeCode;
        private string _staticValue;

        public RtfUserPropertyBuilder(RtfDocumentPropertyCollection collectedProperties) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (collectedProperties == null)
                throw new ArgumentNullException(nameof(collectedProperties));
            _collectedProperties = collectedProperties;
        } // RtfUserPropertyBuilder

        public IRtfDocumentProperty CreateProperty()
        {
            return new RtfDocumentProperty(_propertyTypeCode, _propertyName, _staticValue, _linkValue);
        } // CreateProperty

        public void Reset()
        {
            _propertyTypeCode = 0;
            _propertyName = null;
            _staticValue = null;
            _linkValue = null;
        } // Reset

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
                    _collectedProperties.Add(CreateProperty());
                    break;
                case RtfSpec.TagUserPropertyName:
                    _textBuilder.Reset();
                    _textBuilder.VisitGroup(group);
                    _propertyName = _textBuilder.CombinedText;
                    break;
                case RtfSpec.TagUserPropertyValue:
                    _textBuilder.Reset();
                    _textBuilder.VisitGroup(group);
                    _staticValue = _textBuilder.CombinedText;
                    break;
                case RtfSpec.TagUserPropertyLink:
                    _textBuilder.Reset();
                    _textBuilder.VisitGroup(group);
                    _linkValue = _textBuilder.CombinedText;
                    break;
            }
        } // DoVisitGroup

        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagUserPropertyType:
                    _propertyTypeCode = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag
    } // class RtfUserPropertyBuilder
}