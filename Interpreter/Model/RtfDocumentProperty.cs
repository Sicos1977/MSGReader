// name       : RtfDocumentProperty.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfDocumentProperty : IRtfDocumentProperty
    {
        // members

        public RtfDocumentProperty(int propertyKindCode, string name, string staticValue) :
            this(propertyKindCode, name, staticValue, null)
        {
        } // RtfDocumentProperty

        public RtfDocumentProperty(int propertyKindCode, string name, string staticValue, string linkValue)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (staticValue == null)
                throw new ArgumentNullException("staticValue");
            PropertyKindCode = propertyKindCode;
            switch (propertyKindCode)
            {
                case RtfSpec.PropertyTypeInteger:
                    PropertyKind = RtfPropertyKind.IntegerNumber;
                    break;
                case RtfSpec.PropertyTypeRealNumber:
                    PropertyKind = RtfPropertyKind.RealNumber;
                    break;
                case RtfSpec.PropertyTypeDate:
                    PropertyKind = RtfPropertyKind.Date;
                    break;
                case RtfSpec.PropertyTypeBoolean:
                    PropertyKind = RtfPropertyKind.Boolean;
                    break;
                case RtfSpec.PropertyTypeText:
                    PropertyKind = RtfPropertyKind.Text;
                    break;
                default:
                    PropertyKind = RtfPropertyKind.Unknown;
                    break;
            }
            Name = name;
            StaticValue = staticValue;
            LinkValue = linkValue;
        } // RtfDocumentProperty

        public int PropertyKindCode { get; } // PropertyKindCode

        public RtfPropertyKind PropertyKind { get; } // PropertyKind

        public string Name { get; } // Name

        public string StaticValue { get; } // StaticValue

        public string LinkValue { get; } // LinkValue

        public override bool Equals(object obj)
        {
            if (obj == this)
                return true;

            if (obj == null || GetType() != obj.GetType())
                return false;

            return IsEqual(obj);
        } // Equals

        private bool IsEqual(object obj)
        {
            var compare = obj as RtfDocumentProperty; // guaranteed to be non-null
            return
                compare != null &&
                PropertyKindCode == compare.PropertyKindCode &&
                PropertyKind == compare.PropertyKind &&
                Name.Equals(compare.Name) &&
                CompareTool.AreEqual(StaticValue, compare.StaticValue) &&
                CompareTool.AreEqual(LinkValue, compare.LinkValue);
        } // IsEqual

        public override int GetHashCode()
        {
            return HashTool.AddHashCode(GetType().GetHashCode(), ComputeHashCode());
        } // GetHashCode

        private int ComputeHashCode()
        {
            var hash = PropertyKindCode;
            hash = HashTool.AddHashCode(hash, PropertyKind);
            hash = HashTool.AddHashCode(hash, Name);
            hash = HashTool.AddHashCode(hash, StaticValue);
            hash = HashTool.AddHashCode(hash, LinkValue);
            return hash;
        } // ComputeHashCode

        public override string ToString()
        {
            var buf = new StringBuilder(Name);
            if (StaticValue != null)
            {
                buf.Append("=");
                buf.Append(StaticValue);
            }
            if (LinkValue != null)
            {
                buf.Append("@");
                buf.Append(LinkValue);
            }
            return buf.ToString();
        } // ToString
    } // class RtfDocumentProperty
}