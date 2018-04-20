// name       : RtfHtmlElementPath.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.09
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections;
using System.Text;
using System.Web.UI;

namespace Itenso.Rtf.Converter.Html
{
    public class RtfHtmlElementPath
    {
        // Members
        private readonly Stack _elements = new Stack();

        public int Count => _elements.Count;
        // Count

        public HtmlTextWriterTag Current => (HtmlTextWriterTag) _elements.Peek();
        // Current

        public bool IsCurrent(HtmlTextWriterTag tag)
        {
            return Current == tag;
        } // IsCurrent

        public bool Contains(HtmlTextWriterTag tag)
        {
            return _elements.Contains(tag);
        } // Contains

        public void Push(HtmlTextWriterTag tag)
        {
            _elements.Push(tag);
        } // Push

        public void Pop()
        {
            _elements.Pop();
        } // Pop

        public override string ToString()
        {
            if (_elements.Count == 0)
                return base.ToString();

            var sb = new StringBuilder();
            var first = true;
            foreach (var element in _elements)
            {
                if (!first)
                    sb.Insert(0, " > ");
                sb.Insert(0, element.ToString());
                first = false;
            }

            return sb.ToString();
        } // ToString
    }
}