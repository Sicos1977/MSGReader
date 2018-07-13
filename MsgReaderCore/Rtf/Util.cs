namespace MsgReader.Rtf
{
    /// <summary>
    /// Some utility functions
    /// </summary>
    internal static class Util
    {
        #region HasContentElement
        /// <summary>
        /// Checks if the root element has content elemens
        /// </summary>
        /// <param name="rootElement"></param>
        /// <returns>True when there are content elements</returns>
        public static bool HasContentElement(DomElement rootElement)
        {
            if (rootElement.Elements.Count == 0)
            {
                return false;
            }
            if (rootElement.Elements.Count == 1)
            {
                if (rootElement.Elements[0] is DomParagraph)
                {
                    var p = (DomParagraph) rootElement.Elements[0];
                    if (p.Elements.Count == 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion
    }
}