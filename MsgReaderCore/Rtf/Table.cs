using System.Collections;
using System.Text;

namespace MsgReader.Rtf;

/// <summary>
///     Font table
/// </summary>
internal class FontTable : CollectionBase
{
    #region ByIndex
    /// <summary>
    ///     Get font information special index
    /// </summary>
    public Font this[int index]
    {
        get
        {
            foreach (Font font in this)
                if (font.Index == index)
                    return font;
            return null;
        }
    }
    #endregion

    #region ByName
    /// <summary>
    ///     Get font object special name
    /// </summary>
    /// <param name="fontName">font name</param>
    /// <returns>font object</returns>
    public Font this[string fontName]
    {
        get
        {
            foreach (Font font in this)
                if (font.Name == fontName)
                    return font;
            return null;
        }
    }
    #endregion

    #region Add
    /// <summary>
    ///     Add font
    /// </summary>
    /// <param name="font">Font object</param>
    public void Add(Font font)
    {
        List.Add(font);
    }
    #endregion

    #region Clone
    /// <summary>
    ///     Clone object
    /// </summary>
    /// <returns>new object</returns>
    public FontTable Clone()
    {
        var table = new FontTable();
        foreach (Font item in this)
        {
            var newItem = item.Clone();
            table.List.Add(newItem);
        }

        return table;
    }
    #endregion

    #region MixedEncodings
    /// <summary>
    ///     Returns <c>true</c> when mixed encodings are used in the font table
    /// </summary>
    public bool MixedEncodings
    {
        get
        {
            Encoding currentEncoding = null;

            foreach (Font font in this)
            {
                if (font.Encoding != null && currentEncoding == null)
                    currentEncoding = font.Encoding;
                else if (font.Encoding != null && !Equals(currentEncoding.IsSingleByte, font.Encoding.IsSingleByte))
                    return true;
            }

            return false;
        }
    }
    #endregion
}