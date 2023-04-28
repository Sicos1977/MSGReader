using System.Collections;
using System.Text;

namespace MsgReader.Rtf;

/// <summary>
///     Font table
/// </summary>
internal class Table : CollectionBase
{
    #region ByIndex
    /// <summary>
    ///     Get font information special index
    /// </summary>
    public Font this[int fontIndex]
    {
        get
        {
            foreach (Font font in this)
                if (font.Index == fontIndex)
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

    #region GetFontName
    /// <summary>
    ///     Get font object special font index
    /// </summary>
    /// <param name="fontIndex">Font index</param>
    /// <returns>font object</returns>
    public string GetFontName(int fontIndex)
    {
        var font = this[fontIndex];
        return font?.Name;
    }
    #endregion

    #region Add
    /// <summary>
    ///     Add font
    /// </summary>
    /// <param name="name">Font name</param>
    public Font Add(string name)
    {
        return Add(Count, name, Encoding.Default);
    }

    /// <summary>
    ///     Add font
    /// </summary>
    /// <param name="name">font name</param>
    /// <param name="encoding"></param>
    public Font Add(string name, Encoding encoding)
    {
        return Add(Count, name, encoding);
    }

    /// <summary>
    ///     Add font
    /// </summary>
    /// <param name="index">special font index</param>
    /// <param name="name">font name</param>
    /// <param name="encoding"></param>
    public Font Add(int index, string name, Encoding encoding)
    {
        if (this[name] == null)
        {
            var font = new Font(index, name);
            if (encoding != null) font.Charset = Font.GetCharset(encoding);
            List.Add(font);
            return font;
        }

        return this[name];
    }

    /// <summary>
    ///     Add font
    /// </summary>
    /// <param name="font">Font object</param>
    public void Add(Font font)
    {
        List.Add(font);
    }
    #endregion

    #region Remove
    /// <summary>
    ///     Remove font
    /// </summary>
    /// <param name="name">font name</param>
    public void Remove(string name)
    {
        var font = this[name];
        if (font != null)
            List.Remove(font);
    }

    /// <summary>
    ///     Get font index special font name
    /// </summary>
    /// <param name="name">font name</param>
    /// <returns>font index</returns>
    public int IndexOf(string name)
    {
        foreach (Font font in this)
            if (font.Name == name)
                return font.Index;
        return -1;
    }
    #endregion

    #region Clone
    /// <summary>
    ///     Clone object
    /// </summary>
    /// <returns>new object</returns>
    public Table Clone()
    {
        var table = new Table();
        foreach (Font item in this)
        {
            var newItem = item.Clone();
            table.List.Add(newItem);
        }

        return table;
    }
    #endregion
}