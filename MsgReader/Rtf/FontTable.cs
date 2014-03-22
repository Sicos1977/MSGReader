using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace DocumentServices.Modules.Readers.MsgReader.Rtf
{
    /// <summary>
    /// Font table
    /// </summary>
    internal class Table : CollectionBase
    {
        #region ByIndex
        /// <summary>
        /// Get font information special index
        /// </summary>
        public Font this[int fontIndex]
        {
            get
            {
                foreach (Font font in this)
                {
                    if (font.Index == fontIndex)
                        return font;
                }
                return null;
            }
        }
        #endregion

        #region ByName
        /// <summary>
        /// Get font object special name
        /// </summary>
        /// <param name="fontName">font name</param>
        /// <returns>font object</returns>
        public Font this[string fontName]
        {
            get
            {
                foreach (Font font in this)
                {
                    if (font.Name == fontName)
                    {
                        return font;
                    }
                }
                return null;
            }
        }
        #endregion

        #region GetFontName
        /// <summary>
        /// Get font object special font index
        /// </summary>
        /// <param name="fontIndex">Font index</param>
        /// <returns>font object</returns>
        public string GetFontName(int fontIndex)
        {
            var font = this[fontIndex];
            return font != null ? font.Name : null;
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
                if (encoding != null)
                {
                    font.Charset = Font.GetCharset(encoding);
                }
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
        /// Remove font
        /// </summary>
        /// <param name="name">font name</param>
        public void Remove(string name)
        {
            var font = this[name];
            if (font != null)
                List.Remove(font);
        }

        /// <summary>
        /// Get font index special font name
        /// </summary>
        /// <param name="f">font name</param>
        /// <returns>font index</returns>
        public int IndexOf(string name)
        {
            foreach (Font font in this)
            {
                if (font.Name == name)
                {
                    return font.Index;
                }
            }
            return -1;
        }
        #endregion

        #region Write
        /// <summary>
        /// Write font table rtf
        /// </summary>
        /// <param name="writer">rtf text writer</param>
        public void Write(Writer writer)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword(Consts.Fonttbl);
            foreach (Font font in this)
            {
                writer.WriteStartGroup();
                writer.WriteKeyword("f" + font.Index);
                if (font.Charset != 0)
                {
                    writer.WriteKeyword("fcharset" + font.Charset);
                }
                writer.WriteText(font.Name);
                writer.WriteEndGroup();
            }
            writer.WriteEndGroup();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var str = new StringBuilder();
            foreach (Font item in this)
            {
                str.Append(Environment.NewLine);
                str.Append("Index " + item.Index + "   Name:" + item.Name);
            }
            return str.ToString();
        }
        #endregion

        #region Clone
        /// <summary>
        /// Clone object
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

    /// <summary>
    /// Rtf font information
    /// </summary>
    internal class Font
    {
        #region Fields
        private static Dictionary<int, Encoding> _encodingCharsets;
        private int _charset = 1;
        #endregion

        #region Properties
        /// <summary>
        /// Font index
        /// </summary>
        public int Index { get; set; }

        public bool NilFlag { get; set; }

        /// <summary>
        ///     font name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Charset
        /// </summary>
        public int Charset
        {
            get { return _charset; }
            set
            {
                _charset = value;
                Encoding = GetRTFEncoding(_charset);
            }
        }

        /// <summary>
        /// Encoding
        /// </summary>
        public Encoding Encoding { get; private set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        /// <param name="index">font index</param>
        /// <param name="name">font name</param>
        public Font(int index, string name)
        {
            NilFlag = false;
            Encoding = null;
            Index = index;
            Name = name;
        }
        #endregion

        #region Clone
        public Font Clone()
        {
            return new Font(Index, Name)
            {
                _charset = _charset,
                Index = Index,
                Encoding = Encoding,
                Name = Name
            };
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return Index + ":" + Name + " Charset:" + _charset;
        }
        #endregion

        #region CheckEncodingCharsets
        private static void CheckEncodingCharsets()
        {
            if (_encodingCharsets == null)
            {
                _encodingCharsets = new Dictionary<int, Encoding>();
                _encodingCharsets[77] = Encoding.GetEncoding(10000);    // Mac ,macintosh Î÷Å·×Ö·û(Mac)
                _encodingCharsets[128] = Encoding.GetEncoding(932);     // Shift Jis ;ANSI/OEM - Japanese, Shift-JIS 
                _encodingCharsets[130] = Encoding.GetEncoding(1361);    // Johab;Korean (Johab) 
                _encodingCharsets[134] = Encoding.GetEncoding(936);     // GB2312
                _encodingCharsets[136] = Encoding.GetEncoding(10002);   // Big5
                _encodingCharsets[161] = Encoding.GetEncoding(1253);    // Greek
                _encodingCharsets[162] = Encoding.GetEncoding(1254);    // Turkish
                _encodingCharsets[163] = Encoding.GetEncoding(1258);    // Vietnamese;ANSI/OEM - Vietnamese 
                _encodingCharsets[177] = Encoding.GetEncoding(1255);    // Hebrw
                _encodingCharsets[178] = Encoding.GetEncoding(864);     // Arabic
                _encodingCharsets[179] = Encoding.GetEncoding(864);     // Arabic Traditional
                _encodingCharsets[180] = Encoding.GetEncoding(864);     // Arabic user
                _encodingCharsets[181] = Encoding.GetEncoding(864);     // Hebrew user
                _encodingCharsets[186] = Encoding.GetEncoding(775);     // Baltic
                _encodingCharsets[204] = Encoding.GetEncoding(866);     // Russian
                _encodingCharsets[222] = Encoding.GetEncoding(874);     // Thai
                _encodingCharsets[255] = Encoding.GetEncoding(437);     // OEM
            }
        }
        #endregion

        #region GetCharset
        internal static int GetCharset(Encoding encoding)
        {
            CheckEncodingCharsets();
            foreach (var key in _encodingCharsets.Keys)
            {
                if (_encodingCharsets[key] == encoding)
                {
                    return key;
                }
            }

            return 1;
        }
        #endregion

        #region GetRTFEncoding
        internal static Encoding GetRTFEncoding(int chartset)
        {
            if (chartset == 0)
                return AnsiEncoding.Instance;

            if (chartset == 1)
                return Encoding.Default;

            CheckEncodingCharsets();
            
            return _encodingCharsets.ContainsKey(chartset) ? _encodingCharsets[chartset] : null;
        }
        #endregion
    }

    #region Internal class ANSIEncoding
    /// <summary>
    ///     Internal encoding for ansi
    /// </summary>
    internal class AnsiEncoding : Encoding
    {
        public static AnsiEncoding Instance = new AnsiEncoding();

        public override string GetString(byte[] bytes, int index, int count)
        {
            var stringBuilder = new StringBuilder();
            var endIndex = Math.Min(bytes.Length - 1, index + count - 1);

            for (var iCount = index; iCount <= endIndex; iCount++)
                stringBuilder.Append(System.Convert.ToChar(bytes[iCount]));

            return stringBuilder.ToString();
        }

        public override int GetByteCount(char[] chars, int index, int count)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override int GetCharCount(byte[] bytes, int index, int count)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override int GetMaxByteCount(int charCount)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override int GetMaxCharCount(int byteCount)
        {
            throw new Exception("The method or operation is not implemented.");
        }
    }
    #endregion
}