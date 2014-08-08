using System.Collections;
using System.Drawing;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf color table
    /// </summary>
    internal class ColorTable
    {
        #region Fields
        private readonly ArrayList _arrayList = new ArrayList();
        #endregion

        #region Constructor
        public ColorTable()
        {
            CheckValueExistWhenAdd = true;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Get color at special index
        /// </summary>
        public Color this[int index]
        {
            get { return (Color) _arrayList[index]; }
        }

        /// <summary>
        /// Check color value exist when add color to list
        /// </summary>
        public bool CheckValueExistWhenAdd { get; set; }

        /// <summary>
        /// Count the items in the color table
        /// </summary>
        public int Count
        {
            get { return _arrayList.Count; }
        }
        #endregion

        #region GetColor
        /// <summary>
        /// Get color at special index , if index out of range , return default color
        /// </summary>
        /// <param name="index">index</param>
        /// <param name="defaultValue">default value</param>
        /// <returns>color value</returns>
        public Color GetColor(int index, Color defaultValue)
        {
            index --;
            if (index >= 0 && index < _arrayList.Count)
            {
                return (Color) _arrayList[index];
            }
            return defaultValue;
        }
        #endregion

        #region Add
        /// <summary>
        /// Add color to list
        /// </summary>
        /// <param name="c">New color value</param>
        public void Add(Color c)
        {
            if (c.IsEmpty)
                return;

            if (c.A == 0)
                return;

            if (c.A != 255)
                c = Color.FromArgb(255, c);

            if (CheckValueExistWhenAdd)
            {
                if (IndexOf(c) < 0)
                    _arrayList.Add(c);
            }
            else
                _arrayList.Add(c);
        }
        #endregion

        #region Remove
        /// <summary>
        /// Remove special color
        /// </summary>
        /// <param name="c">color value</param>
        public void Remove(Color c)
        {
            var index = IndexOf(c);
            if (index >= 0)
                _arrayList.RemoveAt(index);
        }
        #endregion

        #region IndexOf
        /// <summary>
        /// Get color index
        /// </summary>
        /// <param name="c">color</param>
        /// <returns>index , if not found return -1</returns>
        public int IndexOf(Color c)
        {
            if (c.A == 0)
                return -1;

            if (c.A != 255)
                c = Color.FromArgb(255, c);

            for (var count = 0; count < _arrayList.Count; count++)
            {
                var color = (Color) _arrayList[count];
                if (color.ToArgb() == c.ToArgb())
                    return count;
            }
            return -1;
        }
        #endregion

        #region Clear
        /// <summary>
        /// Clear
        /// </summary>
        public void Clear()
        {
            _arrayList.Clear();
        }
        #endregion

        #region Write
        /// <summary>
        /// Write
        /// </summary>
        /// <param name="writer"></param>
        public void Write(Writer writer)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword(Consts.Colortbl);
            writer.WriteRaw(";");
            foreach (var item in _arrayList)
            {
                var c = (Color) item;
                writer.WriteKeyword("red" + c.R);
                writer.WriteKeyword("green" + c.G);
                writer.WriteKeyword("blue" + c.B);
                writer.WriteRaw(";");
            }
            writer.WriteEndGroup();
        }
        #endregion

        #region Clone
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns></returns>
        public ColorTable Clone()
        {
            var colorTable = new ColorTable();
            for (var count = 0; count < _arrayList.Count; count++)
            {
                var c = (Color) _arrayList[count];
                colorTable._arrayList.Add(c);
            }
            return colorTable;
        }
        #endregion
    }
}