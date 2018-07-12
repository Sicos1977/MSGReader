using System.Collections;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Table row
    /// </summary>
    internal class DomTableRow : DomElement
    {
        #region Properties
        /// <summary>
        /// Cell settings
        /// </summary>
        internal ArrayList CellSettings { get; set; }

        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }

        /// <summary>
        /// Document level
        /// </summary>
        public int Level { get; set; }

        /// <summary>
        /// row index
        /// </summary>
        internal int RowIndex { get; set; }

        /// <summary>
        /// Is the last row
        /// </summary>
        public bool IsLastRow { get; set; }

        /// <summary>
        /// Is header row
        /// </summary>
        public bool Header { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }
        
        /// <summary>
        /// Padding left
        /// </summary>
        public int PaddingLeft { get; set; }

        /// <summary>
        /// Top padding
        /// </summary>
        public int PaddingTop { get; set; }


        /// <summary>
        /// Right padding
        /// </summary>
        public int PaddingRight { get; set; }

        /// <summary>
        /// Bottom padding
        /// </summary>
        public int PaddingBottom { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }
        #endregion

        #region Constructor
        public DomTableRow()
        {
            Width = 0;
            PaddingBottom = int.MinValue;
            PaddingRight = int.MinValue;
            PaddingTop = int.MinValue;
            PaddingLeft = int.MinValue;
            Height = 0;
            Header = false;
            IsLastRow = false;
            RowIndex = 0;
            Level = 1;
            Format = new DocumentFormatInfo();
            CellSettings = new ArrayList();
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "Row " + RowIndex;
        }
        #endregion
    }
}
