namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf raw lay
    /// </summary>
    internal class LayerInfo
    {
        #region Fields
        private int _ucValue;
        private int _ucValueCount;
        #endregion

        #region Properties
        public int UcValue
        {
            get { return _ucValue; }
            set
            {
                _ucValue = value;
                _ucValueCount = 0;
            }
        }

        public int UCValueCount
        {
            get { return _ucValueCount; }
            set { _ucValueCount = value; }
        }
        #endregion

        #region CheckUcValueCount
        /// <summary>
        /// Checks if Uc value count is greater then zero
        /// </summary>
        /// <returns>True when greater</returns>
        public bool CheckUcValueCount()
        {
            _ucValueCount--;
            return _ucValueCount < 0;
        }
        #endregion

        #region Clone
        public LayerInfo Clone()
        {
            return (LayerInfo) MemberwiseClone();
        }
        #endregion
    }
}