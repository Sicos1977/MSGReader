using System;

namespace DocumentServices.Modules.Readers.MsgReader.Rtf
{
    #region Delegates
    /// <summary>
    /// Progress event handler type
    /// </summary>
    /// <param name="sender">sender</param>
    /// <param name="args">event arguments</param>
    public delegate void ProgressEventHandler( object sender , ProgressEventArgs args );
    #endregion

    /// <summary>
    /// Progress event arguments
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        #region Properties
        /// <summary>
        /// Progress max value
        /// </summary>
        public int MaxValue { get; private set; }

        /// <summary>
        /// Current value
        /// </summary>
        public int Value { get; private set; }

        /// <summary>
        /// Progress message
        /// </summary>
        public string Message { get; private set; }

        /// <summary>
        /// Cancel operation
        /// </summary>
        public bool Cancel { get; set; }
        #endregion

        #region Constructor
        public ProgressEventArgs(int max, int value, string message)
        {
            Cancel = false;
            MaxValue = max;
            Value = value;
            Message = message;
        }
        #endregion
    }
}
