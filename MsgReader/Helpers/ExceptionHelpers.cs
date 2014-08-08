using System;

namespace MsgReader.Helpers
{
    internal static class ExceptionHelpers
    {
        #region GetInnerException
        /// <summary>
        /// Get the complete inner exception tree
        /// </summary>
        /// <param name="e">The exception object</param>
        /// <returns></returns>
        public static string GetInnerException(Exception e)
        {
            var exception = e.Message + Environment.NewLine;
            if (e.InnerException != null)
                exception += GetInnerException(e.InnerException);
            return exception;
        }
        #endregion
    }
}
