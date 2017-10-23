using Itenso.Rtf;
using Itenso.Rtf.Converter.Html;
using Itenso.Rtf.Converter.Image;
using Itenso.Rtf.Support;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
{
    /// <summary>
    /// This class is used to convert RTF to HTML by using the RichTextEditBox
    /// </summary>
    internal class RtfToHtmlConverter
    {
        #region Fields
        /// <summary>
        /// The RTF string that needs to be converted
        /// </summary>
        private string _rtf;

        /// <summary>
        /// The RTF string that is converted to HTML
        /// </summary>
        private string _convertedRtf;
        #endregion

        #region ConvertRtfToHtml
        /// <summary>
        /// Convert RTF to HTML by using the Windows RichTextBox
        /// </summary>
        /// <param name="rtf">The rtf string</param>
        /// <returns></returns>
        public string ConvertRtfToHtml(string rtf)
        {
            _rtf = rtf;
            Convert();
            return _convertedRtf;
        }
        #endregion

        

        #region Convert
        /// <summary>
        /// Do the actual conversion by using a RichTextBox
        /// </summary>
        private void Convert()
        {
            if (string.IsNullOrEmpty(_rtf))
            {
                _convertedRtf = string.Empty;
                return;
            }
            
            IRtfDocument rtfDocument = RtfInterpreterTool.BuildDoc(_rtf);
            RtfHtmlConverter converter = new RtfHtmlConverter(rtfDocument);

            _convertedRtf = converter.Convert();
        }
        #endregion
        
   
        
    }
}
