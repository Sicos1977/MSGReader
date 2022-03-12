//
// DocumentInfo.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2022 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Collections;
using System.Collections.Specialized;
using System.Globalization;
// ReSharper disable UnusedMember.Global

namespace MsgReader.Rtf
{
    /// <summary>
    /// Document information
    /// </summary>
    internal class DocumentInfo
    {
        #region Fields
        private readonly StringDictionary _infoStringDictionary = new StringDictionary();
        private DateTime _backupTime = DateTime.Now;
        private DateTime _creationTime = DateTime.Now;
        private DateTime _printTime = DateTime.Now;
        private DateTime _revisionTime = DateTime.Now;
        #endregion

        #region Properties
        /// <summary>
        /// Document title
        /// </summary>
        public string Title
        {
            get => _infoStringDictionary["title"];
            set => _infoStringDictionary["title"] = value;
        }

        /// <summary>
        /// Document subject
        /// </summary>
        public string Subject
        {
            get => _infoStringDictionary["subject"];
            set => _infoStringDictionary["subject"] = value;
        }

        /// <summary>
        /// Document author
        /// </summary>
        public string Author
        {
            get => _infoStringDictionary["author"];
            set => _infoStringDictionary["author"] = value;
        }

        /// <summary>
        /// Document manager
        /// </summary>
        public string Manager
        {
            get => _infoStringDictionary["manager"];
            set => _infoStringDictionary["manager"] = value;
        }

        /// <summary>
        /// Document company
        /// </summary>
        public string Company
        {
            get => _infoStringDictionary["company"];
            set => _infoStringDictionary["company"] = value;
        }

        /// <summary>
        /// Document operator
        /// </summary>
        public string Operator
        {
            get => _infoStringDictionary["operator"];
            set => _infoStringDictionary["operator"] = value;
        }

        /// <summary>
        /// Document category
        /// </summary>
        public string Category
        {
            get => _infoStringDictionary["category"];
            set => _infoStringDictionary["categroy"] = value;
        }

        /// <summary>
        /// Document keywords
        /// </summary>
        public string Keywords
        {
            get => _infoStringDictionary["keywords"];
            set => _infoStringDictionary["keywords"] = value;
        }

        /// <summary>
        /// Document comment
        /// </summary>
        public string Comment
        {
            get => _infoStringDictionary["comment"];
            set => _infoStringDictionary["comment"] = value;
        }

        /// <summary>
        /// Document doccomm
        /// </summary>
        public string Doccomm
        {
            get => _infoStringDictionary["doccomm"];
            set => _infoStringDictionary["doccomm"] = value;
        }

        /// <summary>
        /// Document base
        /// </summary>
        public string HLinkbase
        {
            get => _infoStringDictionary["hlinkbase"];
            set => _infoStringDictionary["hlinkbase"] = value;
        }

        /// <summary>
        /// Total edit minutes
        /// </summary>
        public int EditMinutes
        {
            get
            {
                if (!_infoStringDictionary.ContainsKey("edmins")) return 0;
                var v = Convert.ToString(_infoStringDictionary["edmins"]);
                return int.TryParse(v, out var result) ? result : 0;
            }
            set => _infoStringDictionary["edmins"] = value.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Document version
        /// </summary>
        public string Version
        {
            get => _infoStringDictionary["vern"];
            set => _infoStringDictionary["vern"] = value;
        }

        /// <summary>
        /// Document number of pages
        /// </summary>
        public string NumberOfPages
        {
            get => _infoStringDictionary["nofpages"];
            set => _infoStringDictionary["nofpages"] = value;
        }

        /// <summary>
        /// Document number of words
        /// </summary>
        public string NumberOfWords
        {
            get => _infoStringDictionary["nofwords"];
            set => _infoStringDictionary["nofwords"] = value;
        }

        /// <summary>
        /// Document number of characters , include whitespace
        /// </summary>
        public string NumberOfCharaktersWithWhiteSpace
        {
            get => _infoStringDictionary["nofchars"];
            set => _infoStringDictionary["nofchars"] = value;
        }

        /// <summary>
        /// Document number of characters , exclude white space
        /// </summary>
        public string NumberOfCharaktersWithoutWhiteSpace
        {
            get => _infoStringDictionary["nofcharsws"];
            set => _infoStringDictionary["nofcharsws"] = value;
        }

        /// <summary>
        /// Document inner id
        /// </summary>
        public string Id
        {
            get => _infoStringDictionary["id"];
            set => _infoStringDictionary["id"] = value;
        }

        /// <summary>
        /// Document Creation time
        /// </summary>
        public DateTime CreationTime
        {
            get => _creationTime;
            set => _creationTime = value;
        }

        /// <summary>
        /// Document modified time
        /// </summary>
        public DateTime RevisionTime
        {
            get => _revisionTime;
            set => _revisionTime = value;
        }

        /// <summary>
        /// Document last print time
        /// </summary>
        public DateTime PrintTime
        {
            get => _printTime;
            set => _printTime = value;
        }

        /// <summary>
        /// Document last backup time
        /// </summary>
        public DateTime BackupTime
        {
            get => _backupTime;
            set => _backupTime = value;
        }
        #endregion

        #region StringItems
        internal string[] StringItems
        {
            get
            {
                var list = new ArrayList();
                foreach (string key in _infoStringDictionary.Keys)
                {
                    list.Add(key + "=" + _infoStringDictionary[key]);
                }
                list.Add("Creatim=" + CreationTime.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Revtim=" + RevisionTime.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Printim=" + PrintTime.ToString("yyyy-MM-dd HH:mm:ss"));
                list.Add("Buptim=" + BackupTime.ToString("yyyy-MM-dd HH:mm:ss"));
                return (string[]) list.ToArray(typeof (string));
            }
        }
        #endregion

        #region GetInfo
        /// <summary>
        /// get information specify name
        /// </summary>
        /// <param name="strName">name</param>
        /// <returns>value</returns>
        public string GetInfo(string strName)
        {
            return _infoStringDictionary[strName];
        }
        #endregion

        #region SetInfo
        /// <summary>
        /// set information specify name
        /// </summary>
        /// <param name="strName">name</param>
        /// <param name="strValue">value</param>
        public void SetInfo(string strName, string strValue)
        {
            _infoStringDictionary[strName] = strValue;
        }
        #endregion

        #region Clear
        public void Clear()
        {
            _infoStringDictionary.Clear();
            _creationTime = DateTime.Now;
            _revisionTime = DateTime.Now;
            _printTime = DateTime.Now;
            _backupTime = DateTime.Now;
        }
        #endregion
    }
}