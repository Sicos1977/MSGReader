using System;
using System.Collections;
using System.Collections.Specialized;
using System.Globalization;

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
            get { return _infoStringDictionary["title"]; }
            set { _infoStringDictionary["title"] = value; }
        }

        /// <summary>
        /// Document subject
        /// </summary>
        public string Subject
        {
            get { return _infoStringDictionary["subject"]; }
            set { _infoStringDictionary["subject"] = value; }
        }

        /// <summary>
        /// Document author
        /// </summary>
        public string Author
        {
            get { return _infoStringDictionary["author"]; }
            set { _infoStringDictionary["author"] = value; }
        }

        /// <summary>
        /// Document manager
        /// </summary>
        public string Manager
        {
            get { return _infoStringDictionary["manager"]; }
            set { _infoStringDictionary["manager"] = value; }
        }

        /// <summary>
        /// Document company
        /// </summary>
        public string Company
        {
            get { return _infoStringDictionary["company"]; }
            set { _infoStringDictionary["company"] = value; }
        }

        /// <summary>
        /// Document operator
        /// </summary>
        public string Operator
        {
            get { return _infoStringDictionary["operator"]; }
            set { _infoStringDictionary["operator"] = value; }
        }

        /// <summary>
        /// Document category
        /// </summary>
        public string Category
        {
            get { return _infoStringDictionary["category"]; }
            set { _infoStringDictionary["categroy"] = value; }
        }

        /// <summary>
        /// Document keywords
        /// </summary>
        public string Keywords
        {
            get { return _infoStringDictionary["keywords"]; }
            set { _infoStringDictionary["keywords"] = value; }
        }

        /// <summary>
        /// Document comment
        /// </summary>
        public string Comment
        {
            get { return _infoStringDictionary["comment"]; }
            set { _infoStringDictionary["comment"] = value; }
        }

        /// <summary>
        /// Document doccomm
        /// </summary>
        public string Doccomm
        {
            get { return _infoStringDictionary["doccomm"]; }
            set { _infoStringDictionary["doccomm"] = value; }
        }

        /// <summary>
        /// Document base
        /// </summary>
        public string HLinkbase
        {
            get { return _infoStringDictionary["hlinkbase"]; }
            set { _infoStringDictionary["hlinkbase"] = value; }
        }

        /// <summary>
        /// Total edit minutes
        /// </summary>
        public int EditMinutes
        {
            get
            {
                if (_infoStringDictionary.ContainsKey("edmins"))
                {
                    var v = Convert.ToString(_infoStringDictionary["edmins"]);
                    int result;
                    if (int.TryParse(v, out result))
                        return result;
                }
                return 0;
            }
            set { _infoStringDictionary["edmins"] = value.ToString(CultureInfo.InvariantCulture); }
        }

        /// <summary>
        /// Document version
        /// </summary>
        public string Version
        {
            get { return _infoStringDictionary["vern"]; }
            set { _infoStringDictionary["vern"] = value; }
        }

        /// <summary>
        /// Document number of pages
        /// </summary>
        public string NumberOfPages
        {
            get { return _infoStringDictionary["nofpages"]; }
            set { _infoStringDictionary["nofpages"] = value; }
        }

        /// <summary>
        /// Document number of words
        /// </summary>
        public string NumberOfWords
        {
            get { return _infoStringDictionary["nofwords"]; }
            set { _infoStringDictionary["nofwords"] = value; }
        }

        /// <summary>
        /// Document number of characters , include whitespace
        /// </summary>
        public string NumberOfCharaktersWithWhiteSpace
        {
            get { return _infoStringDictionary["nofchars"]; }
            set { _infoStringDictionary["nofchars"] = value; }
        }

        /// <summary>
        /// Document number of characters , exclude white space
        /// </summary>
        public string NumberOfCharaktersWithoutWhiteSpace
        {
            get { return _infoStringDictionary["nofcharsws"]; }
            set { _infoStringDictionary["nofcharsws"] = value; }
        }

        /// <summary>
        /// Document inner id
        /// </summary>
        public string Id
        {
            get { return _infoStringDictionary["id"]; }
            set { _infoStringDictionary["id"] = value; }
        }

        /// <summary>
        /// Document Creation time
        /// </summary>
        public DateTime CreationTime
        {
            get { return _creationTime; }
            set { _creationTime = value; }
        }

        /// <summary>
        /// Document modified time
        /// </summary>
        public DateTime RevisionTime
        {
            get { return _revisionTime; }
            set { _revisionTime = value; }
        }

        /// <summary>
        /// Document last print time
        /// </summary>
        public DateTime PrintTime
        {
            get { return _printTime; }
            set { _printTime = value; }
        }

        /// <summary>
        /// Document last backup time
        /// </summary>
        public DateTime BackupTime
        {
            get { return _backupTime; }
            set { _backupTime = value; }
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

        #region Write
        public void Write(Writer writer)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword("info");
            foreach (string strKey in _infoStringDictionary.Keys)
            {
                writer.WriteStartGroup();
                if (strKey == "edmins"
                    || strKey == "vern"
                    || strKey == "nofpages"
                    || strKey == "nofwords"
                    || strKey == "nofchars"
                    || strKey == "nofcharsws"
                    || strKey == "id")
                {
                    writer.WriteKeyword(strKey + _infoStringDictionary[strKey]);
                }
                else
                {
                    writer.WriteKeyword(strKey);
                    writer.WriteText(_infoStringDictionary[strKey]);
                }
                writer.WriteEndGroup();
            }
            writer.WriteStartGroup();

            WriteTime(writer, "creatim", _creationTime);
            WriteTime(writer, "revtim", _revisionTime);
            WriteTime(writer, "printim", _printTime);
            WriteTime(writer, "buptim", _backupTime);

            writer.WriteEndGroup();
        }
        #endregion

        #region WriteTime
        private void WriteTime(Writer writer, string name, DateTime value)
        {
            writer.WriteStartGroup();
            writer.WriteKeyword(name);
            writer.WriteKeyword("yr" + value.Year);
            writer.WriteKeyword("mo" + value.Month);
            writer.WriteKeyword("dy" + value.Day);
            writer.WriteKeyword("hr" + value.Hour);
            writer.WriteKeyword("min" + value.Minute);
            writer.WriteKeyword("sec" + value.Second);
            writer.WriteEndGroup();
        }
        #endregion
    }
}