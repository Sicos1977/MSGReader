using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

/*
   Copyright 2013-2015 Kees van Spelde

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

namespace MsgMapper
{
    [Serializable]
    public class Folders : List<Folder>
    {
        #region Add
        /// <summary>
        /// Adds a new folder
        /// </summary>
        /// <param name="path"></param>
        public void Add(string path)
        {
            Add(new Folder {Path = path }); 
        }
        #endregion

        #region LoadFromString
        /// <summary>
        /// Create this object from an XML string
        /// </summary>
        /// <param name="xml">The XML string</param>
        public static Folders LoadFromString(string xml)
        {
            try
            {
                var xmlSerializer = new XmlSerializer(typeof(Folders));
                var rdr = new StringReader(xml);
                return (Folders)xmlSerializer.Deserialize(rdr);
            }
            catch (Exception e)
            {
                throw new Exception("The XML string contains invalid XML, error: " + e.Message);
            }
        }
        #endregion
        
        #region SerializeToString
        /// <summary>
        /// Serialize this object to a string
        /// </summary>
        /// <returns>The object as an XML string</returns>
        public string SerializeToString()
        {
            var stringWriter = new StringWriter(new StringBuilder());
            var s = new XmlSerializer(GetType());

            s.Serialize(stringWriter, this);

            var output = stringWriter.ToString();
            output = output.Replace("<?xml version=\"1.0\" encoding=\"utf-8\"?>", string.Empty);
            output = output.Replace("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"", string.Empty);
            output = output.Replace("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"", string.Empty);

            return output;
        }
        #endregion
    }

    public class Folder
    {
        public string Path { get; set; }
    }
}
