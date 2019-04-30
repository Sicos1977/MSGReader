//
// Folders.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
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
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

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
