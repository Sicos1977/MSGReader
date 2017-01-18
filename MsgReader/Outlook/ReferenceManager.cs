using System.Collections.Generic;
using System.Runtime.InteropServices;

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
    public partial class Storage
    {
        /// <summary>
        /// Class used to track all the COM objects
        /// </summary>
        private class ReferenceManager
        {
            #region Fields
            /// <summary>
            /// Static instance to the <see cref="ReferenceManager"/>
            /// </summary>
            private static readonly ReferenceManager Instance = new ReferenceManager();

            /// <summary>
            /// Used to keep track of all the objects
            /// </summary>
            private readonly List<object> _trackingObjects = new List<object>();
            #endregion

            #region Destructor
            /// <summary>
            /// Dispose all tracked objects
            /// </summary>
            ~ReferenceManager()
            {
                foreach (var trackingObject in _trackingObjects)
                    Marshal.ReleaseComObject(trackingObject);
            }
            #endregion

            #region AddItem
            /// <summary>
            /// Add an object to track
            /// </summary>
            /// <param name="track"></param>
            public static void AddItem(object track)
            {
                lock (Instance)
                {
                    if (!Instance._trackingObjects.Contains(track))
                        Instance._trackingObjects.Add(track);
                }
            }
            #endregion

            #region RemoveItem
            /// <summary>
            /// Remove a tracked object
            /// </summary>
            /// <param name="track"></param>
            public static void RemoveItem(object track)
            {
                lock (Instance)
                {
                    if (Instance._trackingObjects.Contains(track))
                        Instance._trackingObjects.Remove(track);
                }
            }
            #endregion
        }
    }
}