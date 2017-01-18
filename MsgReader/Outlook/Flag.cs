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
        /// Class used to contain all the flag (follow up) information of a <see cref="Storage.Message"/>.
        /// </summary>
        public sealed class Flag : Storage
        {
            #region Public enum FlagStatus
            /// <summary>
            /// The flag state of an message
            /// </summary>
            public enum FlagStatus
            {
                /// <summary>
                /// The msg object has been flagged as completed
                /// </summary>
                Complete = 1,

                /// <summary>
                /// The msg object has been flagged and marked as a task
                /// </summary>
                Marked = 2
            }
            #endregion

            #region Properties
            /// <summary>
            /// Returns the flag request text
            /// </summary>
            public string Request { get; private set; }

            /// <summary>
            /// Returns the <see cref="FlagStatus">Status</see> of the flag
            /// </summary>
            public FlagStatus? Status { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Flag" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Flag(Storage message) : base(message._storage)
            {
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

                Request = GetMapiPropertyString(MapiTags.FlagRequest);

                var status = GetMapiPropertyInt32(MapiTags.PR_FLAG_STATUS);
                if (status == null)
                    Status = null;
                else
                    Status = (FlagStatus)status; 
            }
            #endregion
        }
    }
}