using System;
using System.IO;
using System.Linq;
using MsgReader.Outlook;

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

namespace EmailExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Expected 2 command line arguments!");
                Console.WriteLine("For example: EmailExtractor.exe \"c:\\<folder to read>\" \"c:\\filetowriteto.txt\"");
                return;
            }

            var folderIn = args[0];
            if (!Directory.Exists(folderIn))
            {
                Console.WriteLine("The directory '" + folderIn + "' does not exist");
                return;;
            }

            var toFile = args[1];

            // Get all the .msg files from the folder
            var files = new DirectoryInfo(folderIn).GetFiles("*.msg");
            
            Console.WriteLine("Found '" + files.Count() + "' files to process");

            // Loop through all the files
            foreach (var file in files)
            {
                Console.WriteLine("Checking file '" + file.FullName + "'");

                // Load the msg file
                using (var message = new Storage.Message(file.FullName))
                {
                    Console.WriteLine("Found '" + message.Attachments.Count + "' attachments");

                    // Loop through all the attachments
                    foreach (var attachment in message.Attachments)
                    {
                        // Try to cast the attachment to a Message file
                        var msg = attachment as Storage.Message;

                        // If the file not is null then we have an msg file
                        if (msg != null)
                        {
                            using (msg)
                            {
                                Console.WriteLine("Found msg file '" + msg.Subject + "'");

                                if (!string.IsNullOrWhiteSpace(msg.MailingListSubscribe))
                                    Console.WriteLine("Mailing list subscribe page: '" + msg.MailingListSubscribe + "'");

                                foreach (var recipient in msg.Recipients)
                                {
                                    if (!string.IsNullOrWhiteSpace(recipient.Email))
                                    {
                                        Console.WriteLine("Recipient E-mail: '" + recipient.Email + "'");
                                        File.AppendAllText(toFile, recipient.Email + Environment.NewLine);
                                    }
                                    else if (!string.IsNullOrWhiteSpace(recipient.DisplayName))
                                    {
                                        Console.WriteLine("Recipient display name: '" + recipient.DisplayName + "'");
                                        File.AppendAllText(toFile, recipient.DisplayName + Environment.NewLine);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
