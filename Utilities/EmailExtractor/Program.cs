//
// EmailExtractor
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
using System.IO;
using System.Linq;
using MsgReader.Outlook;

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
                return;
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
