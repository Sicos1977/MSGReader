using System;
using System.IO;
using System.Linq;
using MsgReader.Outlook;
// ReSharper disable LocalizableElement

/*
   Copyright 2013-2016 Kees van Spelde

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
            //if (args.Length != 2)
            //{
            //    Console.WriteLine("Expected 2 command line arguments!");
            //    Console.WriteLine("For example: EmailExtractor.exe \"c:\\<folder to read>\" \"c:\\filetowriteto.txt\"");
            //    return;
            //}

            //var folderIn = args[0];
            //if (!Directory.Exists(folderIn))
            //{
            //    Console.WriteLine("The directory '" + folderIn + "' does not exist");
            //    return;
            //}
            var folderIn = @"d:\Progsoft\msg files\";

            var toFile = @"d:\kees.txt";

            // Get all the .msg files from the folder
            var files = new DirectoryInfo(folderIn).GetFiles("*.msg");
            var reader = new MsgReader.Reader();
            Console.WriteLine("Found '" + files.Count() + "' files to process");

            // Loop through all the files
            foreach (var file in files)
            {
                Console.WriteLine("Checking file '" + file.FullName + "'");

                // Load the msg file
                var test = reader.ExtractMsgEmailBody(File.OpenRead(file.FullName), false, string.Empty);
            }
        }
    }
}
