using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

namespace MsgReader.Helpers
{
    /// <summary>
    /// This class is used as a placeholder for the filetype information
    /// </summary>
    internal class FileTypeFileInfo
    {
        #region Fields
        private byte[] _magicBytes;
        #endregion

        #region Properties
        /// <summary>
        /// The magic bytes
        /// </summary>
        public byte[] MagicBytes
        {
            get { return _magicBytes; }
            set
            {
                _magicBytes = value;
                if (_magicBytes != null)
                    MagicBytesAsString = Encoding.ASCII.GetString(_magicBytes);
            }
        }

        /// <summary>
        /// The magic bytes as a string
        /// </summary>
        public string MagicBytesAsString { get; internal set; }

        /// <summary>
        /// The file type extension that belongs to the magic bytes (e.g. doc, xls, msg)
        /// </summary>
        public string Extension { get; set; }

        /// <summary>
        /// Description of the file
        /// </summary>
        public string Description { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="magicBytes">The magic bytes</param>
        /// <param name="extension">The file type extension that belongs to the magic bytes (e.g. doc, xls, msg)</param>
        /// <param name="description">Description of the file (e.g. Word document, Outlook message, etc...)</param>
        internal FileTypeFileInfo(byte[] magicBytes, string extension, string description)
        {
            MagicBytes = magicBytes;
            Extension = extension;
            Description = description;
        }
        #endregion
    }

    /// <summary>
    /// This class can be used to recognize files by their magic bytes
    /// </summary>
    internal class FileTypeSelector
    {
        #region Consts
        private const string MicroSoftOffice = "MICROSOFTOFFICE";
        private const string ZipOrOffice2007 = "ZIPOROFFICE2007";
        private const string ByteOrderMarkerFile = "BYTEORDERMARKERFILE";
        #endregion

        #region Fields
        /// <summary>
        /// Contains all the magic bytes and their description
        /// </summary>
        private static readonly List<FileTypeFileInfo> FileTypes = GetFileTypes();
        #endregion

        #region STB
        /// <summary>
        /// Converts a string to a byte array
        /// </summary>
        /// <param name="line">String to convert</param>
        /// <returns>Byte array</returns>
        private static Byte[] Stb(string line)
        {
            return Encoding.ASCII.GetBytes(line);
        }
        #endregion

        #region GetFileTypes
        /// <summary>
        /// A list with most of the used file types and their magic bytes
        /// </summary>
        /// <returns></returns>
        private static List<FileTypeFileInfo> GetFileTypes()
        {
            // ReSharper disable UseObjectOrCollectionInitializer
            var fileTypes = new List<FileTypeFileInfo>();
            // ReSharper restore UseObjectOrCollectionInitializer

            // Microsoft binary formats
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, MicroSoftOffice, "Microsoft Office applications (Word, Powerpoint, Excel, Works)"));

            // Microsoft open document file format or zip
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x50, 0x4B }, ZipOrOffice2007, "Zip or Microsoft Office 2007, 2010 or 2013 document"));

            // PDF
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.7"), "pdf", "Adobe Portable Document file (version 1.7)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.6"), "pdf", "Adobe Portable Document file (version 1.6)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.5"), "pdf", "Adobe Portable Document file (version 1.5)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.4"), "pdf", "Adobe Portable Document file (version 1.4)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.3"), "pdf", "Adobe Portable Document file (version 1.3)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.2"), "pdf", "Adobe Portable Document file (version 1.2)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.1"), "pdf", "Adobe Portable Document file (version 1.1)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF-1.0"), "pdf", "Adobe Portable Document file (version 1.0)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("%PDF"), "pdf", "Adobe Portable Document file"));

            // RTF
            fileTypes.Add(new FileTypeFileInfo(Stb("{\\rtf1"), "rtf", "Rich Text Format"));

            // FileNet COLD document
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xC5, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0D }, "cold", "FileNet COLD document"));

            // Developing
            fileTypes.Add(new FileTypeFileInfo(Stb("# Microsoft Developer Studio"), "dsp", "Microsoft Developer Studio project file"));
            fileTypes.Add(new FileTypeFileInfo(Stb("dswfile"), "dsp", "Microsoft Visual Studio workspace file"));
            fileTypes.Add(new FileTypeFileInfo(Stb("#!/usr/bin/perl"), "pl", "Perl script file"));

            // Corel Paint Shop Pro
            fileTypes.Add(new FileTypeFileInfo(Stb("Paint Shop Pro Image File"), "pspimage", "Corel Paint Shop Pro Image file"));
            fileTypes.Add(new FileTypeFileInfo(Stb("JASC BROWS FILE"), "jbf", "Corel Paint Shop Pro browse file"));

            // Microsoft Access
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x01, 0x00, 0x00, 0x53, 0x74, 0x61, 0x6E, 0x64, 0x61, 0x72, 0x64, 0x20, 0x4A, 0x65, 0x74, 0x20, 0x44, 0x42 }, "mdb", "Microsoft Access file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x01, 0x00, 0x00, 0x53, 0x74, 0x61, 0x6E, 0x64, 0x61, 0x72, 0x64, 0x20, 0x41, 0x43, 0x45, 0x20, 0x44, 0x42 }, "accdb", "Microsoft Access 2007 file"));

            // Microsoft Outlook
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x9C, 0xCB, 0xCB, 0x8D, 0x13, 0x75, 0xD2, 0x11, 0x91, 0x58, 0x00, 0xC0, 0x4F, 0x79, 0x56, 0xA4 }, "wab", "Outlook address file"));
            fileTypes.Add(new FileTypeFileInfo(Stb("!BD"), "pst", "Microsoft  Outlook Personal Folder File"));

            // ZIP
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x01, 0x00 }, "zip", "ZLock Pro encrypted ZIP"));
            fileTypes.Add(new FileTypeFileInfo(Stb("WinZip"), "zip", "WinZip compressed archive"));
            fileTypes.Add(new FileTypeFileInfo(Stb("PKLITE"), "zip", "PKLITE compressed ZIP archive (see also PKZIP)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x37, 0x7A, 0xBC, 0xAF, 0x27, 0x1C }, "7z", "7-Zip compressed file")); // 7Z zip formaat	
            fileTypes.Add(new FileTypeFileInfo(Stb("PKSFX"), "zip", "PKSFX self-extracting executable compressed file (see also PKZIP)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1F, 0x8B, 0x08 }, "gz", "GZIP archive file"));

            // XML
            fileTypes.Add(new FileTypeFileInfo(Stb("<?xml version=\"1.0\"?>"), "xml", "XML File (UTF16 encoding)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("<?xml version=\"1.0\" encoding=\"utf-16\""), "xml", "XML File (UTF16 encoding)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("<?xml version=\"1.0\" encoding=\"utf-8\""), "xml", "XML File (UTF8 encoding)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("<?xml version=\"1.0\" encoding=\"utf-7\""), "xml", "XML File (UTF7 encoding)"));

            // EML
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x52, 0x65, 0x74, 0x75, 0x72, 0x6E, 0x2D, 0x50, 0x61, 0x74, 0x68, 0x3A, 0x20 }, "eml", "A commmon file extension for e-mail files"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x72, 0x6F, 0x6D, 0x20, 0x3F, 0x3F, 0x3F }, "eml", "E-mail markup language file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x72, 0x6F, 0x6D, 0x20, 0x20, 0x20 }, "eml", "E-mail markup language file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x72, 0x6F, 0x6D, 0x3A, 0x20 }, "eml", "E-mail markup language file"));

            // TIF
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4D, 0x4D, 0x00, 0x2B }, "tif", "BigTIFF files; Tagged Image File Format files > 4 GB"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4D, 0x4D, 0x00, 0x2A }, "tif", "Tagged Image File Format file (big endian, i.e., LSB last in the byte; Motorola)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x49, 0x49, 0x2A, 0x00 }, "tif", "Tagged Image File Format file (little endian, i.e., LSB first in the byte; Intel)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("I I"), "tif", "Tagged Image File Format file"));

            // AutoCAD
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x30, 0x32 }, "dwg", "Generic AutoCAD drawing - AutoCAD R2.5"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x30, 0x33 }, "dwg", "Generic AutoCAD drawing - AutoCAD R2.6"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x30, 0x34 }, "dwg", "Generic AutoCAD drawing - AutoCAD R9"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x30, 0x36 }, "dwg", "Generic AutoCAD drawing - AutoCAD R10"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x30, 0x39 }, "dwg", "Generic AutoCAD drawing - AutoCAD R11/R12"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x30 }, "dwg", "Generic AutoCAD drawing - AutoCAD R13 (subtype 10)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x31 }, "dwg", "Generic AutoCAD drawing - AutoCAD R13 (subtype 11)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x32 }, "dwg", "Generic AutoCAD drawing - AutoCAD R13 (subtype 12)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x33 }, "dwg", "Generic AutoCAD drawing - AutoCAD R13 (subtype 13)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x34 }, "dwg", "Generic AutoCAD drawing - AutoCAD R13 (subtype 14)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x35 }, "dwg", "Generic AutoCAD drawing - AutoCAD R2000"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x31, 0x38 }, "dwg", "Generic AutoCAD drawing - AutoCAD R2004"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x41, 0x43, 0x31, 0x30, 0x32, 0x31 }, "dwg", "Generic AutoCAD drawing - AutoCAD R2007"));

            // JPG
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xDB }, "jpg", "Samsung D807 JPEG file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 }, "jpg", "JPEG/JIFF file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xE1 }, "jpg", "JPEG/Exif file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xE2 }, "jpg", "Canon EOS-1D JPEG file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xE3 }, "jpg", "Samsung D500 JPEG file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0xFF, 0xD8, 0xFF, 0xE8 }, "jpg", "Still Picture Interchange File Format (SPIFF)"));

            // PNG
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "png", "Portable Network Graphics"));

            // RealAudio
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x2E, 0x52, 0x4D, 0x46, 0x00, 0x00, 0x00, 0x12, 0x00 }, "ra", "RealAudio file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x2E, 0x72, 0x61, 0xFD, 0x00 }, "ra", "RealAudio streaming media file"));
            fileTypes.Add(new FileTypeFileInfo(Stb(".REC"), "ivr", "RealPlayer video file (V11 and later)"));
            fileTypes.Add(new FileTypeFileInfo(Stb(".RMF"), "rm", "RealMedia streaming media file"));

            // MP3
            fileTypes.Add(new FileTypeFileInfo(Stb("ID3"), "mp3", "MPEG-1 Audio Layer 3 (MP3) audio file"));

            // IMG
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x01, 0x00, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01 }, "img", "Image Format Bitmap file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x50, 0x49, 0x43, 0x54, 0x00, 0x08 }, "img", "ADEX Corp. ChromaGraph Graphics Card Bitmap Graphic file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x53, 0x43, 0x4D, 0x49 }, "img", "Img Software Set Bitmap"));

            // GIF
            fileTypes.Add(new FileTypeFileInfo(Stb("GIF87a"), "gif", "Graphics interchange format file (GIF87a)"));
            fileTypes.Add(new FileTypeFileInfo(Stb("GIF89a"), "gif", "Graphics interchange format file (GIF89a)"));

            // BMP
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x42, 0x4D }, "BMP", "Windows (or device-independent) bitmap image"));

            // MDI
            fileTypes.Add(new FileTypeFileInfo(Stb("MThd"), "mdi", "Musical Instrument Digital Interface (MIDI) sound file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x45, 0x50 }, "mdi", "Microsoft Document Imaging file"));

            // WRI
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x32, 0xBE }, "wri", "Microsoft Write file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x31, 0xBE }, "wri", "Microsoft Write file"));

            // ARC
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x02 }, "arc", "LH archive file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x03 }, "arc", "LH archive file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x04 }, "arc", "LH archive file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x08 }, "arc", "LH archive file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x09 }, "arc", "LH archive file"));

            // Windows Event viewer 
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x45, 0x6C, 0x66, 0x46, 0x69, 0x6C, 0x65, 0x00 }, "evtx", "Windows Vista event log file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x30, 0x00, 0x00, 0x00, 0x4C, 0x66, 0x4C, 0x65 }, "evt", "Windows Event Viewer file"));

            // Microsoft Help File
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF }, "hlp", "Windows help file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4C, 0x4E, 0x02, 0x00 }, "hlp", "Windows Help file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x3F, 0x5F, 0x03, 0x00 }, "hlp", "Windows help file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x49, 0x54, 0x53, 0x46 }, "chm", "Microsoft Compiled HTM   L Help File"));

            // SWF
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x57, 0x53 }, "swf", "Macromedia Shockwave Flash player file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x43, 0x57, 0x53 }, "swf", "Shockwave Flash file (v5+)"));

            // CAB
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4D, 0x53, 0x43, 0x46 }, "cab", "Microsoft cabinet file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x49, 0x53, 0x63, 0x28 }, "cab", "Install Shield v5.x or 6.x compressed file"));

            // vCard
            fileTypes.Add(new FileTypeFileInfo(Stb("BEGIN:VCARD"), "vcf", "vCard file"));

            // RAR
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x52, 0x61, 0x72, 0x21, 0x1A, 0x07, 0x00 }, "rar", "WinRAR compressed archive file"));

            // LHA
            fileTypes.Add(new FileTypeFileInfo(Stb("-lh"), "lha", "Compressed archive file"));

            // Adobe PhotoShop
            fileTypes.Add(new FileTypeFileInfo(Stb("8BPS"), "psd", "Photoshop image file"));

            // Apple Quicktime
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x0B, 0x77 }, "ac3", "Apple QuickTime movie"));

            // MKV
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x1A, 0x45, 0xD5, 0xA3, 0x93, 0x42, 0x82, 0x88, 0x6D, 0x61, 0x74, 0x72, 0x6F, 0x73, 0x6B }, "mkv", "Matroska open movie format"));

            // AVI RIFF
            fileTypes.Add(new FileTypeFileInfo(Stb("RIFF"), "avi", "Audio Video Interleave"));

            // Windows Media
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x30, 0x26, 0xB2, 0x75, 0x8E, 0x66, 0xCF, 0x11, 0xA6, 0xD9, 0x00, 0xAA, 0x00, 0x62, 0xCE, 0x6C }, "wmv", "Microsoft Windows Media Audio/Video File (Advanced Streaming Format"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x30, 0x26, 0xB2, 0x75, 0x8E, 0x66, 0xCF, 0x11, 0xA6, 0xD9, 0x00, 0xAA, 0x00, 0x62, 0xCE, 0x6C }, "wma", "	Microsoft Windows Media Audio/Video File (Advanced Streaming Format)"));

            // NT Backup
            fileTypes.Add(new FileTypeFileInfo(Stb("TAPE"), "bkf", "Windows NT Backup file (NTBackup)"));

            // Windows registry file
            fileTypes.Add(new FileTypeFileInfo(Stb("Windows Registry Editor Version 5.00"), "reg", "Windows Registry Editor Version 5.00 file"));

            // Others
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x33, 0x67, 0x70, 0x35 }, "mp4", "MPEG-4 video files"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x49, 0x49, 0x1A, 0x00, 0x00, 0x00, 0x48, 0x45, 0x41, 0x50, 0x43, 0x43, 0x44, 0x52, 0x02, 0x00 }, "crw", "Canon digital camera RAW file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0x00, 0x20, 0x66, 0x74, 0x79, 0x70, 0x4D, 0x34, 0x41, 0x20, 0x00, 0x00, 0x00, 0x00 }, "mov", "Apple QuickTime movie file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4C, 0x00, 0x00, 0x00, 0x01, 0x14, 0x02, 0x00 }, "lnk", "Windows shortcut file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x52, 0x45, 0x47, 0x45, 0x44, 0x49, 0x54 }, "reg", "Windows NT Registry and Registry Undo files"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x43, 0x50, 0x54, 0x46, 0x49, 0x4C, 0x45 }, "cpt", "Corel Photopaint file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4A, 0x41, 0x52, 0x43, 0x53, 0x00 }, "jar", "JARCS compressed archive"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x4F, 0x52, 0x4D, 0x00 }, "aiff", "Audio Interchange File"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x4B, 0x49, 0x00, 0x00 }, "shd", "Windows 9x printer spool file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x46, 0x4C, 0x56, 0x01 }, "flv", "Flash video file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x01, 0x0F, 0x00, 0x00 }, "mdf", "Microsoft SQL Server 2000 database"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0x02, 0x00 }, "cur", "Windows cursor file"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0x01, 0xBA }, "vob", "DVD Video Movie File (video/dvd, video/mpeg)"));
            fileTypes.Add(new FileTypeFileInfo(new byte[] { 0x00, 0x00, 0x01, 0x00 }, "ico", "Windows icon file"));

            return fileTypes;
        }
        #endregion

        #region IndexOf
        /// <summary>
        /// IndexOf function for byte arrays
        /// </summary>
        /// <param name="input"></param>
        /// <param name="pattern"></param>
        /// <param name="startIndex"></param>
        /// <returns></returns>
        private static int IndexOf(byte[] input, byte[] pattern, int startIndex)
        {
            var firstByte = pattern[0];
            int index;

            if ((index = Array.IndexOf(input, firstByte, startIndex)) >= 0)
            {
                for (var i = 0; i < pattern.Length; i++)
                {
                    if (index + i >= input.Length)
                        return -1;

                    if (pattern[i] == input[index + i]) continue;
                    index += i;
                    IndexOf(input, pattern, index);
                }
            }

            return index;
        }
        #endregion

        #region GetFileTypeFileInfo
        /// <summary>
        /// Returns a <see cref="FileTypeFileInfo"/> object by looking to the magic bytes of the <param ref="fileBytes"/> array.
        /// A <see cref="FileTypeFileInfo"/> object
        /// </summary>
        /// <param name="fileBytes">The bytes of the file</param>
        /// <returns></returns>
        public static FileTypeFileInfo GetFileTypeFileInfo(byte[] fileBytes)
        {
            var dataAsString = Encoding.ASCII.GetString(fileBytes);
            var result = FileTypes.FirstOrDefault(fileType => dataAsString.StartsWith(fileType.MagicBytesAsString));

            // Omdat er bepaalde bestanden zijn die we nog nader moeten onderzoeken gooien we deze door de swtich statement
            if (result != null)
            {
                switch (result.Extension)
                {
                    case MicroSoftOffice:
                        return CheckMicrosoftOfficeFormatsWithAsciiReading(fileBytes);

                    case ZipOrOffice2007:
                        return CheckZipOrOpenOfficeFormat(fileBytes);

                    case ByteOrderMarkerFile:
                        // Convert unicode file to ascii format
                        var temp = StripByteOrderMarker(fileBytes);
                        temp = Encoding.ASCII.GetBytes(Encoding.Unicode.GetString(temp));
                        return GetFileTypeFileInfo(temp);

                    default:
                        return result;
                }
            }

            return InspectTextBasedFileFormats(fileBytes);
        }
        #endregion

        #region InspectTextBasedFileFormats
        private static FileTypeFileInfo InspectTextBasedFileFormats(byte[] fileBytes)
        {
            var fileAscii = Encoding.ASCII.GetString(fileBytes);
            var lines = fileAscii.Split('\n');
            foreach(var line in lines)
            {
                var tempLine = line.ToLowerInvariant();

                if (tempLine.Contains("<?xml"))
                    return new FileTypeFileInfo(null, "xml", "Extensible Markup Language");

                if (tempLine.Contains("mime-version: 1.0"))
                    return new FileTypeFileInfo(null, "eml", "Extended Markup Language");

                if (tempLine.Contains("<html"))
                    return new FileTypeFileInfo(null, "htm", "Hypertext Markup Language");
            }

            return null;
        }
        #endregion

        #region CheckMicrosoftOfficeFormatsWithAsciiReading
        /// <summary>
        /// Tries to recognize an Microsoft Office file by looking to it's internal bytes
        /// </summary>
        /// <param name="fileBytes">The bytes of the file</param>
        /// <returns></returns>
        private static FileTypeFileInfo CheckMicrosoftOfficeFormatsWithAsciiReading(byte[] fileBytes)
        {
            var fileAscii = Encoding.ASCII.GetString(fileBytes);

            if (fileAscii.Contains("Microsoft Office Word"))
            {
                if (fileAscii.Contains("Word_sjabloon") || fileAscii.Contains("Word_template"))
                    return new FileTypeFileInfo(null, "dot", "Microsoft Word template");

                return new FileTypeFileInfo(null, "doc", "Microsoft Word binary format");
            }

            if (fileAscii.Contains("Microsoft Excel"))
                return new FileTypeFileInfo(null, "xls", "Microsoft Excel binary format");

            if (fileAscii.Contains("Microsoft Office PowerPoint"))
                return new FileTypeFileInfo(null, "ppt", "Microsoft PowerPoint binary format");

            if (fileAscii.Contains("Microsoft Visio"))
                return new FileTypeFileInfo(null, "vsd", "Microsoft Visio binary format");

            return new FileTypeFileInfo(null, string.Empty, "Unknown file type");
        }
        #endregion

        #region CheckZipOrOpenOfficeFormat
        /// <summary>
        /// Tries to recognize an Microsoft Office 2007+ file by looking to it's internal bytes
        /// </summary>
        /// <param name="fileBytes">The bytes of the file</param>
        private static FileTypeFileInfo CheckZipOrOpenOfficeFormat(byte[] fileBytes)
        {
            var fileAscii = Encoding.ASCII.GetString(fileBytes);

            if (fileAscii.Contains("word/_rels/"))
                return new FileTypeFileInfo(null, "docx", "Microsoft Word open XML document format");

            if (fileAscii.Contains("xl/_rels/workbook"))
                return new FileTypeFileInfo(null, "xlsx", "Microsoft Excel open XML document format");

            if (fileAscii.Contains("ppt/slides/_rels"))
                return new FileTypeFileInfo(null, "pptx", "Microsoft PowerPoint open XML document format");

            if (fileAscii.Contains("CHNKWKS"))
                return new FileTypeFileInfo(null, "wks", "Microsoft Works");

            // Anders is het waarschijnlijk een ZIP bestand
            return new FileTypeFileInfo(null, "zip", "Zip compressed archive");
        }
        #endregion

        #region StripByteOrderMarker
        /// <summary>
        /// Strips the BOM from the magic bytes array
        /// </summary>
        /// <param name="magicBytes"></param>
        /// <returns>Byte array without BOM</returns>
        private static byte[] StripByteOrderMarker(byte[] magicBytes)
        {
            if (IndexOf(magicBytes, new byte[] { 0xDD, 0x73, 0x66, 0x73 }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0xFF, 0xFE, 0x00, 0x00 }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0x00, 0x00, 0xFE, 0xFF }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0x84, 0x31, 0x95, 0x33 }, 0) == 0)
                return magicBytes.Select(m => m).Skip(4).ToArray();

            if (IndexOf(magicBytes, new byte[] { 0xEF, 0xBB, 0xBF }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0x2B, 0x2F, 0x76 }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0x0E, 0xFE, 0xFF }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0xFB, 0xEE, 0x28 }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0xF7, 0x64, 0x4C }, 0) == 0)
                return magicBytes.Select(m => m).Skip(3).ToArray();

            if (IndexOf(magicBytes, new byte[] { 0xFE, 0xFF }, 0) == 0 ||
                IndexOf(magicBytes, new byte[] { 0xFF, 0xFE }, 0) == 0)
                return magicBytes.Select(m => m).Skip(2).ToArray();

            return null;
        }
        #endregion
    }
}
