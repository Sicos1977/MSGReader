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
    /// This class contains all known mimetypes
    /// </summary>
    internal class MimeType
    {
        #region GetExtensionFromMimeType
        /// <summary>
        /// Returns the file extension for the given <paramref name="mimeType"/>. An empty string
        /// is returned when the mimetype is not found.
        /// </summary>
        /// <param name="mimeType">The mime type</param>
        /// <returns></returns>
        public static string GetExtensionFromMimeType(string mimeType)
        {
            if (mimeType == null)
                return string.Empty;

            switch (mimeType.ToLowerInvariant())
            {
                case "application/fractals":
                    return ".fif";

                case "application/futuresplash":
                    return ".spl";

                case "application/hta":
                    return ".hta";

                case "application/mac-binhex40":
                    return ".hqx";

                case "application/ms-infopath.xml":

                case "application/ms-vsi":
                    return ".vsi";

                case "application/msaccess":
                    return ".accdb";

                case "application/msaccess.addin":
                    return ".accda";

                case "application/msaccess.cab":
                    return ".accdc";

                case "application/msaccess.exec":
                    return ".accde";

                case "application/msaccess.ftemplate":
                    return ".accft";

                case "application/msaccess.runtime":
                    return ".accdr";

                case "application/msaccess.template":
                    return ".accdt";

                case "application/msaccess.webapplication":
                    return ".accdw"
                        ;
                case "application/msonenote":
                    return ".one";

                case "application/msword":
                    return ".doc";

                case "application/opensearchdescription+xml":
                    return ".osdx";

                case "application/oxps":
                    return ".oxps";

                case "application/pdf":
                    return ".pdf";

                case "application/pkcs10":
                    return ".p10";

                case "application/pkcs7-mime":
                    return ".p7c";

                case "application/pkcs7-signature":
                    return ".p7s";

                case "application/pkix-cert":
                    return ".cer";

                case "application/pkix-crl":
                    return ".crl";

                case "application/postscript":
                    return ".ps";

                case "application/tif":
                    return ".tif";

                case "application/tiff":
                    return ".tiff";

                case "application/vnd.adobe.acrobat-security-settings":
                    return ".acrobatsecuritysettings";

                case "application/vnd.adobe.pdfxml":
                    return ".pdfxml";

                case "application/vnd.adobe.pdx":
                    return ".pdx";

                case "application/vnd.adobe.xdp+xml":
                    return ".xdp";

                case "application/vnd.adobe.xfd+xml":
                    return ".xfd";

                case "application/vnd.adobe.xfdf":
                    return ".xfdf";

                case "application/vnd.fdf":
                    return ".fdf";

                case "application/vnd.ms-excel":
                    return ".xls";

                case "application/vnd.ms-excel.12":
                    return ".xlsx";

                case "application/vnd.ms-excel.addin.macroEnabled.12":
                    return ".xlam";

                case "application/vnd.ms-excel.sheet.binary.macroEnabled.12":
                    return ".xlsb";

                case "application/vnd.ms-excel.sheet.macroEnabled.12":
                    return ".xlsm";

                case "application/vnd.ms-excel.template.macroEnabled.12":
                    return ".xltm";

                case "application/vnd.ms-officetheme":
                    return ".thmx";

                case "application/vnd.ms-pki.certstore":
                    return ".sst";

                case "application/vnd.ms-pki.pko":
                    return ".pko";

                case "application/vnd.ms-pki.seccat":
                    return ".cat";

                case "application/vnd.ms-powerpoint":
                    return ".ppt";

                case "application/vnd.ms-powerpoint.12":
                    return ".pptx";

                case "application/vnd.ms-powerpoint.addin.macroEnabled.12":
                    return ".ppam";

                case "application/vnd.ms-powerpoint.presentation.macroEnabled.12":
                    return ".pptm";

                case "application/vnd.ms-powerpoint.slide.macroEnabled.12":
                    return ".sldm";

                case "application/vnd.ms-powerpoint.slideshow.macroEnabled.12":
                    return ".ppsm";

                case "application/vnd.ms-powerpoint.template.macroEnabled.12":
                    return ".potm";

                case "application/vnd.ms-publisher":
                    return ".pub";

                case "application/vnd.ms-visio.viewer":
                    return ".vsd";

                case "application/vnd.ms-word.document.12":
                    return ".docx";

                case "application/vnd.ms-word.document.macroEnabled.12":
                    return ".docm";

                case "application/vnd.ms-word.template.12":
                    return ".dotx";

                case "application/vnd.ms-word.template.macroEnabled.12":
                    return ".dotm";

                case "application/vnd.ms-wpl":
                    return ".wpl";

                case "application/vnd.ms-xpsdocument":
                    return ".xps";

                case "application/vnd.oasis.opendocument.presentation":
                    return ".odp";

                case "application/vnd.oasis.opendocument.spreadsheet":
                    return ".ods";

                case "application/vnd.oasis.opendocument.text":
                    return ".odt";

                case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
                    return ".pptx";

                case "application/vnd.openxmlformats-officedocument.presentationml.slide":
                    return ".sldx";

                case "application/vnd.openxmlformats-officedocument.presentationml.slideshow":
                    return ".ppsx";

                case "application/vnd.openxmlformats-officedocument.presentationml.template":
                    return ".potx";

                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    return ".xlsx";

                case "application/vnd.openxmlformats-officedocument.spreadsheetml.template":
                    return ".xltx";

                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    return ".docx";

                case "application/vnd.openxmlformats-officedocument.wordprocessingml.template":
                    return ".dotx";

                case "application/windows-appcontent+xml":
                    return ".appcontent-ms";

                case "application/x-compress":
                    return ".z";

                case "application/x-compressed":
                    return ".tgz";

                case "application/x-dtcp1":
                    return ".dtcp-ip";

                case "application/x-gzip":
                    return ".gz";

                case "application/x-jtx+xps":
                    return ".jtx";

                case "application/x-latex":
                    return ".latex";

                case "application/x-mix-transfer":
                    return ".nix";

                case "application/x-mplayer2":
                    return ".asx";

                case "application/x-ms-application":
                    return ".application";

                case "application/x-ms-vsto":
                    return ".vsto";

                case "application/x-ms-wmd":
                    return ".wmd";

                case "application/x-ms-wmz":
                    return ".wmz";

                case "application/x-ms-xbap":
                    return ".xbap";

                case "application/x-mswebsite":
                    return ".website";

                case "application/x-pkcs12":
                    return ".p12";

                case "application/x-pkcs7-certificates":
                    return ".p7b";

                case "application/x-pkcs7-certreqresp":
                    return ".p7r";

                case "application/x-shockwave-flash":
                    return ".swf";

                case "application/x-skype":
                    return ".skype";

                case "application/x-stuffit":
                    return ".sit";

                case "application/x-tar":
                    return ".tar";

                case "application/x-troff-man":
                    return ".man";

                case "application/x-wlpg-detect":
                    return ".wlpginstall";

                case "application/x-wlpg3-detect":
                    return ".wlpginstall3";

                case "application/x-wmplayer":
                    return ".asx";

                case "application/x-x509-ca-cert":
                    return ".cer";

                case "application/x-zip-compressed":
                    return ".zip";

                case "application/xaml+xml":
                    return ".xaml";

                case "application/xhtml+xml":
                    return ".xht";

                case "application/xml":
                    return ".xml";

                case "application/xps":
                    return ".xps";

                case "application/zip":
                    return ".zip";

                case "audio/3gpp":
                    return ".3gp";

                case "audio/3gpp2":
                    return ".3g2";

                case "audio/aiff":
                    return ".aiff";

                case "audio/basic":
                    return ".au";

                case "audio/ec3":
                    return ".ec3";

                case "audio/l16":
                    return ".lpcm";

                case "audio/mid":
                    return ".mid";

                case "audio/midi":
                    return ".mid";

                case "audio/mp3":
                    return ".mp3";

                case "audio/mp4":
                    return ".m4a";

                case "audio/mpeg":
                    return ".mp3";

                case "audio/mpegurl":
                    return ".m3u";

                case "audio/mpg":
                    return ".mp3";

                case "audio/scpls":
                    return ".pls";

                case "audio/vnd.dlna.adts":
                    return ".adts";

                case "audio/vnd.dolby.dd-raw":
                    return ".ac3";

                case "audio/wav":
                    return ".wav";

                case "audio/x-aiff":
                    return ".aiff";
                        
                case "audio/x-mid":
                    return ".mid";

                case "audio/x-midi":
                    return ".mid";
                      
                case "audio/x-mp3":
                    return ".mp3";

                case "audio/x-mpeg":
                    return ".mp3";

                case "audio/x-mpegurl":
                    return ".m3u";
                        
                case "audio/x-mpg":
                    return ".mp3";

                case "audio/x-ms-wax":
                    return ".wax";

                case "audio/x-ms-wma":
                    return ".wma";

                case "audio/x-scpls":
                    return ".pls";

                case "audio/x-wav":
                    return ".wav";

                case "image/bmp":
                    return ".bmp";

                case "image/gif":
                    return ".gif";

                case "image/jpeg":
                    return ".jpg";

                case "image/pjpeg":
                    return ".jpg";

                case "image/png":
                    return ".png";
                case "image/svg+xml":
                    return ".svg";

                case "image/tiff":
                    return ".tiff";

                case "image/vnd.ms-dds":
                    return ".dds";

                case "image/vnd.ms-photo":
                    return ".wdp";

                case "image/x-emf":
                    return ".emf";

                case "image/x-icon":
                    return ".ico";

                case "image/x-png":
                    return ".png";

                case "image/x-wmf":
                    return ".wmf";


                case "interface/x-winamp-lang":
                    return ".wlz";

                case "interface/x-winamp-skin":
                    return ".wsz";

                case "interface/x-winamp3-skin":
                    return ".wal";

                case "message/rfc822":
                    return ".eml";

                case "midi/mid":
                    return ".mid";

                case "model/vnd.dwfx+xps":
                    return ".dwfx";

                case "model/vnd.easmx+xps":
                    return ".easmx";

                case "model/vnd.edrwx+xps":
                    return ".edrwx";

                case "model/vnd.eprtx+xps":
                    return ".eprtx";

                case "pkcs10":
                    return ".p10";

                case "pkcs7-mime":
                    return ".p7m";

                case "pkcs7-signature":
                    return ".p7s";

                case "pkix-cert":
                    return ".cer";

                case "pkix-crl":
                    return ".crl";
                        
                case "text/calendar":
                    return ".ics";

                case "text/css":
                    return ".css";

                case "text/directory":
                    return ".vcf";

                case "text/directory;profile=vCard":
                    return ".vcf";

                case "text/html":
                    return ".htm";

                case "text/json":
                    return ".mcjson";

                case "text/plain":
                    return ".txt";

                case "text/scriptlet":
                    return ".wsc";

                case "text/vcard":
                    return ".vcf";

                case "text/x-component":
                    return ".htc";

                case "text/x-ms-contact":
                    return ".contact";

                case "text/x-ms-iqy":
                    return ".iqy";

                case "text/x-ms-odc":
                    return ".odc";

                case "text/x-ms-rqy":
                    return ".rqy";

                case "text/x-vcard":
                    return ".vcf";

                case "text/xml":
                    return ".xml";

                case "video/3gpp":
                    return ".3gp";

                case "video/3gpp2":
                    return ".3g2";

                case "video/asx":
                    return ".asx";

                case "video/avi":
                    return ".avi";

                case "video/mp4":
                    return ".mp4";

                case "video/mpeg":
                    return ".mpeg";

                case "video/mpg":
                    return ".mpeg";

                case "video/msvideo":
                    return ".avi";

                case "video/quicktime":
                    return ".mov";

                case "video/vnd.dece.mp4":
                    return ".uvu";

                case "video/vnd.dlna.mpeg-tts":
                    return ".tts";

                case "video/wtv":
                    return ".wtv";

                case "video/x-asx":
                    return ".asx";

                case "video/x-mpeg":
                    return ".mpeg";
                    
                case "video/x-mpeg2a":
                    return ".mpeg";
                   
                case "video/x-ms-asf":
                    return ".asx";

                case "video/x-ms-asf-plugin":
                    return ".asx";

                case "video/x-ms-dvr":
                    return ".dvr-ms";

                case "video/x-ms-wm":
                    return ".wm";

                case "video/x-ms-wmv":
                    return ".wmv";

                case "video/x-ms-wmx":
                    return ".wmx";

                case "video/x-ms-wvx":
                    return ".wvx";

                case "video/x-msvideo":
                    return ".avi";

                case "vnd.ms-pki.certstore":
                    return ".sst";

                case "vnd.ms-pki.pko":
                    return ".pko";

                case "vnd.ms-pki.seccat":
                    return ".cat";

                case "x-pkcs12":
                    return ".p12";

                case "x-pkcs7-certificates":
                    return ".p7b";

                case "x-pkcs7-certreqresp":
                    return ".p7r";

                case "x-x509-ca-cert":
                    return ".cer";

                default:
                    return string.Empty;
            }
        }
        #endregion
    }
}
