MSGReader
=========

- 2014-07-06 Version 1.7
    - Added support for signed MSG files
    - Added support for EML (Mime 1.0 encoded) files
    - Fixed issue with double quotes not getting parsed correctly from HTML embedded into RTF

- 2014-06-20 Version 1.6.1
    - Added support for multi value STRING8 type properties
    - Added option to msgviewer to save the E-mail as a text file
    - Fixed issue with window positioning in the msgviewer program
    - Added CPOL license file

- 2014-06-12 Version 1.6
    - Fixed bug in E-mail CC field that also went into the BCC field
    - Fixed bug in appointment mapping
    - Added AllAttendees, ToAttendees and CCAttendees properties to Appointment class
    - Added MSGMapper tool, with this tool properties from msg files can be mapped to extended file properties 
      (Windows 7 or higher is needed for this)
    - Added Outlook properties to extended file properties mapping for:
        - E-mails
        - Appointments
        - Tasks
        - Contacts

- 2014-04-29 Version 1.5
    - Added Outlook contact(card) support
    - Made properties late binding

- 2014-04-21 Version 1.4
    - Full support for MAPI named properties
    - Added support for OLE attachments
    - Added Outlook Appointment support
    - Added Outlook Sticky notes support
    - Added support so that ole images in RTF files get rendered correctly after converison to HTML (Tried to be as        close as possible to how it looks in Outlook).
    - Extended E-mail support with RTF to HTML conversion. When there is no HTML body part and there is an RTF part 
      then this one gets converted to HTML.
    - Moved all language specific things to a separate class so that this component can be easily translated to other       languages. Please send me the translations if you do this.
    - Fixed a lot of bugs and made speed improvements

- 2014-03-03 Version 1.3
      - Completed implementing Outlook flag system on E-mail MSG object
      - Made all the MAPI functions private
      - Moved all language related things into a separate file so that it is easy to translate
      - Moved all MAPI constants to a seperate file and added comment
      - Removed some unused classes
      - Cleaned up the code
      - Added option to convert attachment locations to hyperlinks
      - Fixed remove < and > when there is no E-mail address and only a displayname
      - Fixed issue with SentOn and Received on not being parsed as a DateTime values
      - Added support for categories in msg files (Outlook 2007 or later)
      - Fixed issue with E-mail address and displayname being swapped
      - Added RtfToHtmlConverter class (to convert RTF to HTML)
      - Added Message Type property so that we know what kind of MSG object we have

- 2014-03-20 Version 1.2.1
      -  Added support for double byte char sets like Chinese

- 2014-03-18 Version 1.2
      -  Fixed an issue with the Sent On (this was not set to the local timezone)
      -  Added Received On, this is now added to the injected Outlook header
      -  Added using statement to message object
      -  Made some Native Methods internal
      -  Fixed some disposing issues, this was done more the once on some places
      -  Refactored the code so that everything was correct according to the Microsoft code rules


- 2014-03-06 Version 1.1
      -  Added support for special characters like German umlauts, they are parsed out of the HTML text that is                 embedded inside the RTF
      -  The RTFBody was loaded 4 times instead of once
      -  A CC was always added even when there was no CC (the To was taken in that case)
      -  Fixed some minor issues and cleaned up the code 

      
- 2014-01-16 First release

Translations
============

- Kees van Spelde
    - English (US)
    - Dutch

- Ronald Kohl
    - German

Side note
=========

This project can also be used from a COM based language like VB script or VB6.
To use it first compile the code and register the com visible assembly with the command:

Regasm.exe /codebase DocumentServices.Modules.Readers.MsgReader.dll

After that you can call it like this:

dim msgreader
set msgreader = createobject("DocumentServices.Modules.Readers.MsgReader.Reader")
msgreader.ExtractToFolder "the msg file to read", "the folder where to place the extracted files"


License
=======

Copyright 2013-2014 Kees van Spelde.

Licensed under the The Code Project Open License (CPOL) 1.02; you may not use this software except in compliance with the License. You may obtain a copy of the License at:

http://www.codeproject.com/info/cpol10.aspx

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

Core Team
=========
    Sicos1977 (Kees van Spelde)

