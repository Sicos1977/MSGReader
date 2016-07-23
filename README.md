What is MSGReader
=========

MSGReader is a C# .NET 4.0 library to read Outlook MSG and EML (Mime 1.0) files. Almost all common object in Outlook are supported:

- E-mail
- Appointment
- Task
- Contact card
- Sticky note

It supports all body types there are in MSG files, this includes:

- Text
- HTML
- HTML embedded into RTF
- RTF

MSGReader has only a few options to manipulate an MSG file. The only option you have is that you can remove attachments and then save the file to a new one.

If you realy want to write MSG files then see my MsgKit project on GitHub (https://github.com/Sicos1977/MsgKit)

Translations
============

- Kees van Spelde
    - English (US)
    - Dutch

- Ronald Kohl
    - German

- Yan Grenier (@ygrenier on GitHub)
    - French

Installing via NuGet
====================

The easiest way to install MSGReader is via NuGet.

In Visual Studio's Package Manager Console, simply enter the following command:

    Install-Package MSGReader


Side note
=========

This project can also be used from a COM based language like VB script or VB6.
To use it first compile the code and register the com visible assembly with the command:

Regasm.exe /codebase MsgReader.dll

After that you can call it like this:

dim msgreader

set msgreader = createobject("MsgReader.Reader")

msgreader.ExtractToFolderFromCom "the msg file to read", "the folder where to place the extracted files"


License
=======

Copyright 2013-2016 Kees van Spelde.

Licensed under the The Code Project Open License (CPOL) 1.02; you may not use this software except in compliance with the License. You may obtain a copy of the License at:

http://www.codeproject.com/info/cpol10.aspx

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

Core Team
=========
    Sicos1977 (Kees van Spelde)

Support
=======
If you like my work then please consider a donation as a thank you.

<a href="https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=NS92EXB2RDPYA" target="_blank"><img src="https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif" /></a>
