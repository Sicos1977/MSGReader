What is MSGReader
=========

MSGReader is a C# .NET 4.6.1 and Standard 2.1 library to read Outlook MSG and EML (Mime 1.0) files. Almost all common object in Outlook are supported:

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

Read properties from an Outlook (msg) message
============
```c#
using (var msg = new MsgReader.Outlook.Storage.Message("d:\\testfile.msg"))
{
        var from = msg.Sender;
        var sentOn = msg.SentOn;
        var recipientsTo = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false);
        var recipientsCc = msg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false);
        var subject = msg.Subject;
        var htmlBody = msg.BodyHtml;
        // etc...
}
```

Read properties from an Outlook (eml) message
============
```c#
var fileInfo = new FileInfo("d:\\testfile.eml");
var eml = MsgReader.Mime.Message.Load(fileInfo);

if (eml.Headers != null)
{
        if (eml.Headers.To != null)
        {
            foreach (var recipient in eml.Headers.To)
            {
                var to = recipient.Address;            
            }
        }
}

var subject = eml.Headers.Subject;

if (eml.TextBody != null)
{
        var textBody = System.Text.Encoding.UTF8.GetString(eml.TextBody.Body);
}

if (eml.HtmlBody != null)
{
        var htmlBody = System.Text.Encoding.UTF8.GetString(eml.HtmlBody.Body);
}

// etc...
```

Delete attachment from an Outlook message
============

This example deletes the first attachment

```c#
var outlook = new Storage.Message(fileName, FileAccess.ReadWrite);
outlook.DeleteAttachment(outlook.Attachments[0]);
outlook.Save("d:\\deleted.msg");
```

Translations
============

- Kees van Spelde
    - English (US)
    - Dutch

- Ronald Kohl
    - German

- Yan Grenier (@ygrenier on GitHub)
    - French

- xupefei
    - Simpl Chinese

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

```vb
dim msgreader

set msgreader = createobject("MsgReader.Reader")
msgreader.ExtractToFolderFromCom "the msg file to read", "the folder where to place the extracted files"
```

## License Information

MsgReader is Copyright (C) 2013-2021 Magic-Sessions and is licensed under the MIT license:

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.

Core Team
=========
    Sicos1977 (Kees van Spelde)