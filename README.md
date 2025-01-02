What is MSGReader
=========

MSGReader is a C# .NET 4.6.2, .NET Standard 2.0 and .NET Standard 2.1 assembly to read Outlook MSG and EML (Mime 1.0) files. Almost all common object in Outlook are supported:

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

Detecting charset encoding in MSG files with HTML encapuslated into RTF that use different font set encodings
============

Most of the times when an HTML body is used in an MSG file this HTML body is encapsulated into RTF.
See this link for more info --> https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfex/4f09a809-9910-43f3-a67c-3506b09ca5ac

When an HTML body contains chars that are not in the default extended ASCII range then these chars are encoded. This is normally not a problem when just one language is used.
When multiple languages are used then it is quite often that the RTF is not build correctly in a way so that MSGReader can figure out what kind of encoding needs to be used to decode the chars. Because of this MSGReader uses the nuget package UTF.Unknown (https://www.nuget.org/packages/UTF.Unknown/) to try to figure out in what kind of encoding a char is stored. Most of the times this works correctly and because of that a threshold is set to a value of 0.90 so that when the detection level passes this value it will be seen as a valid char.

If you still have bad results you can control this confidence level yourself by using the property `CharsetDetectionEncodingConfidenceLevel` in the `Reader` or `Message` class


```c#
/// <summary>
///     When an MSG file contains an RTF file with encapsulated HTML and the RTF
///     uses fonts with different encodings then this levels set the threshold that
///     an encoded string detection levels needs to be before recognizing it as a valid
///     string. When the detection level is lower than this setting then the default RTF
///     encoding is used to decode the encoded char 
/// </summary>
/// <remarks>
///     Default this value is set to 0.90, any values lower then 0.70 probably give bad
///     results
/// </remarks>
public float CharsetDetectionEncodingConfidenceLevel { get; set; } = 0.90f;
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

[![NuGet](https://img.shields.io/nuget/v/MSGReader.svg?style=flat-square)](https://www.nuget.org/packages/MSGReader)

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

MsgReader is Copyright (C) 2013-2025 Magic-Sessions and is licensed under the MIT license:

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

Support
=======
If you like my work then please consider a donation as a thank you by using the donate button at the top
