using System;
using System.IO;
using System.Net.Mail;
using System.Reflection;

namespace MsgReader.Helpers;

internal static class MailMessageExtension
{
    internal static void WriteTo(this MailMessage mail, Stream stream)
    {
        var assembly = typeof(SmtpClient).Assembly;
        var mailWriterType = assembly.GetType("System.Net.Mail.MailWriter");

        var mailWriterConstructor = mailWriterType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { typeof(Stream), typeof(bool) }, null);

        if (mailWriterConstructor == null) 
            throw new Exception("Failed to find internal constructor for MailWriter");
        
        var mailWriter = mailWriterConstructor.Invoke(new object[] { stream, true });
        var sendMethod = typeof(MailMessage).GetMethod("Send", BindingFlags.Instance | BindingFlags.NonPublic);
        
        if (sendMethod == null) 
            throw new Exception("Failed to find internal 'Send' method on MailMessage");
        
        sendMethod.Invoke(mail, BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { mailWriter, true, true }, null!);

        var closeMethod = mailWriter.GetType().GetMethod("Close", BindingFlags.Instance | BindingFlags.NonPublic);
        
        if (closeMethod != null)
            closeMethod.Invoke(mailWriter, BindingFlags.Instance | BindingFlags.NonPublic, null, Array.Empty<object>(),
                null!);
        else
            throw
                new Exception("Failed to find internal 'Close' method on MailWriter");
    }
}