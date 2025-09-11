﻿using System;
using System.IO;
using System.Net.Mail;
using System.Reflection;

namespace MsgReader.Helpers;

internal static class MailMessageExtension
{
    #region WriteTo
    /// <summary>
    ///     Extension method to get the raw content of a <see cref="MailMessage" />
    /// </summary>
    /// <param name="mail"><see cref="MailMessage"/></param>
    /// <param name="stream">The stream to write to</param>
    /// <exception cref="Exception"></exception>
    internal static void WriteTo(this MailMessage mail, Stream stream)
    {
        var assembly = typeof(SmtpClient).Assembly;
        
        var mailWriterType = assembly.GetType("System.Net.Mail.MailWriter") ??
                             throw new Exception("Failed to find internal constructor for MailWriterType");

#if (NETSTANDARD2_0_OR_GREATER)
        var mailWriterConstructor = mailWriterType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null,
                                        [typeof(Stream), typeof(bool)], null) ??
                                    throw new Exception("Failed to find internal constructor for MailWriter");

        var mailWriter = mailWriterConstructor.Invoke([stream, true]);
        var sendMethod = typeof(MailMessage).GetMethod("Send", BindingFlags.Instance | BindingFlags.NonPublic) ??
                         throw new Exception("Failed to find internal 'Send' method on MailMessage");
        
        sendMethod.Invoke(mail, BindingFlags.Instance | BindingFlags.NonPublic, null, [mailWriter, true, true], null!);

        var closeMethod = mailWriter.GetType().GetMethod("Close", BindingFlags.Instance | BindingFlags.NonPublic);

        if (closeMethod != null)
            closeMethod.Invoke(mailWriter, BindingFlags.Instance | BindingFlags.NonPublic, null, [], null!);
        else
            throw
                new Exception("Failed to find internal 'Close' method on MailWriter");

        return;
#elif(NETFRAMEWORK)
        var mailWriterConstructor =
            mailWriterType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null, [typeof(Stream)], null) ??
            throw new Exception("Failed to find internal constructor for MailWriter");

        var mailWriter = mailWriterConstructor.Invoke([stream]);
        var sendMethod = typeof(MailMessage).GetMethod("Send", BindingFlags.Instance | BindingFlags.NonPublic) ??
                         throw new Exception("Failed to find internal 'Send' method on MailMessage");
        
        sendMethod.Invoke(mail, BindingFlags.Instance | BindingFlags.NonPublic, null, [mailWriter, true, true], null!);

        var closeMethod = mailWriter.GetType().GetMethod("Close", BindingFlags.Instance | BindingFlags.NonPublic);

        if (closeMethod != null)
            closeMethod.Invoke(mailWriter, BindingFlags.Instance | BindingFlags.NonPublic, null, [], null!);
        else
            throw
                new Exception("Failed to find internal 'Close' method on MailWriter");
#else
        throw new Exception("Unsupported platform");
#endif
    }
    #endregion
}