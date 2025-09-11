using System;
using System.IO;
using MsgReader.Helpers;
// ReSharper disable InconsistentNaming

namespace MsgReader.Mime.Decode;

internal static class UUEncode
{
    #region Fields
    private static readonly byte[] UUDecMap =
    [
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
        0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F,
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17,
        0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F,
        0x20, 0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27,
        0x28, 0x29, 0x2A, 0x2B, 0x2C, 0x2D, 0x2E, 0x2F,
        0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37,
        0x38, 0x39, 0x3A, 0x3B, 0x3C, 0x3D, 0x3E, 0x3F,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
    ];
    #endregion

    #region Decode
    public static byte[] Decode(byte[] encodeBytes)
    {
        using Stream input = StreamHelpers.Manager.GetStream("UUEncode.cs", encodeBytes, 0, encodeBytes.Length);
        using var output = StreamHelpers.Manager.GetStream();
        try
        {
            if (input == null)
                throw new ArgumentNullException(nameof(encodeBytes));

            var len = input.Length;
            if (len == 0)
                return [];

            long didX = 0;
            var nextByte = input.ReadByte();
            while (nextByte >= 0)
            {
                // Get line length (in number of encoded octets)
                int lineLen = UUDecMap[nextByte];

                // Ascii printable to 0-63 and 4-byte to 3-byte conversion
                var end = didX + lineLen;
                byte a, b, c;
                if (end > 2)
                    while (didX < end - 2)
                    {
                        a = UUDecMap[input.ReadByte()];
                        b = UUDecMap[input.ReadByte()];
                        c = UUDecMap[input.ReadByte()];
                        var d = UUDecMap[input.ReadByte()];

                        output.WriteByte((byte)(((a << 2) & 255) | ((b >> 4) & 3)));
                        output.WriteByte((byte)(((b << 4) & 255) | ((c >> 2) & 15)));
                        output.WriteByte((byte)(((c << 6) & 255) | (d & 63)));
                        didX += 3;
                    }

                if (didX < end)
                {
                    a = UUDecMap[input.ReadByte()];
                    b = UUDecMap[input.ReadByte()];
                    output.WriteByte((byte)(((a << 2) & 255) | ((b >> 4) & 3)));
                    didX++;
                }

                if (didX < end)
                {
                    b = UUDecMap[input.ReadByte()];
                    c = UUDecMap[input.ReadByte()];
                    output.WriteByte((byte)(((b << 4) & 255) | ((c >> 2) & 15)));
                    didX++;
                }

                // Skip padding
                do
                {
                    nextByte = input.ReadByte();
                } while (nextByte >= 0 && nextByte != '\n' && nextByte != '\r');

                // Skip end of line
                do
                {
                    nextByte = input.ReadByte();
                } while (nextByte >= 0 && nextByte is '\n' or '\r');
            }

            return output.ToArray();
        }
        catch (Exception)
        {
            return [];
        }
    }
    #endregion
}