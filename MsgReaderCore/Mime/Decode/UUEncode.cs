using System;
using System.IO;
using MsgReader.Helpers;

namespace MsgReader.Mime.Decode;

internal static class UUEncode
{
    private static readonly byte[] UUDecMap =
    {
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
    };

    public static byte[] Decode(byte[] encodeBytes)
    {
        using (Stream input = StreamHelpers.Manager.GetStream("UUEncode.cs", encodeBytes, 0, encodeBytes.Length))
        using (var output = StreamHelpers.Manager.GetStream())
        {
            try
            {
                if (input == null)
                    throw new ArgumentNullException("input");

                var len = input.Length;
                if (len == 0)
                    return new byte[0];

                long didx = 0;
                var nextByte = input.ReadByte();
                while (nextByte >= 0)
                {
                    // Get line length (in number of encoded octets)
                    int line_len = UUDecMap[nextByte];

                    // Ascii printable to 0-63 and 4-byte to 3-byte conversion
                    var end = didx + line_len;
                    byte A, B, C, D;
                    if (end > 2)
                        while (didx < end - 2)
                        {
                            A = UUDecMap[input.ReadByte()];
                            B = UUDecMap[input.ReadByte()];
                            C = UUDecMap[input.ReadByte()];
                            D = UUDecMap[input.ReadByte()];

                            output.WriteByte((byte)(((A << 2) & 255) | ((B >> 4) & 3)));
                            output.WriteByte((byte)(((B << 4) & 255) | ((C >> 2) & 15)));
                            output.WriteByte((byte)(((C << 6) & 255) | (D & 63)));
                            didx += 3;
                        }

                    if (didx < end)
                    {
                        A = UUDecMap[input.ReadByte()];
                        B = UUDecMap[input.ReadByte()];
                        output.WriteByte((byte)(((A << 2) & 255) | ((B >> 4) & 3)));
                        didx++;
                    }

                    if (didx < end)
                    {
                        B = UUDecMap[input.ReadByte()];
                        C = UUDecMap[input.ReadByte()];
                        output.WriteByte((byte)(((B << 4) & 255) | ((C >> 2) & 15)));
                        didx++;
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
                    } while (nextByte >= 0 && (nextByte == '\n' || nextByte == '\r'));
                }

                return output.ToArray();
            }
            catch (Exception)
            {
                return new byte[0];
            }
        }
    }
}