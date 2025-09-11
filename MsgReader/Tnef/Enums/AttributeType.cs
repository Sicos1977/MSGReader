namespace MsgReader.Tnef.Enums;

internal enum AttributeType
{
    Triples = 0x00000000,
    String = 0x00010000,
    Text = 0x00020000,
    Date = 0x00030000,
    Short = 0x00040000,
    Long = 0x00050000,
    Byte = 0x00060000,
    Word = 0x00070000,
    DWord = 0x00080000,
    Max = 0x00090000
}