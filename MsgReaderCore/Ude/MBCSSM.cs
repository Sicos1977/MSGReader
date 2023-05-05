//
// Version: MPL 1.1/GPL 2.0/LGPL 2.1
//
// The contents of this file are subject to the Mozilla Public License Version
// 1.1 (the "License"); you may not use this file except in compliance with
// the License. You may obtain a copy of the License at
// http://www.mozilla.org/MPL/
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
// for the specific language governing rights and limitations under the
// License.
//
// The Original Code is Mozilla Universal charset detector code.
//
// The Initial Developer of the Original Code is
// Netscape Communications Corporation.
// Portions created by the Initial Developer are Copyright (C) 2001
// the Initial Developer. All Rights Reserved.
//
// Contributor(s):
//          Shy Shalom <shooshX@gmail.com>
//          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
// 
// Alternatively, the contents of this file may be used under the terms of
// either the GNU General Public License Version 2 or later (the "GPL"), or
// the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
// in which case the provisions of the GPL or the LGPL are applicable instead
// of those above. If you wish to allow use of your version of this file only
// under the terms of either the GPL or the LGPL, and not to allow others to
// use your version of this file under the terms of the MPL, indicate your
// decision by deleting the provisions above and replace them with the notice
// and other provisions required by the GPL or the LGPL. If you do not delete
// the provisions above, a recipient may use your version of this file under
// the terms of any one of the MPL, the GPL or the LGPL.
//

namespace MsgReader.Ude;

internal class Utf8SmModel : SmModel
{
    #region Fields
    private static readonly int[] Utf8Cls =
    {
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 00 - 07
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 10 - 17 
        BitPackage.Pack4Bits(1, 1, 1, 0, 1, 1, 1, 1), // 18 - 1f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 20 - 27 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 28 - 2f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 30 - 37 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 38 - 3f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 40 - 47 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 48 - 4f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 50 - 57 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 58 - 5f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 60 - 67 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 68 - 6f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 70 - 77 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 78 - 7f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 3, 3, 3, 3), // 80 - 87 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 88 - 8f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 90 - 97 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 98 - 9f 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // a0 - a7 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // a8 - af 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // b0 - b7 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // b8 - bf 
        BitPackage.Pack4Bits(0, 0, 6, 6, 6, 6, 6, 6), // c0 - c7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // c8 - cf 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // d0 - d7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // d8 - df 
        BitPackage.Pack4Bits(7, 8, 8, 8, 8, 8, 8, 8), // e0 - e7 
        BitPackage.Pack4Bits(8, 8, 8, 8, 8, 9, 8, 8), // e8 - ef 
        BitPackage.Pack4Bits(10, 11, 11, 11, 11, 11, 11, 11), // f0 - f7 
        BitPackage.Pack4Bits(12, 13, 13, 13, 14, 15, 0, 0) // f8 - ff 
    };

    private static readonly int[] Utf8St =
    {
        BitPackage.Pack4Bits(Error, Start, Error, Error, Error, Error, 12, 10), //00-07 
        BitPackage.Pack4Bits(9, 11, 8, 7, 6, 5, 4, 3), //08-0f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //10-17 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //18-1f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe), //20-27 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe), //28-2f 
        BitPackage.Pack4Bits(Error, Error, 5, 5, 5, 5, Error, Error), //30-37 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //38-3f 
        BitPackage.Pack4Bits(Error, Error, Error, 5, 5, 5, Error, Error), //40-47 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //48-4f 
        BitPackage.Pack4Bits(Error, Error, 7, 7, 7, 7, Error, Error), //50-57 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //58-5f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, 7, 7, Error, Error), //60-67 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //68-6f 
        BitPackage.Pack4Bits(Error, Error, 9, 9, 9, 9, Error, Error), //70-77 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //78-7f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, 9, Error, Error), //80-87 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //88-8f 
        BitPackage.Pack4Bits(Error, Error, 12, 12, 12, 12, Error, Error), //90-97 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //98-9f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, 12, Error, Error), //a0-a7 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //a8-af 
        BitPackage.Pack4Bits(Error, Error, 12, 12, 12, Error, Error, Error), //b0-b7 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error), //b8-bf 
        BitPackage.Pack4Bits(Error, Error, Start, Start, Start, Start, Error, Error), //c0-c7 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, Error, Error) //c8-cf  
    };

    private static readonly int[] Utf8CharLenTable = { 0, 1, 0, 0, 0, 0, 2, 3, 3, 3, 4, 4, 5, 5, 6, 6 };
    #endregion

    #region Constructor
    internal Utf8SmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Utf8Cls),
        16,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Utf8St),
        Utf8CharLenTable, "UTF-8")
    {
    }
    #endregion
}

internal class Gb18030SmModel : SmModel
{
    #region Fields
    private static readonly int[] Gb18030Cls =
    {
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 00 - 07 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 10 - 17 
        BitPackage.Pack4Bits(1, 1, 1, 0, 1, 1, 1, 1), // 18 - 1f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 20 - 27 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 28 - 2f 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // 30 - 37 
        BitPackage.Pack4Bits(3, 3, 1, 1, 1, 1, 1, 1), // 38 - 3f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 40 - 47 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 48 - 4f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 50 - 57 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 58 - 5f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 60 - 67 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 68 - 6f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 70 - 77 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 4), // 78 - 7f 
        BitPackage.Pack4Bits(5, 6, 6, 6, 6, 6, 6, 6), // 80 - 87 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // 88 - 8f 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // 90 - 97 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // 98 - 9f 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // a0 - a7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // a8 - af 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // b0 - b7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // b8 - bf 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // c0 - c7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // c8 - cf 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // d0 - d7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // d8 - df 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // e0 - e7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // e8 - ef 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 6), // f0 - f7 
        BitPackage.Pack4Bits(6, 6, 6, 6, 6, 6, 6, 0) // f8 - ff 
    };

    private static readonly int[] Gb18030St =
    {
        BitPackage.Pack4Bits(Error, Start, Start, Start, Start, Start, 3, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, Error, Error, Start), //10-17 
        BitPackage.Pack4Bits(4, Error, Start, Start, Error, Error, Error, Error), //18-1f 
        BitPackage.Pack4Bits(Error, Error, 5, Error, Error, Error, ItsMe, Error), //20-27 
        BitPackage.Pack4Bits(Error, Error, Start, Start, Start, Start, Start, Start) //28-2f 
    };

    // To be accurate, the length of class 6 can be either 2 or 4. 
    // But it is not necessary to discriminate between the two since 
    // it is used for frequency analysis only, and we are validating 
    // each code range there as well. So it is safe to set it to be 
    // 2 here. 
    private static readonly int[] Gb18030CharLenTable = { 0, 1, 1, 1, 1, 1, 2 };
    #endregion

    #region Constructor
    internal Gb18030SmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Gb18030Cls),
        7,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Gb18030St),
        Gb18030CharLenTable, "GB18030")
    {
    }
    #endregion
}

internal class Big5SmModel : SmModel
{
    #region Fields
    private static readonly int[] Big5Cls =
    {
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 00 - 07
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 10 - 17 
        BitPackage.Pack4Bits(1, 1, 1, 0, 1, 1, 1, 1), // 18 - 1f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 20 - 27 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 28 - 2f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 30 - 37 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 38 - 3f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 40 - 47 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 48 - 4f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 50 - 57 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 58 - 5f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 60 - 67 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 68 - 6f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 70 - 77 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 1), // 78 - 7f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 80 - 87 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 88 - 8f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 90 - 97 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 98 - 9f 
        BitPackage.Pack4Bits(4, 3, 3, 3, 3, 3, 3, 3), // a0 - a7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // a8 - af 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // b0 - b7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // b8 - bf 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // c0 - c7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // c8 - cf 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // d0 - d7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // d8 - df 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // e0 - e7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // e8 - ef 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // f0 - f7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 0) // f8 - ff 
    };

    private static readonly int[] Big5St =
    {
        BitPackage.Pack4Bits(Error, Start, Start, 3, Error, Error, Error, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, Error), //08-0f 
        BitPackage.Pack4Bits(Error, Start, Start, Start, Start, Start, Start, Start) //10-17 
    };

    private static readonly int[] Big5CharLenTable = { 0, 1, 1, 2, 0 };
    #endregion

    #region Constructor
    internal Big5SmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Big5Cls),
        5,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Big5St),
        Big5CharLenTable, "Big5")
    {
    }
    #endregion
}

internal class EucjpsmModel : SmModel
{
    #region Fields
    private static readonly int[] EucjpCls =
    {
        //BitPacket.Pack4bits(5,4,4,4,4,4,4,4),  // 00 - 07 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 00 - 07 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 5, 5), // 08 - 0f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 10 - 17 
        BitPackage.Pack4Bits(4, 4, 4, 5, 4, 4, 4, 4), // 18 - 1f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 20 - 27 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 28 - 2f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 30 - 37 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 38 - 3f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 40 - 47 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 48 - 4f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 50 - 57 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 58 - 5f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 60 - 67 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 68 - 6f 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 70 - 77 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // 78 - 7f 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // 80 - 87 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 1, 3), // 88 - 8f 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // 90 - 97 
        BitPackage.Pack4Bits(5, 5, 5, 5, 5, 5, 5, 5), // 98 - 9f 
        BitPackage.Pack4Bits(5, 2, 2, 2, 2, 2, 2, 2), // a0 - a7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // e0 - e7 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // e8 - ef 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // f0 - f7 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 5) // f8 - ff 
    };

    private static readonly int[] EucjpSt =
    {
        BitPackage.Pack4Bits(3, 4, 3, 5, Start, Error, Error, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, ItsMe, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, Start, Error, Start, Error, Error, Error), //10-17 
        BitPackage.Pack4Bits(Error, Error, Start, Error, Error, Error, 3, Error), //18-1f 
        BitPackage.Pack4Bits(3, Error, Error, Error, Start, Start, Start, Start) //20-27 
    };

    private static readonly int[] EucjpCharLenTable = { 2, 2, 2, 3, 1, 0 };
    #endregion

    #region Constructor
    internal EucjpsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EucjpCls),
        6,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EucjpSt),
        EucjpCharLenTable, "EUC-JP")
    {
    }
    #endregion
}

internal class EuckrsmModel : SmModel
{
    #region Fields
    private static readonly int[] EuckrCls =
    {
        //BitPacket.Pack4bits(0,1,1,1,1,1,1,1),  // 00 - 07 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 00 - 07 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 10 - 17 
        BitPackage.Pack4Bits(1, 1, 1, 0, 1, 1, 1, 1), // 18 - 1f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 20 - 27 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 28 - 2f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 30 - 37 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 38 - 3f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 40 - 47 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 48 - 4f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 50 - 57 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 58 - 5f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 60 - 67 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 68 - 6f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 70 - 77 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 78 - 7f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 80 - 87 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 88 - 8f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 90 - 97 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 98 - 9f 
        BitPackage.Pack4Bits(0, 2, 2, 2, 2, 2, 2, 2), // a0 - a7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 3, 3, 3), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 3, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e0 - e7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e8 - ef 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // f0 - f7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 0) // f8 - ff 
    };

    private static readonly int[] EuckrSt =
    {
        BitPackage.Pack4Bits(Error, Start, 3, Error, Error, Error, Error, Error), //00-07 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, Error, Error, Start, Start) //08-0f 
    };

    private static readonly int[] EuckrCharLenTable = { 0, 1, 2, 0 };
    #endregion

    #region Constructor
    internal EuckrsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EuckrCls),
        4,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EuckrSt),
        EuckrCharLenTable, "EUC-KR")
    {
    }
    #endregion
}

internal class EuctwsmModel : SmModel
{
    #region Fields
    private static readonly int[] EuctwCls =
    {
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 00 - 07 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 10 - 17 
        BitPackage.Pack4Bits(2, 2, 2, 0, 2, 2, 2, 2), // 18 - 1f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 20 - 27 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 28 - 2f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 30 - 37 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 38 - 3f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 40 - 47 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 48 - 4f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 50 - 57 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 58 - 5f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 60 - 67 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 68 - 6f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 70 - 77 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 78 - 7f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 80 - 87 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 6, 0), // 88 - 8f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 90 - 97 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 98 - 9f 
        BitPackage.Pack4Bits(0, 3, 4, 4, 4, 4, 4, 4), // a0 - a7 
        BitPackage.Pack4Bits(5, 5, 1, 1, 1, 1, 1, 1), // a8 - af 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // b0 - b7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // b8 - bf 
        BitPackage.Pack4Bits(1, 1, 3, 1, 3, 3, 3, 3), // c0 - c7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // c8 - cf 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // d0 - d7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // d8 - df 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // e0 - e7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // e8 - ef 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // f0 - f7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 0) // f8 - ff 
    };

    private static readonly int[] EuctwSt =
    {
        BitPackage.Pack4Bits(Error, Error, Start, 3, 3, 3, 4, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, Error, Start, Error), //10-17 
        BitPackage.Pack4Bits(Start, Start, Start, Error, Error, Error, Error, Error), //18-1f 
        BitPackage.Pack4Bits(5, Error, Error, Error, Start, Error, Start, Start), //20-27 
        BitPackage.Pack4Bits(Start, Error, Start, Start, Start, Start, Start, Start) //28-2f 
    };

    private static readonly int[] EuctwCharLenTable = { 0, 0, 1, 2, 2, 2, 3 };
    #endregion

    #region Constructor
    internal EuctwsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EuctwCls),
        7,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, EuctwSt),
        EuctwCharLenTable, "EUC-TW")
    {
    }
    #endregion
}

internal class SjissmModel : SmModel
{
    #region Fields
    private static readonly int[] SjisCls =
    {
        //BitPacket.Pack4bits(0,1,1,1,1,1,1,1),  // 00 - 07 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 00 - 07 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 10 - 17 
        BitPackage.Pack4Bits(1, 1, 1, 0, 1, 1, 1, 1), // 18 - 1f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 20 - 27 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 28 - 2f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 30 - 37 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 38 - 3f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 40 - 47 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 48 - 4f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 50 - 57 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 58 - 5f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 60 - 67 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 68 - 6f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 70 - 77 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 1), // 78 - 7f 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // 80 - 87 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // 88 - 8f 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // 90 - 97 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // 98 - 9f 
        //0xa0 is illegal in sjis encoding, but some pages does 
        //contain such byte. We need to be more error forgiven.
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a0 - a7     
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 3, 3, 3), // e0 - e7 
        BitPackage.Pack4Bits(3, 3, 3, 3, 3, 4, 4, 4), // e8 - ef 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 4, 4, 4), // f0 - f7 
        BitPackage.Pack4Bits(4, 4, 4, 4, 4, 0, 0, 0) // f8 - ff 
    };

    private static readonly int[] SjisSt =
    {
        BitPackage.Pack4Bits(Error, Start, Start, 3, Error, Error, Error, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, ItsMe, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, Error, Error, Start, Start, Start, Start) //10-17        
    };

    private static readonly int[] SjisCharLenTable = { 0, 1, 1, 2, 0, 0 };
    #endregion

    #region Constructor
    internal SjissmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, SjisCls),
        6,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, SjisSt),
        SjisCharLenTable, "Shift_JIS")
    {
    }
    #endregion
}