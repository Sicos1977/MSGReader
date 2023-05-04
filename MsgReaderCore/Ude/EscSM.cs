/* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1/GPL 2.0/LGPL 2.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is Mozilla Universal charset detector code.
 *
 * The Initial Developer of the Original Code is
 * Netscape Communications Corporation.
 * Portions created by the Initial Developer are Copyright (C) 2001
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *          Kohei TAKETA <k-tak@void.in> (Java port)
 *          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
 *
 * Refactoring to the code done by Kees van Spelde so that it works in this project
 * Copyright (c) 2023 Magic-Sessions. (www.magic-sessions.com)
 *
 * Alternatively, the contents of this file may be used under the terms of
 * either the GNU General Public License Version 2 or later (the "GPL"), or
 * the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
 * in which case the provisions of the GPL or the LGPL are applicable instead
 * of those above. If you wish to allow use of your version of this file only
 * under the terms of either the GPL or the LGPL, and not to allow others to
 * use your version of this file under the terms of the MPL, indicate your
 * decision by deleting the provisions above and replace them with the notice
 * and other provisions required by the GPL or the LGPL. If you do not delete
 * the provisions above, a recipient may use your version of this file under
 * the terms of any one of the MPL, the GPL or the LGPL.
 *
 * ***** END LICENSE BLOCK ***** */

namespace MsgReader.Ude;

/// <summary>
/// Escaped charsets state machines
/// </summary>
internal class HzsmModel : SmModel
{
    #region Fields
    private static readonly int[] HzCls =
    {
        BitPackage.Pack4Bits(1, 0, 0, 0, 0, 0, 0, 0), // 00 - 07 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 10 - 17 
        BitPackage.Pack4Bits(0, 0, 0, 1, 0, 0, 0, 0), // 18 - 1f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 20 - 27 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 28 - 2f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 30 - 37 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 38 - 3f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 40 - 47 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 48 - 4f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 50 - 57 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 58 - 5f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 60 - 67 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 68 - 6f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 70 - 77 
        BitPackage.Pack4Bits(0, 0, 0, 4, 0, 5, 2, 0), // 78 - 7f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 80 - 87 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 88 - 8f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 90 - 97 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // 98 - 9f 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // a0 - a7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // a8 - af 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // b0 - b7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // b8 - bf 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // c0 - c7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // c8 - cf 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // d0 - d7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // d8 - df 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // e0 - e7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // e8 - ef 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1), // f0 - f7 
        BitPackage.Pack4Bits(1, 1, 1, 1, 1, 1, 1, 1) // f8 - ff 
    };

    private static readonly int[] HzSt =
    {
        BitPackage.Pack4Bits(Start, Error, 3, Start, Start, Start, Error, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, ItsMe, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, Error, Error, Start, Start, 4, Error), //10-17 
        BitPackage.Pack4Bits(5, Error, 6, Error, 5, 5, 4, Error), //18-1f 
        BitPackage.Pack4Bits(4, Error, 4, 4, 4, Error, 4, Error), //20-27 
        BitPackage.Pack4Bits(4, ItsMe, Start, Start, Start, Start, Start, Start) //28-2f 
    };

    private static readonly int[] HzCharLenTable = { 0, 0, 0, 0, 0, 0 };
    #endregion

    #region Constructor
    public HzsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, HzCls),
        6,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, HzSt),
        HzCharLenTable, "HZ-GB-2312")
    {
    }
    #endregion
}

internal class Iso2022CnsmModel : SmModel
{
    #region Fields
    private static readonly int[] Iso2022CnCls =
    {
        BitPackage.Pack4Bits(2, 0, 0, 0, 0, 0, 0, 0), // 00 - 07 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 10 - 17 
        BitPackage.Pack4Bits(0, 0, 0, 1, 0, 0, 0, 0), // 18 - 1f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 20 - 27 
        BitPackage.Pack4Bits(0, 3, 0, 0, 0, 0, 0, 0), // 28 - 2f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 30 - 37 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 38 - 3f 
        BitPackage.Pack4Bits(0, 0, 0, 4, 0, 0, 0, 0), // 40 - 47 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 48 - 4f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 50 - 57 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 58 - 5f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 60 - 67 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 68 - 6f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 70 - 77 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 78 - 7f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 80 - 87 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 88 - 8f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 90 - 97 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 98 - 9f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a0 - a7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e0 - e7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e8 - ef 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // f0 - f7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2) // f8 - ff 
    };

    private static readonly int[] Iso2022CnSt =
    {
        BitPackage.Pack4Bits(Start, 3, Error, Start, Start, Start, Start, Start), //00-07 
        BitPackage.Pack4Bits(Start, Error, Error, Error, Error, Error, Error, Error), //08-0f 
        BitPackage.Pack4Bits(Error, Error, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe), //10-17 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, Error, Error, Error, 4, Error), //18-1f 
        BitPackage.Pack4Bits(Error, Error, Error, ItsMe, Error, Error, Error, Error), //20-27 
        BitPackage.Pack4Bits(5, 6, Error, Error, Error, Error, Error, Error), //28-2f 
        BitPackage.Pack4Bits(Error, Error, Error, ItsMe, Error, Error, Error, Error), //30-37 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, ItsMe, Error, Start) //38-3f 
    };

    private static readonly int[] Iso2022CnCharLenTable = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
    #endregion

    #region Constructor
    internal Iso2022CnsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022CnCls),
        9,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022CnSt),
        Iso2022CnCharLenTable, "ISO-2022-CN")
    {
    }
    #endregion
}

internal class Iso2022JpsmModel : SmModel
{
    #region Fields
    private static readonly int[] Iso2022JpCls =
    {
        BitPackage.Pack4Bits(2, 0, 0, 0, 0, 0, 0, 0), // 00 - 07 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 2, 2), // 08 - 0f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 10 - 17 
        BitPackage.Pack4Bits(0, 0, 0, 1, 0, 0, 0, 0), // 18 - 1f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 7, 0, 0, 0), // 20 - 27 
        BitPackage.Pack4Bits(3, 0, 0, 0, 0, 0, 0, 0), // 28 - 2f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 30 - 37 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 38 - 3f 
        BitPackage.Pack4Bits(6, 0, 4, 0, 8, 0, 0, 0), // 40 - 47 
        BitPackage.Pack4Bits(0, 9, 5, 0, 0, 0, 0, 0), // 48 - 4f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 50 - 57 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 58 - 5f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 60 - 67 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 68 - 6f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 70 - 77 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 78 - 7f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 80 - 87 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 88 - 8f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 90 - 97 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 98 - 9f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a0 - a7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e0 - e7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e8 - ef 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // f0 - f7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2) // f8 - ff 
    };

    private static readonly int[] Iso2022JpSt =
    {
        BitPackage.Pack4Bits(Start, 3, Error, Start, Start, Start, Start, Start), //00-07 
        BitPackage.Pack4Bits(Start, Start, Error, Error, Error, Error, Error, Error), //08-0f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, ItsMe, ItsMe, ItsMe), //10-17 
        BitPackage.Pack4Bits(ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, ItsMe, Error, Error), //18-1f 
        BitPackage.Pack4Bits(Error, 5, Error, Error, Error, 4, Error, Error), //20-27 
        BitPackage.Pack4Bits(Error, Error, Error, 6, ItsMe, Error, ItsMe, Error), //28-2f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, Error, Error, ItsMe, ItsMe), //30-37 
        BitPackage.Pack4Bits(Error, Error, Error, ItsMe, Error, Error, Error, Error), //38-3f 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, Error, Start, Start) //40-47 
    };

    private static readonly int[] Iso2022JpCharLenTable = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
    #endregion

    #region Constructor
    public Iso2022JpsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022JpCls),
        10,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022JpSt),
        Iso2022JpCharLenTable, "ISO-2022-JP")
    {
    }
    #endregion
}

internal class Iso2022KrsmModel : SmModel
{
    #region Fields
    private static readonly int[] Iso2022KrCls =
    {
        BitPackage.Pack4Bits(2, 0, 0, 0, 0, 0, 0, 0), // 00 - 07 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 08 - 0f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 10 - 17 
        BitPackage.Pack4Bits(0, 0, 0, 1, 0, 0, 0, 0), // 18 - 1f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 3, 0, 0, 0), // 20 - 27 
        BitPackage.Pack4Bits(0, 4, 0, 0, 0, 0, 0, 0), // 28 - 2f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 30 - 37 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 38 - 3f 
        BitPackage.Pack4Bits(0, 0, 0, 5, 0, 0, 0, 0), // 40 - 47 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 48 - 4f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 50 - 57 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 58 - 5f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 60 - 67 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 68 - 6f 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 70 - 77 
        BitPackage.Pack4Bits(0, 0, 0, 0, 0, 0, 0, 0), // 78 - 7f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 80 - 87 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 88 - 8f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 90 - 97 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // 98 - 9f 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a0 - a7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // a8 - af 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b0 - b7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // b8 - bf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c0 - c7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // c8 - cf 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d0 - d7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // d8 - df 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e0 - e7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // e8 - ef 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2), // f0 - f7 
        BitPackage.Pack4Bits(2, 2, 2, 2, 2, 2, 2, 2) // f8 - ff 
    };

    private static readonly int[] Iso2022KrSt =
    {
        BitPackage.Pack4Bits(Start, 3, Error, Start, Start, Start, Error, Error), //00-07 
        BitPackage.Pack4Bits(Error, Error, Error, Error, ItsMe, ItsMe, ItsMe, ItsMe), //08-0f 
        BitPackage.Pack4Bits(ItsMe, ItsMe, Error, Error, Error, 4, Error, Error), //10-17 
        BitPackage.Pack4Bits(Error, Error, Error, Error, 5, Error, Error, Error), //18-1f 
        BitPackage.Pack4Bits(Error, Error, Error, ItsMe, Start, Start, Start, Start) //20-27 
    };

    private static readonly int[] Iso2022KrCharLenTable = { 0, 0, 0, 0, 0, 0 };
    #endregion

    #region Constructor
    public Iso2022KrsmModel() : base(
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022KrCls),
        6,
        new BitPackage(BitPackage.IndexShift4Bits,
            BitPackage.ShiftMask4Bits,
            BitPackage.BitShift4Bits,
            BitPackage.UnitMask4Bits, Iso2022KrSt),
        Iso2022KrCharLenTable, "ISO-2022-KR")
    {
    }
    #endregion
}