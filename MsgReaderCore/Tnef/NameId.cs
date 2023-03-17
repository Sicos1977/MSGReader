//
// TnefNameId.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
//
// Refactoring to the code done by Kees van Spelde so that it works in this project
// Copyright (c) 2023 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using MsgReader.Tnef.Enums;
// ReSharper disable UnusedMember.Global

namespace MsgReader.Tnef;

/// <summary>
///     A TNEF name identifier.
/// </summary>
/// <remarks>
///     A TNEF name identifier.
/// </remarks>
internal readonly struct NameId : IEquatable<NameId>
{
    #region Fields
    private readonly Guid _guid;
    #endregion

    #region Properties
    /// <summary>
    ///     Get the property set GUID.
    /// </summary>
    /// <remarks>
    ///     Gets the property set GUID.
    /// </remarks>
    /// <value>The property set GUID.</value>
    public Guid PropertySetGuid => _guid;

    /// <summary>
    ///     Get the kind of TNEF name identifier.
    /// </summary>
    /// <remarks>
    ///     Gets the kind of TNEF name identifier.
    /// </remarks>
    /// <value>The kind of identifier.</value>
    public NameIdKind Kind { get; }

    /// <summary>
    ///     Get the name, if available.
    /// </summary>
    /// <remarks>
    ///     If the <see cref="Kind" /> is <see cref="NameIdKind.Name" />, then this property will be available.
    /// </remarks>
    /// <value>The name.</value>
    public string Name { get; }

    /// <summary>
    ///     Get the identifier, if available.
    /// </summary>
    /// <remarks>
    ///     If the <see cref="Kind" /> is <see cref="NameIdKind.Id" />, then this property will be available.
    /// </remarks>
    /// <value>The identifier.</value>
    public int Id { get; }
    #endregion

    #region Constructors
    /// <summary>
    ///     Initialize a new instance of the <see cref="NameId" /> struct.
    /// </summary>
    /// <remarks>
    ///     Creates a new <see cref="NameId" /> with the specified integer identifier.
    /// </remarks>
    /// <param name="propertySetGuid">The property set GUID.</param>
    /// <param name="id">The identifier.</param>
    public NameId(Guid propertySetGuid, int id)
    {
        Kind = NameIdKind.Id;
        _guid = propertySetGuid;
        Id = id;
        Name = null;
    }

    /// <summary>
    ///     Initialize a new instance of the <see cref="NameId" /> struct.
    /// </summary>
    /// <remarks>
    ///     Creates a new <see cref="NameId" /> with the specified string identifier.
    /// </remarks>
    /// <param name="propertySetGuid">The property set GUID.</param>
    /// <param name="name">The name.</param>
    public NameId(Guid propertySetGuid, string name)
    {
        Kind = NameIdKind.Name;
        _guid = propertySetGuid;
        Name = name;
        Id = 0;
    }
    #endregion

    #region GetHashCode
    /// <summary>
    ///     Serves as a hash function for a <see cref="NameId" /> object.
    /// </summary>
    /// <remarks>
    ///     Serves as a hash function for a <see cref="NameId" /> object.
    /// </remarks>
    /// <returns>
    ///     A hash code for this instance that is suitable for use in hashing algorithms
    ///     and data structures such as a hash table.
    /// </returns>
    public override int GetHashCode()
    {
        var hash = Kind == NameIdKind.Id ? Id : Name.GetHashCode();

        return Kind.GetHashCode() ^ _guid.GetHashCode() ^ hash;
    }
    #endregion

    #region Equals
    /// <summary>
    ///     Determine whether the specified <see cref="object" /> is equal to the current <see cref="NameId" />.
    /// </summary>
    /// <remarks>
    ///     Determines whether the specified <see cref="object" /> is equal to the current <see cref="NameId" />.
    /// </remarks>
    /// <param name="obj">The <see cref="object" /> to compare with the current <see cref="NameId" />.</param>
    /// <returns>
    ///     <c>true</c> if the specified <see cref="object" /> is equal to the current
    ///     <see cref="NameId" />; otherwise, <c>false</c>.
    /// </returns>
    public override bool Equals(object obj)
    {
        return obj is NameId other && Equals(other);
    }

    /// <summary>
    ///     Determine whether the specified <see cref="NameId" /> is equal to the current <see cref="NameId" />.
    /// </summary>
    /// <remarks>
    ///     Compares two TNEF name identifiers to determine if they are identical or not.
    /// </remarks>
    /// <param name="other">The <see cref="NameId" /> to compare with the current <see cref="NameId" />.</param>
    /// <returns>
    ///     <c>true</c> if the specified <see cref="NameId" /> is equal to the current
    ///     <see cref="NameId" />; otherwise, <c>false</c>.
    /// </returns>
    public bool Equals(NameId other)
    {
        if (Kind != other.Kind || _guid != other._guid)
            return false;

        return Kind is NameIdKind.Id ? other.Id == Id : other.Name == Name;
    }
    #endregion
}