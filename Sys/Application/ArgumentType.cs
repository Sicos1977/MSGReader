// name       : ArgumentInfo.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Sys.Application
{
    [Flags]
    public enum ArgumentType
    {
        None = 0x00000000,
        Mandatory = 0x00000001,
        HasName = 0x00000002,
        ContainsValue = 0x00000004
    }
}