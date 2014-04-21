using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Contains all the used native methods
        /// </summary>
        public static class NativeMethods
        {
            #region Stgm enum
            [Flags]
            public enum Stgm
            {
                Direct = 0x00000000,
                Transacted = 0x00010000,
                Simple = 0x08000000,
                Read = 0x00000000,
                Write = 0x00000001,
                Readwrite = 0x00000002,
                ShareDenyNone = 0x00000040,
                ShareDenyRead = 0x00000030,
                ShareDenyWrite = 0x00000020,
                ShareExclusive = 0x00000010,
                Priority = 0x00040000,
                Deleteonrelease = 0x04000000,
                Noscratch = 0x00100000,
                Create = 0x00001000,
                Convert = 0x00020000,
                Failifthere = 0x00000000,
                Nosnapshot = 0x00200000,
                DirectSwmr = 0x00400000
            }
            #endregion

            #region DllImports
            [DllImport("ole32.DLL")]
            internal static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out ILockBytes ppLkbyt);

            [DllImport("ole32.DLL")]
            internal static extern int StgIsStorageILockBytes(ILockBytes plkbyt);

            [DllImport("ole32.DLL")]
            internal static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, Stgm grfMode, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            internal static extern void StgOpenStorageOnILockBytes(ILockBytes plkbyt, IStorage pstgPriority, Stgm grfMode,
                IntPtr snbExclude, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            internal static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string wcsName);

            [DllImport("ole32.DLL")]
            internal static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string wcsName,
                IStorage pstgPriority, Stgm grfMode, IntPtr snbExclude, int reserved,
                out IStorage ppstgOpen);
            #endregion

            #region CloneStorage
            internal static IStorage CloneStorage(IStorage source, bool closeSource)
            {
                IStorage memoryStorage = null;
                ILockBytes memoryStorageBytes = null;
                try
                {
                    //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    StgCreateDocfileOnILockBytes(memoryStorageBytes, Stgm.Create | Stgm.Readwrite | Stgm.ShareExclusive,
                        0, out memoryStorage);

                    //copy the source storage into the new storage
                    source.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                    memoryStorageBytes.Flush();
                    memoryStorage.Commit(0);

                    // Ensure memory is released
                    ReferenceManager.AddItem(memoryStorage);
                }
                catch
                {
                    if (memoryStorage != null)
                        Marshal.ReleaseComObject(memoryStorage);
                }
                finally
                {
                    if (memoryStorageBytes != null)
                        Marshal.ReleaseComObject(memoryStorageBytes);

                    if (closeSource)
                        Marshal.ReleaseComObject(source);
                }

                return memoryStorage;
            }
            #endregion

            #region IEnumSTATSTG
            [ComImport, Guid("0000000D-0000-0000-C000-000000000046"),
             InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IEnumSTATSTG
            {
                void Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] STATSTG[] rgelt, out uint pceltFetched);
                void Skip(uint celt);
                void Reset();

                [return: MarshalAs(UnmanagedType.Interface)]
                IEnumSTATSTG Clone();
            }
            #endregion

            #region ILockBytes
            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"),
             InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            internal interface ILockBytes
            {
                void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset,
                    [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                    [In, MarshalAs(UnmanagedType.U4)] int cb,
                    [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);

                void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset,
                    [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv,
                    [In, MarshalAs(UnmanagedType.U4)] int cb,
                    [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);

                void Flush();
                void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);

                void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset,
                    [In, MarshalAs(UnmanagedType.U8)] long cb,
                    [In, MarshalAs(UnmanagedType.U4)] int dwLockType);

                void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset,
                    [In, MarshalAs(UnmanagedType.U8)] long cb,
                    [In, MarshalAs(UnmanagedType.U4)] int dwLockType);

                void Stat([Out] out STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
            }
            #endregion

            #region IStorage
            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority,
                    [In, MarshalAs(UnmanagedType.U4)] Stgm grfMode, IntPtr snbExclude,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved);

                void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIidExclude,
                    IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);

                void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest,
                    [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName,
                    [In, MarshalAs(UnmanagedType.U4)] int grfFlags);

                void Commit(int grfCommitFlags);
                void Revert();

                void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved3,
                    [MarshalAs(UnmanagedType.Interface)] out IEnumSTATSTG ppVal);

                void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);

                void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName,
                    [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);

                void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] FILETIME pctime,
                    [In] FILETIME patime, [In] FILETIME pmtime);

                void SetClass([In] ref Guid clsid);
                void SetStateBits(int grfStateBits, int grfMask);
                void Stat([Out] out STATSTG pStatStg, int grfStatFlag);
            }
            #endregion
        }
    }
}
