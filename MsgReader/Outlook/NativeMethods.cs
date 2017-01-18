using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using FILETIME = System.Runtime.InteropServices.ComTypes.FILETIME;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
{
    public partial class Storage
    {
        #pragma warning disable 1591
        /// <summary>
        /// Contains all the used native methods
        /// </summary>
        public static class NativeMethods
        {
            #region STGM enum
            // ReSharper disable InconsistentNaming
            [Flags]
            public enum STGM
            {
                DIRECT = 0,
                FAILIFTHERE = 0,
                READ = 0,
                STGM_WRITE = 1,
                READWRITE = 2,
                SHARE_EXCLUSIVE = 0x10,
                SHARE_DENY_WRITE = 0x20,
                SHARE_DENY_READ = 0x30,
                SHARE_DENY_NONE = 0x40,
                CREATE = 0x1000,
                TRANSACTED = 0x10000,
                CONVERT = 0x20000,
                PRIORITY = 0x40000,
                NOSCRATCH = 0x100000,
                NOSNAPSHOT = 0x200000,
                DIRECT_SWMR = 0x400000,
                DELETEONRELEASE = 0x4000000,
                SIMPLE = 0x8000000,
            }
            // ReSharper restore InconsistentNaming
            #endregion

            #region DllImports
            [DllImport("ole32.DLL")]
            internal static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out ILockBytes ppLkbyt);

            [DllImport("ole32.DLL")]
            internal static extern int StgIsStorageILockBytes(ILockBytes plkbyt);

            [DllImport("ole32.DLL")]
            internal static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, STGM grfMode, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            internal static extern void StgOpenStorageOnILockBytes(ILockBytes plkbyt, IStorage pstgPriority, STGM grfMode,
                IntPtr snbExclude, uint reserved,
                out IStorage ppstgOpen);

            [DllImport("ole32.DLL")]
            internal static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string wcsName);

            [DllImport("ole32.DLL")]
            internal static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string wcsName,
                IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, int reserved,
                out IStorage ppstgOpen);
            #endregion

            #region CloneStorage
            /// <summary>
            /// This will clone the give <paramref name="source"/>storage
            /// </summary>
            /// <param name="source">The source to clone</param>
            /// <param name="closeSource">True to close the cloned source after cloning</param>
            /// <returns></returns>
            internal static IStorage CloneStorage(IStorage source, bool closeSource)
            {
                IStorage memoryStorage = null;
                ILockBytes memoryStorageBytes = null;
                try
                {
                    // Create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                    CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                    StgCreateDocfileOnILockBytes(memoryStorageBytes, STGM.CREATE | STGM.READWRITE | STGM.SHARE_EXCLUSIVE,
                        0, out memoryStorage);

                    // Copy the source storage into the new storage
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
            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
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
            /// <summary>
            /// Supports creation and management of structured storage objects which enable. hierarchical storage 
            /// of information within a single file
            /// </summary>
            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] STGM grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] STGM grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName,
                    [In, MarshalAs(UnmanagedType.U4)] STGM grfMode,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved1,
                    [In, MarshalAs(UnmanagedType.U4)] int reserved2);

                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority,
                    [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, IntPtr snbExclude,
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
        #pragma warning restore 1591
    }
}
