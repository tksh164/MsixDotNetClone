using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using VsInterop = Microsoft.VisualStudio.OLE.Interop;
using System.IO;
using System.Diagnostics;

namespace MsixDotNet
{
    class Program
    {
        static void Main(string[] args)
        {
            string extractDir = @"F:\Takatano\Desktop\experiment\storages";
            bool includeExtension = true;

            string msiFilePath = @"F:\Takatano\Projects\MsixEx\NDP45-KB3037581.msp";
            VsInterop.IStorage rootStorage = null;
            HRESULT hr = NativeMethods.StgOpenStorage(msiFilePath, null, STGM.STGM_READ | STGM.STGM_SHARE_EXCLUSIVE, null, 0, out rootStorage);

            if (HResultHelper.IsSucceeded(hr) && rootStorage != null)
            {
                //
                // もしかすると、この辺関係ないかも
                // (構造を維持して展開できれば、話は違うかも)
                //

                //VsInterop.IEnumSTATSTG enumStatStg = null;
                //rootStorage.EnumElements(0, IntPtr.Zero, 0, out enumStatStg);
                //if (enumStatStg != null)
                //{
                //    while (true)
                //    {
                //        VsInterop.STATSTG[] stg = new VsInterop.STATSTG[1];
                //        uint fetched;
                //        hr = (HRESULT)enumStatStg.Next(1, stg, out fetched);
                //        if (hr == HRESULT.S_FALSE)
                //        {
                //            break;
                //        }
                //        if (hr != HRESULT.S_OK)
                //        {
                //            throw new Exception(@"Failed: IEnumSTATSTG.Next(): Failed to enumerate next storage status");
                //        }

                //        Console.WriteLine(@"Name:{0}, Type:{1}", stg[0].pwcsName, stg[0].type);

                //        if ((VsInterop.STGTY)stg[0].type == VsInterop.STGTY.STGTY_STORAGE)
                //        {
                //            saveStorage(rootStorage, extractDir, stg[0].pwcsName, includeExtension ? @".mst" : null);
                //        }

                //        //if (stg.pwcsName)
                //        //{
                //        //    CoTaskMemFree(stg.pwcsName);    // これ必要なのかも, Marshal.FreeCoTaskMem(), string になっているし、不要かな。
                //        //    stg.pwcsName = NULL;
                //        //}
                //    }

                //    Marshal.ReleaseComObject(enumStatStg);
                //}

                //
                // A Database and Patch Example
                // https://msdn.microsoft.com/en-us/library/aa367813(v=vs.85).aspx
                //
                // http://www.pinvoke.net/default.aspx/msi.msirecordsetstring
                // http://blogs.msdn.com/b/heaths/archive/2006/03/31/566288.aspx
                //
                MsiDbOpenPersist msiDbOpenPersist = MsiDbOpenPersist.MSIDBOPEN_READONLY;
                if (isPatch(rootStorage))
                {
                    msiDbOpenPersist = MsiDbOpenPersist.MSIDBOPEN_READONLY | MsiDbOpenPersist.MSIDBOPEN_PATCHFILE;
                }

                // Release COM object of root storage.
                Marshal.ReleaseComObject(rootStorage);

                IntPtr msiDatabaseHandle = IntPtr.Zero;
                uint err = NativeMethods.MsiOpenDatabase(msiFilePath, new IntPtr((uint)msiDbOpenPersist), out msiDatabaseHandle);
                if (err == WinError.ERROR_SUCCESS)
                {
                    IntPtr msiViewHandle = IntPtr.Zero;
                    err = NativeMethods.MsiDatabaseOpenView(msiDatabaseHandle, @"SELECT Name, Data FROM _Streams", out msiViewHandle);
                    if (err == WinError.ERROR_SUCCESS)
                    {
                        err = NativeMethods.MsiViewExecute(msiViewHandle, IntPtr.Zero);
                        if (err == WinError.ERROR_SUCCESS)
                        {
                            while (true)
                            {
                                IntPtr msiRecordHandle = IntPtr.Zero;
                                err = NativeMethods.MsiViewFetch(msiViewHandle, out msiRecordHandle);
                                if (err != WinError.ERROR_SUCCESS) break;

                                saveStream(msiRecordHandle, extractDir, includeExtension);

                                NativeMethods.MsiCloseHandle(msiRecordHandle);
                            }
                        }
                    }

                    // Close MSI view handle.
                    if (msiViewHandle != IntPtr.Zero) NativeMethods.MsiCloseHandle(msiViewHandle);
                }

                // Close MSI database handle.
                if (msiDatabaseHandle != IntPtr.Zero) NativeMethods.MsiCloseHandle(msiDatabaseHandle);
            }
            else
            {
                throw new Exception(string.Format(@"Failed: StgOpenStorage(): Failed to open root storage ""{0}""", msiFilePath));
            }
        }

        private static bool isPatch(VsInterop.IStorage storage)
        {
            Trace.Assert(storage != null);

            Guid CLSID_MsiPatch = new Guid(@"{000c1086-0000-0000-c000-000000000046}");

            VsInterop.STATSTG[] stg = new VsInterop.STATSTG[1];
            storage.Stat(stg, (uint)VsInterop.STATFLAG.STATFLAG_NONAME);
            return stg[0].clsid == CLSID_MsiPatch;
        }

        private static void saveStorage(VsInterop.IStorage storage, string saveDirectory, string storageName, string storageExtension)
        {
            Trace.Assert(storage != null);
            Trace.Assert(!string.IsNullOrWhiteSpace(saveDirectory));

            VsInterop.IStorage sourceStorage = null;
            storage.OpenStorage(storageName, null, (uint)(STGM.STGM_READ | STGM.STGM_SHARE_EXCLUSIVE), IntPtr.Zero, 0, out sourceStorage);

            if (sourceStorage != null)
            {
                string savePath = saveDirectory + Path.DirectorySeparatorChar + storageName + (storageExtension == null ? string.Empty : storageExtension);
                Console.WriteLine(@"Extracting: ""{0}""", savePath);

                VsInterop.IStorage destinationStorage = null;
                HRESULT hr = NativeMethods.StgCreateDocfile(savePath, STGM.STGM_WRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_CREATE, 0, out destinationStorage);
                if (HResultHelper.IsSucceeded(hr) && destinationStorage != null)
                {
                    sourceStorage.CopyTo(0, null, IntPtr.Zero, destinationStorage);
                }
                else
                {
                    throw new Exception(string.Format(@"Failed: StgCreateDocfile(): Failed to create storage ""{0}""", savePath));
                }
            }
            else
            {
                throw new Exception(string.Format(@"Failed: OpenStorage(): Failed to save storage ""{0}""", storageName));
            }
        }

        private static void saveStream(IntPtr msiRecordHandle, string saveDirectory, bool includeExtension)
        {
            Trace.Assert(msiRecordHandle != IntPtr.Zero);
            Trace.Assert(!string.IsNullOrWhiteSpace(saveDirectory));

            // Get the name of the stream
            string streamName = getString(msiRecordHandle, 1);

            // For \005SummaryInformation and \005DigitalSignature streams, display only the name.
            if (streamName.StartsWith("\u0005")) return;

            byte[] fileData;
            using (MemoryStream stream = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(stream))
            {
                while (true)
                {
                    byte[] buffer = new byte[4098];
                    int bufferLength = buffer.Length;

                    uint err = NativeMethods.MsiRecordReadStream(msiRecordHandle, 2, buffer, ref bufferLength);
                    if (err != WinError.ERROR_SUCCESS)
                    {
                        throw new Exception(@"Could not read from stream.");
                    }
                    if (bufferLength == 0) break;

                    writer.Write(buffer);
                }

                fileData = stream.ToArray();
            }

            string filePath = saveDirectory + Path.DirectorySeparatorChar + streamName + getStreamFileExtension(fileData);
            using (FileStream fileStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
            using (BinaryWriter fileWriter = new BinaryWriter(fileStream))
            {
                fileWriter.Write(fileData);
            }
        }

        private static string getString(IntPtr msiRecordHandle, uint field)
        {
            Trace.Assert(msiRecordHandle != IntPtr.Zero);
            Trace.Assert(field > 0);

            int bufferLength = 0;
            uint err = NativeMethods.MsiRecordGetString(msiRecordHandle, field, new StringBuilder(0), ref bufferLength);
            if (err != WinError.ERROR_MORE_DATA)
            {
                throw new Exception(@"Failed: MsiRecordGetString(): First time.");
            }

            bufferLength++;
            StringBuilder buffer = new StringBuilder(bufferLength);
            err = NativeMethods.MsiRecordGetString(msiRecordHandle, field, buffer, ref bufferLength);
            if (err != WinError.ERROR_SUCCESS)
            {
                throw new Exception(@"Failed: MsiRecordGetString(): Second time.");
            }

            return buffer.ToString();
        }

        private static string getStreamFileExtension(byte[] fileData)
        {
            byte[] signature = new byte[] { 0x4d, 0x53, 0x43, 0x46 }; // .cab, MSCF
            if (compareArray<byte>(fileData, signature, signature.Length)) return @".cab";

            signature = new byte[] { 0x4d, 0x5a }; // .exe/.dll, MZ
            if (compareArray<byte>(fileData, signature, signature.Length)) return @".dll";

            signature = new byte[] { 0x0, 0x0, 0x1, 0x0 }; // .ico
            if (compareArray<byte>(fileData, signature, signature.Length)) return @".ico";

            signature = new byte[] { 0x42, 0x4d }; // .bmp, BM
            if (compareArray<byte>(fileData, signature, signature.Length)) return @".bmp";

            signature = new byte[] { 0x47, 0x49, 0x46 }; // .gif, GIF
            if (compareArray<byte>(fileData, signature, signature.Length)) return @".gif";

            return string.Empty;
        }

        private static bool compareArray<T>(T[] a, T[] b, int length)
        {
            Trace.Assert(a.Length >= length);
            Trace.Assert(b.Length >= length);

            EqualityComparer<T> comparer = EqualityComparer<T>.Default;
            for (int i = 0; i < length; i++)
            {
                if (!comparer.Equals(a[i], b[i])) return false;
            }
            return true;
        }

    }
}
