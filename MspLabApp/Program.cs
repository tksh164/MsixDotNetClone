using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using VsInterop = Microsoft.VisualStudio.OLE.Interop;
using System.Diagnostics;
using System.IO;

namespace MspLabApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string extractDir = @"D:\HotfixDB\lab";
            bool includeExtension = true;

            string msiFilePath = @"D:\HotfixDB\lab\NDP45-KB3127229-x64\NDP45-KB3127229.msp";
            VsInterop.IStorage rootStorage = null;
            HRESULT hr = NativeMethods.StgOpenStorage(msiFilePath, null, STGM.STGM_READ | STGM.STGM_SHARE_EXCLUSIVE, null, 0, out rootStorage);

            if (HResultHelper.IsSucceeded(hr) && rootStorage != null)
            {
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
                    //err = NativeMethods.MsiDatabaseOpenView(msiDatabaseHandle, @"SELECT Name, Data FROM _Streams", out msiViewHandle);
                    err = NativeMethods.MsiDatabaseOpenView(msiDatabaseHandle, @"SELECT Name, Data FROM _Storages", out msiViewHandle);
                    //err = NativeMethods.MsiDatabaseOpenView(msiDatabaseHandle, @"SELECT Name, Data FROM _Tables", out msiViewHandle);
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

        //private static void saveStorage(VsInterop.IStorage storage, string saveDirectory, string storageName, string storageExtension)
        //{
        //    Trace.Assert(storage != null);
        //    Trace.Assert(!string.IsNullOrWhiteSpace(saveDirectory));

        //    VsInterop.IStorage sourceStorage = null;
        //    storage.OpenStorage(storageName, null, (uint)(STGM.STGM_READ | STGM.STGM_SHARE_EXCLUSIVE), IntPtr.Zero, 0, out sourceStorage);

        //    if (sourceStorage != null)
        //    {
        //        string savePath = saveDirectory + Path.DirectorySeparatorChar + storageName + (storageExtension == null ? string.Empty : storageExtension);
        //        Console.WriteLine(@"Extracting: ""{0}""", savePath);

        //        VsInterop.IStorage destinationStorage = null;
        //        HRESULT hr = NativeMethods.StgCreateDocfile(savePath, STGM.STGM_WRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_CREATE, 0, out destinationStorage);
        //        if (HResultHelper.IsSucceeded(hr) && destinationStorage != null)
        //        {
        //            sourceStorage.CopyTo(0, null, IntPtr.Zero, destinationStorage);
        //        }
        //        else
        //        {
        //            throw new Exception(string.Format(@"Failed: StgCreateDocfile(): Failed to create storage ""{0}""", savePath));
        //        }
        //    }
        //    else
        //    {
        //        throw new Exception(string.Format(@"Failed: OpenStorage(): Failed to save storage ""{0}""", storageName));
        //    }
        //}

        private static void saveStream(IntPtr msiRecordHandle, string saveDirectory, bool includeExtension)
        {
            Trace.Assert(msiRecordHandle != IntPtr.Zero);
            Trace.Assert(!string.IsNullOrWhiteSpace(saveDirectory));

            // Get the name of the stream
            string streamName = getString(msiRecordHandle, 1);

            Debug.WriteLine(streamName);

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

            // Save stream content.
            //string filePath = saveDirectory + Path.DirectorySeparatorChar + streamName + getStreamFileExtension(fileData);
            //using (FileStream fileStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
            //using (BinaryWriter fileWriter = new BinaryWriter(fileStream))
            //{
            //    fileWriter.Write(fileData);
            //}
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

    }
}
