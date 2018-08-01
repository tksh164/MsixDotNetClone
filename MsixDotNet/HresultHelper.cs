using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsixDotNet
{
    internal static class HResultHelper
    {
        public static bool IsSucceeded(HRESULT hr)
        {
            return hr >= 0;
        }

        public static bool IsSucceeded(int hr)
        {
            return hr >= 0;
        }

        public static bool IsFailed(HRESULT hr)
        {
            return hr < 0;
        }

        public static bool IsFailed(int hr)
        {
            return hr < 0;
        }

        //public static string ToString(HRESULT hr)
        //{
        //    string tmp;

        //    tmp = new string('\u0000', 512);
        //    KERNEL32.FormatMessage(KERNEL32.FORMAT_MESSAGE.FORMAT_MESSAGE_FROM_SYSTEM,
        //                           IntPtr.Zero,
        //                           hr,
        //                           KERNEL32.GetUserDefaultLangID(),
        //                           tmp,
        //                           512,
        //                           IntPtr.Zero);

        //    return tmp.Substring(tmp.IndexOf('\u0000'));

        //}

        //public static string ToString(int hr)
        //{
        //    string tmp;

        //    tmp = new string('\u0000', 512);
        //    KERNEL32.FormatMessage(KERNEL32.FORMAT_MESSAGE.FORMAT_MESSAGE_FROM_SYSTEM,
        //                           IntPtr.Zero,
        //                           hr,
        //                           KERNEL32.GetUserDefaultLangID(),
        //                           tmp,
        //                           512,
        //                           IntPtr.Zero);

        //    return tmp.Substring(tmp.IndexOf('\u0000'));

        //}
    }
}
