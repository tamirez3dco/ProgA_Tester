using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using StudentsLib;

namespace NetExceptionCatcher
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Thread.Sleep(1000);
                //IntPtr ip = WindowsAPI.FindWindow("Microsoft .NET Framework");
                IntPtr ip = WindowsAPI.FindWindow("GUI2_Comparer");
                if (ip != IntPtr.Zero)
                {
                    Console.WriteLine("Found Exceptoin");
                    // Click Details
                    IntPtr buttin  = WindowsAPI.FindWindowEx(ip, IntPtr.Zero, null, "Details");
                }
                Console.WriteLine("continue");
            }
        }
    }
}
