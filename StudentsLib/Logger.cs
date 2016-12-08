using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace StudentsLib
{
    public class Logger
    {
        public static void Log(String str, String hw_name, Student s = null)
        {
            DateTime now = DateTime.Now;
            String fileName = String.Format(@"D:\Tamir\CodeTesterLogs\Log_{0}_{1}_{2}.txt", now.Day, now.Month, now.Year);
            String msg = String.Format("{0}:{1}:{2}({3})-{4}",now.Hour,now.Minute,now.Second,hw_name,str);
            if (s != null) msg += " for " + s.ToString();

            using (StreamWriter sw = new StreamWriter(fileName, true))
            {
                sw.WriteLine(msg);
            }

        }
    }
}
