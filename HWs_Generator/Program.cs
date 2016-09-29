using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentsLib;
using System.IO;

namespace HWs_Generator
{
    class Program
    {


        static void Main(string[] args)
        {
            Students students = new Students();
/*
            HW2 hww = new HW2();
            int tid = 029046117;
            Object[] thw_args = hww.get_random_args(tid);
            hww.Create_HW(thw_args, false);
            hww.SaveArgs(thw_args);
            Console.ReadKey();
            return;
*/
            List<HW0> hws = new List<HW0>();
           // hws.Add(new HW0());
           // hws.Add(new HW1());
            hws.Add(new HW2());

            foreach (HW0 hw in hws)
            {

                foreach (int id in Students.students_dic.Keys)
                {
                    String docPath = hw.Students_Hws_dirs + "\\" + id.ToString() + ".docx";
                    if (File.Exists(docPath)) continue;
                    Object[] hw_args = hw.get_random_args(id);
                    Console.Clear();
                    hw.Create_HW(hw_args, false);
                    //hw0.Create_DocFile(hw_args);
                    hw.SaveArgs(hw_args);

                }

            }


        }
    }
}
