using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentsLib;
using System.IO;

namespace HWs_Generator
{
    // TODO: catch returning errors and handle them differently

    class Program
    {


        static void Main(string[] args)
        {
            /*
                        Student tl = new Student();
                        tl.first_name = "תמיר";
                        tl.last_name = "לוי";
                        tl.id = 029046117;
                        tl.email = "tamirlevi123@gmail.com";
                        Students.students_dic = new Dictionary<int, StudentsLib.Student>();
                        Students.students_dic[tl.id] = tl;


                        HW3 hww = new HW3();
                        int tid = 029046117;
                        Object[] myargs = hww.LoadArgs(tid);
                        String studentOutput = File.ReadAllText(@"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\GeneratedInput\student_output.txt");
                        String benchmarkOutput = File.ReadAllText(@"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\GeneratedInput\benchmark_output.txt");
                        RunResults rr = hww.Test_HW(myargs, @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\Hw3_Arrays_Mine.exe");
                        return;
            */
            /*
                        Object[] thw_args = hww.get_random_args(tid);


                        hww.Create_HW(thw_args, false);
                        hww.SaveArgs(thw_args);
                        //Console.ReadKey();
                        return;
            */
            Students students = new Students();
            List<HW0> hws = new List<HW0>();
             hws.Add(new HW0());
             hws.Add(new HW1());
            //hws.Add(new HW2());
            //hws.Add(new HW3());

            foreach (HW0 hw in hws)
            {

                foreach (int id in Students.students_dic.Keys)
                {
                    String docPath = hw.Students_Hws_dirs + "\\" + id.ToString() + ".docx";
                    if (File.Exists(docPath)) continue;
                    Object[] hw_args = hw.get_random_args(id);
                    Console.Clear();
                    hw.Create_HW(hw_args, false);
                    hw.SaveArgs(hw_args);

                }

            }


        }
    }
}
