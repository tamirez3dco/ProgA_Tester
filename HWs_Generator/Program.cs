using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentsLib;

namespace HWs_Generator
{
    class Program
    {


        static void Main(string[] args)
        {

            Students students = new Students();
            foreach (int id in Students.students_dic.Keys)
            {
                HW0 hw0 = new HW0();
                int[] hw_args = hw0.get_random_args(id);
                hw0.Create_HW(hw_args, false);
                hw0.Create_DocFile(hw_args);
                Console.Clear();
                hw0.SaveArgs(hw_args);

            }


        }
    }
}
