using HWs_Generator;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StamConsoleTester
{
    class Program
    {
        static void Main(string[] args)
        {
            StudentsLib.Student tl = new StudentsLib.Student();
            tl.first_name = "תמיר";
            tl.last_name = "לוי";
            tl.id = 029046117;
            tl.email = "tamirlevi123@gmail.com";
            StudentsLib.Students.students_dic = new Dictionary<int, StudentsLib.Student>();
            StudentsLib.Students.students_dic[tl.id] = tl;
            //StudentsLib.Students
            FileInfo fin = new FileInfo(@"D:\Tamir\Netanya_ProgrammingA\2017\TempSolutions\HW3_Mine\ConsoleApplication1\ConsoleApplication1\bin\Debug\ConsoleApplication1.exe");
            HW4 hw3 = new HW4();
            Object[] argsTest = hw3.get_random_args(tl.id);
            hw3.Create_DocFile_By_Creators(argsTest, new List<Creators>());


            hw3.test_Hw(argsTest, fin);
        }
    }
}
