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


        static String ClassName;
        static void Main(string[] args)
        {

//            GUI3_Form ff = new GUI3_Form();
//            ff.ShowDialog();
//            return;

            String excel_file_path = args[0];
            ClassName = Environment.GetCommandLineArgs()[1];

            Students students;
            switch (ClassName)
            {
                case "ProgrammingA_2017_Summer":
                    HW0.Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs_Summer";
                    students = new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\students_name_id_Shana_B.xlsx");                  
                    break;
                case "EDP_2017":
                    GUI1.Students_All_Hws_dirs = @"D:\Tamir\Netanya_Desktop_App\2017\Students_HWs";
                    students = new Students(@"D:\Tamir\Netanya_Desktop_App\2017\Shana_B_2017.xlsx");
                    break;
                case "ProgrammingA_2017":
                    students = new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\Programming_A_Semester_A_2017.xlsx");
                    GUI1.Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs";
                    break;
                case "Java1_2017_Highschool":
                    students = new Students(@"D:\Tamir\Netanya_Java_1\2017\Highschool\Highschool_Class.xlsx");
                    GUI1.Students_All_Hws_dirs = @"D:\Tamir\Netanya_Java_1\2017\Highschool\Students_HWs";
                    break;
            }

            foreach (Student stud in Students.students_dic.Values)
            {
                GUI3 hww = new GUI3();
                if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                Object[] myargs = hww.get_random_args(stud.id);
                hww.Create_DocFile_By_Creators(myargs,null);
                hww.SaveArgs(myargs);
            }
            return;


            /*
                        Student tl = new Student();
                        tl.first_name = "תמיר";
                        tl.last_name = "לוי";
                        tl.id = 029046117;
                        tl.email = "tamirlevi123@gmail.com";
                        Students.students_dic = new Dictionary<int, StudentsLib.Student>();
                        Students.students_dic[tl.id] = tl;

                        GUI3 hww = new GUI3();
                        Object[] myargs = hww.get_random_args(tl.id);
                        hww.Create_DocFile(myargs);
                        return;
            */
            /*
                        foreach (Student stud in Students.students_dic.Values)
                        {
                            GUI1 hww = new GUI1();
                            if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                            Object[] myargs = hww.get_random_args(stud.id);
                            hww.Create_DocFile(myargs);
                            hww.SaveArgs(myargs);
                        }
                        return;
            */
            foreach (Student stud in Students.students_dic.Values)
            {
                HW4 hww = new HW4();
                if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                Object[] myargs = hww.get_random_args(stud.id);
                //hww.Create_DocFile_By_Creators(myargs, new List<Creators>());
                hww.Create_HW(myargs, false);
                hww.SaveArgs(myargs);
                Console.Clear();
            }

            return;

            //Students allStudents = new Students(@"D:\Tamir\Netanya_Desktop_App\2017\Shana_B_2017.xlsx");

            /*
                        Student tl = new Student();
                                    tl.first_name = "תמיר";
                                    tl.last_name = "לוי";
                                    tl.id = 029046117;
                                    tl.email = "tamirlevi123@gmail.com";
                                    Students.students_dic = new Dictionary<int, StudentsLib.Student>();
                                    Students.students_dic[tl.id] = tl;
            */

            /*
                        // fix all students shitty images...
                        Students allStudents = new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\Programming_A_Semester_A_2017.xlsx");
                        int[] shittyImagesIds = { 332270925, 315815522, 311367213, 309705119, 204578249, 204201602 };// 203599014 };
                        foreach (int id in shittyImagesIds)
                        {
                            Student stud = Students.students_dic[id];
                            HW0 hww = new HW0();
                            //if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                            //hww.Create_DocFile_By_Creators(myargs, new List<Creators>());
                            Object[] myargs = hww.LoadArgs(id);
                            hww.SaveArgs(myargs);
                            hww.Create_HW(myargs, false);
                        }

                        return;
            */

/*
            foreach (Student stud in Students.students_dic.Values)
            {
                HW0 hww = new HW0();
                if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                Object[] myargs = hww.get_random_args(stud.id);
                //hww.Create_DocFile_By_Creators(myargs, new List<Creators>());
                hww.Create_HW(myargs, false);
                hww.SaveArgs(myargs);
            }
*/

/*
            foreach (Student stud in Students.students_dic.Values)
            {
                HW0 hww = new HW0();
                if (File.Exists(hww.Students_Hws_dirs + "\\" + stud.id.ToString() + ".docx")) continue;
                Object[] myargs = hww.get_random_args(stud.id);
                //hww.Create_DocFile_By_Creators(myargs, new List<Creators>());
                hww.Create_HW(myargs, false);
                hww.SaveArgs(myargs);
            }
*/
            return;


            //Object[] myargs = hww.LoadArgs(tid);
            //                        String studentOutput = File.ReadAllText(@"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\GeneratedInput\student_output.txt");
            //                        String benchmarkOutput = File.ReadAllText(@"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\GeneratedInput\benchmark_output.txt");
            //                        RunResults rr = hww.Test_HW(myargs, @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\29046117\4_10_2016_22_05_extracted\Hw3_Arrays_Mine\bin\Debug\Hw3_Arrays_Mine.exe");
            //RunResults rr = hww.Test_HW(myargs, @"D:\Tamir\Netanya_Desktop_App\2017\Students_Submissions\GUI1\029046117\GUI1_Mine\GUI1_Mine\bin\Debug\GUI1_Mine.exe");
            //return;


            //Students students = new Students();
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
