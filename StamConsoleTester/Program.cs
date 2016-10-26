using HWs_Generator;
using Microsoft.Office.Interop.Word;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace StamConsoleTester
{
  


    class Program
    {
        static String lastLine;
        static List<RunLine> lines = new List<RunLine>();
        static bool stop = false;
        static Process p;
        static void testSomething()
        {
            String resulting_exe_path = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\66117466\6_10_2016_06_13_extracted\Summer_HW3_066117466\bin\Debug\Summer_HW3_066117466.exe";
            String randomInputFilesFolder = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW3\66117466\6_10_2016_06_13_extracted\Summer_HW3_066117466\bin\Debug\GeneratedInput";
            String randomInputFile = randomInputFilesFolder + "\\" + "test.txt";
            p = new Process();
            p.StartInfo.FileName = resulting_exe_path;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.WorkingDirectory = randomInputFilesFolder;
            p.ErrorDataReceived += P_ErrorDataReceived;
            p.OutputDataReceived += P_OutputDataReceived;
            p.EnableRaisingEvents = true;


            p.Start();
            p.BeginErrorReadLine();
            p.BeginOutputReadLine();


            StreamWriter inputWriter = p.StandardInput;
            String[] inputLines = File.ReadAllLines(randomInputFile);
            int kk = 0;
            while (kk < inputLines.Length && !stop)
            {

                Thread.Sleep(200);
                String line = inputLines[kk];
                lines.Add(new RunLine(StudentsLib.Source.INPUT, line));
                Debug.WriteLine("line=" + line);
                inputWriter.WriteLine(line);
                Thread.Sleep(200);
                kk++;

                //p.CancelOutputRead();
                //p.ca
            }

            if (!p.WaitForExit(10000))
            {
                p.Kill();
            }
            //            string output = p.StandardOutput.ReadToEnd();
            Worder.LinesToTable(lines, randomInputFilesFolder + "//student_run_table.docx");
            String studentOutputFileName = "Student_output.txt";
            String studentOutputFile = randomInputFilesFolder + "//" + studentOutputFileName;
            //File.WriteAllText(studentOutputFile, output);
            return;

        }

        private static void P_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            lines.Add(new RunLine(StudentsLib.Source.OUTPUT, e.Data));
            Debug.WriteLine("Output received {0} : {1} : {2}", lastLine, sender, e.Data);
        }



        static void Main(string[] args)
        {
            GUI1 hww = new GUI1();
            int tid = 029046117;
            Object[] thw_args = hww.get_random_args(tid);
            hww.test_Hw_by_assembly(thw_args, new FileInfo(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI1_Mine\GUI1_Mine\bin\Debug\GUI1_Mine.exe"));
            return;

/*
            Students students = new Students();
            int tid = 317883007;
            HW2 hww = new HW2();
            Object[] myargs = hww.LoadArgs(tid);
            RunResults rr = hww.Test_HW(myargs, @"D:\Tamir\Temp\14_10_2016_21_43_extracted\ConsoleApplication3\ConsoleApplication3\bin\Debug\ConsoleApplication3.exe");
            return;
*/
            //testSomething();
            //return;

            StudentsLib.Student tl = new StudentsLib.Student();
            tl.first_name = "תמיר";
            tl.last_name = "לוי";
            tl.id = 029046117;
            tl.email = "tamirlevi123@gmail.com";
            StudentsLib.Students.students_dic = new Dictionary<int, StudentsLib.Student>();
            StudentsLib.Students.students_dic[tl.id] = tl;
            //StudentsLib.Students
            FileInfo fin = new FileInfo(@"D:\Tamir\Netanya_ProgrammingA\2017\TempSolutions\HW4_Mine\HW4_Mine\bin\Debug\HW4_Mine.exe");
            HW4 hw4 = new HW4();
            Object[] myargs = hw4.LoadArgs(tl.id);
            hw4.test_Hw_by_assembly(myargs, fin);
            //Object[] argsTest = hw4.get_random_args(tl.id);
            //hw3.Create_DocFile_By_Creators(argsTest, new List<Creators>());


            //            hw3.test_Hw(argsTest, fin);
        }

        private static void P_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            lines.Add(new RunLine(StudentsLib.Source.ERROR, e.Data));
            Debug.WriteLine("Error received {0} : {1} : {2}",lastLine, sender, e.Data);
            stop = true;
            p.CancelOutputRead();
        }
    }
}
