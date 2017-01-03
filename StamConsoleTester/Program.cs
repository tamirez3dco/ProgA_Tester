using HWs_Generator;
using Microsoft.Office.Interop.Word;
using SharpCompress.Common;
using SharpCompress.Reader;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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


        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;


        public static void DoMouseClick(int X, int Y)
        {
            Cursor.Position = new System.Drawing.Point(X, Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, (uint)X, (uint)Y, 0, 0);
        }


        public static void test3()
        {
            Assembly studentApp = Assembly.LoadFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI1_Mine\GUI1_Mine\bin\Debug\GUI1_Mine.exe");
            Type[] appTypes = studentApp.GetTypes();
            //studentApp.get

            Type son_form = null;
            foreach (Type t in appTypes)
            {
                Type parent_form = t.BaseType;
                while (parent_form != null && parent_form != typeof(Object))
                {
                    if (parent_form == typeof(System.Windows.Forms.Form))
                    {
                        son_form = t;
                        break;
                    }
                    parent_form = parent_form.BaseType;
                }
            }


            Type[] constructor_param_types = { typeof(int) };
            ConstructorInfo form_empty_constructor = son_form.GetConstructor(constructor_param_types);

            Object[] constructor_params = { 122 };
            formtoshow = (Form)form_empty_constructor.Invoke(constructor_params);
            
            Button b = (Button)formtoshow.Controls[0];
            ThreadStart ts = new ThreadStart(ShowInThread);
            Thread th = new Thread(ts);
            th.Start();

            Thread.Sleep(2000);
            Debug.WriteLine(formtoshow.BackColor.R);
            Debug.WriteLine(formtoshow.Text);
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(3000);
                DoMouseClick(200, 200);
                formtoshow.Refresh();
                Debug.WriteLine("Form back="+formtoshow.BackColor);
                Debug.WriteLine("Button back="+b.BackColor);
            }
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(3000);
                DoMouseClick(250, 250);
                Debug.WriteLine("Form back=" + formtoshow.BackColor);
                Debug.WriteLine("Button back=" + b.BackColor);
            }


        }

        public static Form formtoshow;
        public static void ShowInThread()
        {
            System.Windows.Forms.Application.Run(formtoshow);
        }

        public static Object[] ObjArrayFromString_GUI3(String s)
        {
            String[] tokeenizer = { "," };
            String[] tokens = s.Split(tokeenizer, StringSplitOptions.RemoveEmptyEntries);
            Object[] res = new Object[tokens.Length];
            res[(int)GUI3_ARGS.ID] = int.Parse(tokens[(int)GUI3_ARGS.ID]);
            String colorStr = tokens[(int)GUI3_ARGS.GATE_DIS_COLOR];
            String[] tokenizer = { "[", " ", "]" };
            String[] miniTokens = colorStr.Split(tokenizer, StringSplitOptions.RemoveEmptyEntries);
            res[(int)GUI3_ARGS.GATE_DIS_COLOR] = Color.FromName(miniTokens[1]);
            res[(int)GUI3_ARGS.GATE_BUTTON_SIDE] = Enum.Parse(typeof(SIDE),tokens[(int)GUI3_ARGS.GATE_BUTTON_SIDE]);
            res[(int)GUI3_ARGS.GATE_PIX_WIDTH] = int.Parse(tokens[(int)GUI3_ARGS.GATE_PIX_WIDTH]);
            res[(int)GUI3_ARGS.MEGA_PATTERN] = int.Parse(tokens[(int)GUI3_ARGS.MEGA_PATTERN]);
            return res;
        }

        static private void doshit(String s)
        {
            s = s + "w";
        }
        static void Main(string[] args)
        {
            /*
                        new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\Programming_A_Semester_A_2017.xlsx");
                        HW4 hww = new HW4();
                        int tid = 301763967;
                        Object[] thw_args = hww.LoadArgs(tid);
                        RunResults rr = hww.test_Hw_by_assembly(thw_args, new FileInfo(
                            @"D:\Tamir\Netanya_Desktop_App\2017\Students_Submissions\GUI2\301763967\23_12_2016_10_00_extracted\HW2\HW2\bin\Debug\HW2.exe"));
                        MessageBox.Show(rr.ToString());
                        MessageBox.Show(rr.ToString());
                        return;
            */
/*
            String tempBin = @"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\results.bin";
            if (File.Exists(tempBin))
            {
                IFormatter formatter = new BinaryFormatter();
                List<GUI3_GateButton_Comparer.GuiResults> ress;
                using (Stream stream = new FileStream(tempBin, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    ress = (List<GUI3_GateButton_Comparer.GuiResults>)formatter.Deserialize(stream);
                }
            }
*/
                new Students(@"D:\Tamir\Netanya_Desktop_App\2017\Shana_B_2017.xlsx");
                        GUI3 hww = new GUI3();
                        int tid = 301763967;
                        //String resulting_exe_path;
                        //Compiler.BuildZippedProject(@"D:\Tamir\Netanya_Desktop_App\2017\Students_Submissions\GUI1\312441710\13_11_2016_14_37.zip", out resulting_exe_path);
                        //Object[] thw_args = hww.get_random_args(tid);
                        Object[] thw_args = hww.LoadArgs(tid);
                        RunResults rr = hww.test_Hw_by_assembly(thw_args, new FileInfo(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\GUI3_Mine.exe"));
                        MessageBox.Show(rr.ToString());
                        MessageBox.Show(rr.ToString());
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
            GUI3 hw = new GUI3();
            Object[] myargs = hw.LoadArgs(tl.id);
            RunResults rr4 = hw.test_Hw_by_assembly(myargs, fin);
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
