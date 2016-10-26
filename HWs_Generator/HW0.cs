using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Diagnostics;
using StudentsLib;
using System.IO;
using System.Threading;
using DiffPlex.DiffBuilder;
using DiffPlex;
using DiffPlex.DiffBuilder.Model;

namespace HWs_Generator
{
    // TODO : Fix output gathering to repreform task with input to get accurate output (without any ununderstandable blank lines at the end)
    // TODO : Fix name of attachments in email to short version...
    public class HW0
    {
        public String Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs";
        public String pattern_dir = @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs";
        public String pattern_file_copy = @"HW0_pattern_Copy.docx";
        public String pattern_file_orig = @"HW0_pattern_Orig.docx";
        public String Students_Hws_dirs;
        public Size exampleRectangleSize = new Size(300, 900);
        public int Num_Of_Test_Tries = 1;

        public HW0()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW0";
        }

        protected List<RunLine> lines = new List<RunLine>();
        protected Process p;
        protected bool stop = false;
        protected void P_ErrorDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (e.Data == null) return;
            lines.Add(new RunLine(StudentsLib.Source.ERROR, e.Data));
            stop = true;
            p.CancelOutputRead();
        }
        protected void P_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            lines.Add(new RunLine(StudentsLib.Source.OUTPUT, e.Data));
        }


        public static Random r = new Random();


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string strClassName, string strWindowName);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        public struct Rect
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        public static String StringfromDoubleArray(double[] arr)
        {
            String temp = String.Empty;
            for (int i = 0; i < arr.Length; i++) temp += arr[i].ToString() + ",";
            return "{" + temp.Substring(0,temp.Length-1)+"}";
        }

        public static String StringfromObjArray(Object[] arr)
        {
            String res = String.Empty;
            for (int i = 0; i < arr.Length; i++) res += arr[i].ToString() + ",";
            return res;
        }

        public static Object[] ObjArrayFromString(String s)
        {
            String[] tokeenizer = { "," };
            String[] tokens = s.Split(tokeenizer, StringSplitOptions.RemoveEmptyEntries);
            Object[] res = new Object[tokens.Length];
            for (int i = 0; i < tokens.Length; i++)
            {
                res[i] = int.Parse(tokens[i]);
            }
            return res;
        }

        public void print_square(int size)
        {
            for (int i = 0; i < size; i++)
            {
                if (i == 0 || i == size - 1)
                {
                    Console.WriteLine(new string('*', size));
                }
                else Console.WriteLine("*" + new string(' ', size - 2) + "*");
            }
        }

        public void print_meshulash(int size)
        {
            for (int i = size-1; i > 0; i--)
            {
                String temp = new String(' ', i) + "*" + new String(' ', size - i - 1);
                String temp_reverse = new String(' ', size -i -1) + "*" + new String(' ', size);
                Console.WriteLine(temp + temp_reverse);
            }
            Console.WriteLine(new String('*', size * 2));
        }

        public Object[] LoadArgs(int id)
        {
            String studentArgsFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_args.txt";
            return ObjArrayFromString(File.ReadAllText(studentArgsFilePath));
        }

        public void SaveArgs(Object[] args)
        {
            int id = (int)(args[0]);
            String studentArgsFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_args.txt";
            if (File.Exists(studentArgsFilePath)) File.Delete(studentArgsFilePath);
            File.WriteAllText(studentArgsFilePath,StringfromObjArray(args));
        }

        public String getRandomString()
        {
            String s = "`1234567890-=qwertyuiop[]asdfghjkl;'\\zxcvbnm,./~!@#$%^&*()_+QWERTYUIOP{}ASDFGHJKL:\"ZXCVBNM<>?             ";
            int stringLength = r.Next(1, 20);
            String res = String.Empty;
            for (int i = 0; i < s.Length; i++) res += s[r.Next(0, s.Length)];
            return res;
        }

        public virtual void createRandomInputFile(int id, String filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                sw.WriteLine(getRandomString());
                sw.WriteLine(getRandomString());
                sw.WriteLine(String.Empty);
                sw.WriteLine(String.Empty);
                sw.WriteLine(String.Empty);
            }
        }


        public void print_meuyan(int size)
        {
            for (int i = size - 1; i >= 0; i--)
            {
                String temp = new String(' ', i) + "*" + new String(' ', size - i - 1);
                String temp_reverse = new String(' ', size - i - 1) + "*" + new String(' ', i);
                Console.WriteLine(temp + temp_reverse);
            }
            for (int i = 0; i < size ; i++)
            {
                String temp = new String(' ', i) + "*" + new String(' ', size - i - 1);
                String temp_reverse = new String(' ', size - i - 1) + "*" + new String(' ', i);
                Console.WriteLine(temp + temp_reverse);
            }
        }

        public static String get_input(bool realinput, int num)
        {
            if (!realinput)
            {
                String res = "kelet" + num;
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(res);
                Console.ForegroundColor = ConsoleColor.White;
                return res;
            }
            else return Console.ReadLine();
        }

        public static String get_input_string(bool realinput, String non_real_input)
        {
            if (!realinput)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(non_real_input);
                Console.ForegroundColor = ConsoleColor.White;
                return non_real_input;
            }
            else return Console.ReadLine();
        }

        public virtual Object[] get_random_args(int id)
        {
            Object[] args = new Object[5];
            args[0] = id;
            args[1] = r.Next(0, 3);
            args[2] = r.Next(4, 8);
            args[3] = r.Next(3, 6);
            args[4] = r.Next(2, 5);
            return args;
        }


        public virtual RunResults Test_Text(Object[] args, String expectedText, String studentText, String prefix)
        {
            RunResults rr = new RunResults();

            SideBySideDiffBuilder diffBuilder = new SideBySideDiffBuilder(new Differ());
            var model = diffBuilder.BuildDiffModel(expectedText ?? string.Empty, studentText ?? string.Empty);
            List<String> comparisonErrors = new List<string>();
            for (int i = 0; i < model.NewText.Lines.Count; i++)
            {
                DiffPiece dp = model.NewText.Lines[i];
                switch (dp.Type)
                {
                    case ChangeType.Unchanged:
                        rr.changes_counter[(int)TextDiffs.No_Diff]++;
                        continue;
                    case ChangeType.Modified:
                        // check if minor diff

                        bool minorDiff = (dp.Text.ToLower().Trim() == model.OldText.Lines[i].Text.ToLower().Trim());
                        if (minorDiff)
                        {
                            rr.error_lines.Add(String.Format("{0} - Minor Diff at line # {1}. Minus 1 pts.", prefix, (int)dp.Position));
                            rr.grade -= 1;
                            rr.changes_counter[(int)TextDiffs.Modified_Minor]++;
                        }
                        else
                        {
                            rr.error_lines.Add(String.Format("{0} - Major Diff at line # {1}. Minus 5 pts.", prefix, (int)dp.Position));
                            rr.grade -= 5;
                            rr.changes_counter[(int)TextDiffs.Modified_Major]++;

                        }

                        rr.error_lines.Add(String.Format("{0} -  Correct line is \"{1}\"", prefix, model.OldText.Lines[i].Text));
                        rr.error_lines.Add(String.Format("{0} -      Your Line is \"{1}\"", prefix, dp.Text));
                        break;
                    case ChangeType.Inserted:
                        if (dp.Text == String.Empty)
                        {
                            rr.grade -= 2;
                            rr.error_lines.Add(String.Format("{0} - Extra empty line at line # {1}. Minus 2 pts.", prefix, (int)dp.Position));
                            rr.changes_counter[(int)TextDiffs.Extra_Empty]++;
                        }
                        else if (dp.Text.Trim() == String.Empty)
                        {
                            rr.grade -= 4;
                            rr.error_lines.Add(String.Format("{0} - Extra line of blanks at line # {1}. Minus 4 pts.", prefix, (int)dp.Position));
                            rr.changes_counter[(int)TextDiffs.Extra_blanks]++;
                        }
                        else
                        {
                            rr.grade -= 7;
                            rr.error_lines.Add(String.Format("{0} - Extra line at line # {1}. Minus 7 pts.", prefix, (int)dp.Position));
                            rr.error_lines.Add(String.Format("{0} - Your Line is \"{1}\"", prefix, dp.Text));
                            rr.changes_counter[(int)TextDiffs.Extra_line]++;
                        }
                        break;
                    case ChangeType.Deleted:
                    case ChangeType.Imaginary:
                        rr.grade -= 5;
                        rr.error_lines.Add(String.Format("{0} - Missing line at line # {1}. Minus 5 pts.", prefix, i + 1));
                        rr.error_lines.Add(String.Format("{0} - expected Line is \"{1}\"", prefix, model.OldText.Lines[i].Text));
                        rr.changes_counter[(int)TextDiffs.Missing]++;
                        break;
                }
            }
            return rr;
        }

        public virtual RunResults Test_HW(Object[] args, String resulting_exe_path)
        {
            RunResults rr = new RunResults();
            String randomInputFilesFolder = new FileInfo(resulting_exe_path).DirectoryName + "//GeneratedInput";
            if (!Directory.Exists(randomInputFilesFolder)) Directory.CreateDirectory(randomInputFilesFolder);


            HW0 hw = (HW0)this;
            Student stud = Students.students_dic[(int)args[0]];

            // create random input file
            String randomFileName = "test.txt";
            String studentOutputFileName = "student_output.txt";
            String benchmarkOutputFileName = "benchmark_output.txt";

            String randomInputFile = randomInputFilesFolder + "//" + randomFileName;
            if (File.Exists(randomInputFile))
            {
                File.Delete(randomInputFile);
                System.Threading.Thread.Sleep(500);
            }

            hw.createRandomInputFile(stud.id, randomInputFile);


            // run through student build and send to output
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
                inputWriter.WriteLine(line);
                Thread.Sleep(200);
                kk++;
            }

            if (!p.WaitForExit(15000))
            {
                p.Kill();
                rr.filesToAttach.Add(randomInputFile);
                if (RunLine.GetErrors(lines).Trim() != String.Empty)
                {
                    String wordTableFilePath = Worder.LinesToTable(lines, randomInputFilesFolder);
                    rr.filesToAttach.Add(wordTableFilePath);
                    rr.grade -= 50;
                    rr.error_lines.Add("Running your program did not complete in 10 seconds. Probably some exception was thrown. Minus 50 pts. The input I tried to feed to your program is attached to the email sent to you. The file \"run_table.docx\" presents the input//output//error data of the running of your code.");
                    return rr;
                }
                else // should be when some abnoctious ReadLine() or ReadKey() was added
                {
                    rr.grade -= 5;
                    rr.error_lines.Add("Running your program did not complete in 10 seconds. Probably some unexpected Console.ReadLine() is blocking it from completion. Minus 5 pts. The input I tried to feed to your program is attached to the email sent to you.");
                }
            }

            string output = RunLine.GetOutputs(lines);

            // run again through student build and send to output
            ProcessStartInfo psi = new ProcessStartInfo(resulting_exe_path);
            psi.UseShellExecute = false;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            
            psi.WorkingDirectory = randomInputFilesFolder;
            p = Process.Start(psi);
            inputWriter = p.StandardInput;
            inputLines = File.ReadAllLines(randomInputFile);
            foreach (String line in inputLines) inputWriter.WriteLine(line);
            
            if (!p.WaitForExit(10000))
            {
                p.Kill();
            }
            output = p.StandardOutput.ReadToEnd();
            String studentOutputFile = randomInputFilesFolder + "//" + studentOutputFileName;
            File.WriteAllText(studentOutputFile, output);

            // run through official HW to get output
            TextReader oldInput = Console.In;
            TextWriter oldOutput = Console.Out;
            String BenchmarkOutputFile = randomInputFilesFolder + "//" + benchmarkOutputFileName;
            using (StreamWriter sw = new StreamWriter(BenchmarkOutputFile, false))
            {
                Console.SetIn(new StreamReader(randomInputFile));
                Console.SetOut(sw);
                hw.Create_HW(args, true);
            }
            Console.SetIn(oldInput);
            Console.SetOut(oldOutput);
            // compare and give feedback

            String studentText = File.ReadAllText(studentOutputFile);
            String benchmarkText = File.ReadAllText(BenchmarkOutputFile);

            RunResults rr_output = new RunResults();

            SideBySideDiffBuilder diffBuilder = new SideBySideDiffBuilder(new Differ());
            var model = diffBuilder.BuildDiffModel(benchmarkText ?? string.Empty, studentText ?? string.Empty);
            for (int i = 0; i < model.NewText.Lines.Count; i++)
            {
                DiffPiece dp = model.NewText.Lines[i];
                switch (dp.Type)
                {
                    case ChangeType.Unchanged:
                        continue;
                    case ChangeType.Modified:
                        rr_output.grade -= 5;
                        rr_output.error_lines.Add(String.Format("Diff at line # {0}. Minus 5 pts.", (int)dp.Position));
                        rr_output.error_lines.Add(String.Format("  Correct line is \"{0}\"", model.OldText.Lines[i].Text));
                        rr_output.error_lines.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                        break;
                    case ChangeType.Inserted:
                        if (dp.Text == String.Empty)
                        {
                            rr_output.grade -= 5;
                            rr_output.error_lines.Add(String.Format("Extra empty line at line # {0}. Minus 5 pts.", (int)dp.Position));
                        }
                        else if (dp.Text.Trim() == String.Empty)
                        {
                            rr_output.grade -= 7;
                            rr_output.error_lines.Add(String.Format("Extra line of blanks at line # {0}. Minus 7 pts.", (int)dp.Position));
                        }
                        else
                        {
                            rr_output.grade -= 10;
                            rr_output.error_lines.Add(String.Format("Extra line at line # {0}. Minus 10 pts.", (int)dp.Position));
                            rr_output.error_lines.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                        }
                        break;
                    case ChangeType.Deleted:
                    case ChangeType.Imaginary:
                        rr_output.grade -= 10;
                        rr_output.error_lines.Add(String.Format("Missing line at line # {0}. Minus 10 pts.", i + 1));
                        rr_output.error_lines.Add(String.Format("     expected Line is \"{0}\"", model.OldText.Lines[i].Text));
                        break;
                }
            }

            if (rr_output.grade < 100)
            {
                rr_output.filesToAttach.Add(randomInputFile);
                rr_output.filesToAttach.Add(studentOutputFile);
                rr_output.filesToAttach.Add(BenchmarkOutputFile);
                rr_output.error_lines.Insert(0, String.Format("Your last submission was not correct.It run but did not give exactly the desired output. Follwoing are the differneces to expected output. The input used to test is attached to this email at file \"{0}\". Your output is attached at file \"{1}\". Expected output is attached at file \"{2}\". Please fix program and upload project again to Moodle. Detailed differences between your output and the expected one are:\n", randomFileName, studentOutputFileName, benchmarkOutputFileName));
            }

            rr = rr + rr_output;
            return rr;

        }

        public virtual void Create_DocFile(Object[] args)
        {
            int id = (int)args[0];
            int shape = (int)args[1];
            int shape_size = (int)args[2];
            int kelet_repetitions = (int)args[3];
            int shave_reps = (int)args[4];

            String orig_file_path = pattern_dir + "//" + pattern_file_orig;
            //ADDING A NEW DOCUMENT TO THE APPLICATION
            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;

            Microsoft.Office.Interop.Word.Document wordDoc = oWord.Documents.Open(orig_file_path);
            String student_full_name = Students.students_dic[id].first_name +" " + Students.students_dic[id].last_name;
            Worder.Replace_in_doc(wordDoc, "AAAA", student_full_name);
            String[] shapes = {"ריבוע","משולש","מעוין"};
            Worder.Replace_in_doc(wordDoc, "BBBB", shapes[shape]);
            Worder.Replace_in_doc(wordDoc, "CCCC", shape_size.ToString());
            Worder.Replace_in_doc(wordDoc, "DDDD", shave_reps.ToString());
            Worder.Replace_in_doc(wordDoc, "EEEE", kelet_repetitions.ToString());
            Worder.Replace_to_picture(wordDoc, "FFFF", Students_Hws_dirs + "\\" + id.ToString() + ".png");
            wordDoc.Save();
            wordDoc.Close();
            oWord.Quit();
            String studentDocFilePath = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
            System.Threading.Thread.Sleep(500);
            if (File.Exists(studentDocFilePath)) File.Delete(studentDocFilePath);
            File.Move(orig_file_path, studentDocFilePath);
            System.Threading.Thread.Sleep(300);
            String copy_file_path = pattern_dir + "//" + pattern_file_copy;
            File.Copy(copy_file_path, orig_file_path);

            return;
        }

        public void GetConsoleRectImage(Object[] args)
        {
            int id = (int)args[0];
            Process lol = Process.GetCurrentProcess();
            IntPtr ptr = lol.MainWindowHandle;
            Rect ConsoleRect = new Rect();
            GetWindowRect(ptr, ref ConsoleRect);

            // Set the bitmap object to the size of the screen
            Bitmap bmpScreenshot = new Bitmap(exampleRectangleSize.Width, exampleRectangleSize.Height, PixelFormat.Format32bppArgb);
            // Create a graphics object from the bitmap
            Graphics gfxScreenshot = Graphics.FromImage(bmpScreenshot);
            // Take the screenshot from the upper left corner to the right bottom corner
            Thread.Sleep(1000);
            gfxScreenshot.CopyFromScreen(ConsoleRect.Left + 8, ConsoleRect.Top, 0, 0, exampleRectangleSize, CopyPixelOperation.SourceCopy);
            Thread.Sleep(1000);

            float ratio = 0.8f;
            Size resized_size = new Size((int)(exampleRectangleSize.Width * ratio), (int)(exampleRectangleSize.Height * ratio));
            Bitmap output = new Bitmap(bmpScreenshot, resized_size);
            if (!Directory.Exists(Students_Hws_dirs)) Directory.CreateDirectory(Students_Hws_dirs);
            output.Save(Students_Hws_dirs + "\\" + id.ToString() + ".png", ImageFormat.Png);

        }


        public virtual void Create_HW(Object[] args,bool real_input)
        {
            
            int id = (int)args[0];
            int shape = (int)args[1];
            int shape_size = (int)args[2];
            int kelet_repetitions = (int)args[3];
            int shave_reps = (int)args[4];
            Console.WriteLine("{0}",id.ToString("D9"));
            Console.WriteLine("** {0} **", id.ToString("D9"));
            Console.WriteLine();

            switch (shape)
            {
                case 0:
                    print_square(shape_size);
                    break;
                case 1:
                    print_meshulash(shape_size);
                    break;
                case 2:
                    print_meuyan(shape_size);
                    break;
            }

            String kelet1 = get_input(real_input, 1);
            Console.WriteLine(new String('=', shave_reps));

            for (int i = 0; i <= kelet_repetitions; i++)
            {
                Console.WriteLine(new String(' ', i) + kelet1);
            }
            for (int i = kelet_repetitions - 1; i >= 0; i--)
            {
                Console.WriteLine(new String(' ', i) + kelet1);
            }
            String kelet2 = get_input(real_input, 2);
            Console.WriteLine("{0} {1}", kelet1, kelet2);
            Console.WriteLine("{0} {1}", kelet2, kelet1);

            System.Threading.Thread.Sleep(1000);

            if (real_input == false)
            {
                Process lol = Process.GetCurrentProcess();
                IntPtr ptr = lol.MainWindowHandle;
                Rect ConsoleRect = new Rect();
                GetWindowRect(ptr, ref ConsoleRect);

                // Set the bitmap object to the size of the screen
                Size screenCaptureSize = new Size(300, 900);
                Bitmap bmpScreenshot = new Bitmap(screenCaptureSize.Width, screenCaptureSize.Height, PixelFormat.Format32bppArgb);
                // Create a graphics object from the bitmap
                Graphics gfxScreenshot = Graphics.FromImage(bmpScreenshot);
                // Take the screenshot from the upper left corner to the right bottom corner
                gfxScreenshot.CopyFromScreen(ConsoleRect.Left + 8, ConsoleRect.Top, 0, 0, screenCaptureSize, CopyPixelOperation.SourceCopy);

                float ratio = 0.8f;
                Size resized_size = new Size((int)(screenCaptureSize.Width * ratio), (int)(screenCaptureSize.Height * ratio));
                Bitmap output = new Bitmap(bmpScreenshot, resized_size);
                output.Save(Students_Hws_dirs + "\\" + id.ToString() + ".png", ImageFormat.Png);

                System.Threading.Thread.Sleep(1000);

                Create_DocFile(args);

            }

        }
    }
}
