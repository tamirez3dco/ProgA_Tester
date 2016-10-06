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
    public class HW0
    {
        public static String Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs";
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
            ProcessStartInfo psi = new ProcessStartInfo(resulting_exe_path);
            psi.UseShellExecute = false;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;

            psi.WorkingDirectory = randomInputFilesFolder;
            Process p = Process.Start(psi);
            StreamWriter inputWriter = p.StandardInput;
            String[] inputLines = File.ReadAllLines(randomInputFile);
            foreach (String line in inputLines) inputWriter.WriteLine(line);

            String[] filesToAttach = new String[3];
            if (!p.WaitForExit(10000))
            {
                filesToAttach[0] = randomInputFile;
                filesToAttach[1] = filesToAttach[2] = String.Empty;
                rr.grade -= 50;
                rr.error_lines.Add("Running your program did not complete in 10 seconds. Probably some exception was thrown or unexpected Console.ReadLine() is blocking it from completion. Minus 50 pts. The input I tried to feed to your program is attached to the email sent to you.");
                String email_body = String.Format("Hi - " + stud.first_name + "\nSorry but the last project you uploaded to Moodle failed to run (hoever, it did compile succesfully). The input I tried to feed to your program is attached to this email at file \"{0}\".\n\n\n. Please check your code and upload again to Moodle!", randomFileName);
                stud.Send_Gmail("Your last submission failed to run.", email_body, filesToAttach);
                return rr;
            }
            string output = p.StandardOutput.ReadToEnd();
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
                        rr.grade -= 5;
                        rr.error_lines.Add(String.Format("Diff at line # {0}. Minus 5 pts.", (int)dp.Position));
                        rr.error_lines.Add(String.Format("  Correct line is \"{0}\"", model.OldText.Lines[i].Text));
                        rr.error_lines.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                        break;
                    case ChangeType.Inserted:
                        if (dp.Text == String.Empty)
                        {
                            rr.grade -= 5;
                            rr.error_lines.Add(String.Format("Extra empty line at line # {0}. Minus 5 pts.", (int)dp.Position));
                        }
                        else if (dp.Text.Trim() == String.Empty)
                        {
                            rr.grade -= 7;
                            rr.error_lines.Add(String.Format("Extra line of blanks at line # {0}. Minus 7 pts.", (int)dp.Position));
                        }
                        else
                        {
                            rr.grade -= 10;
                            rr.error_lines.Add(String.Format("Extra line at line # {0}. Minus 10 pts.", (int)dp.Position));
                            rr.error_lines.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                        }
                        break;
                    case ChangeType.Deleted:
                    case ChangeType.Imaginary:
                        rr.grade -= 10;
                        rr.error_lines.Add(String.Format("Missing line at line # {0}. Minus 10 pts.", i + 1));
                        rr.error_lines.Add(String.Format("     expected Line is \"{0}\"", model.OldText.Lines[i].Text));
                        break;
                }
            }

            if (rr.grade == 100)
            {
                stud.Send_Gmail("Your last submission was perfect!!", "Good job - " + stud.first_name, filesToAttach);
            }
            else
            {
                filesToAttach[0] = randomInputFile;
                filesToAttach[1] = studentOutputFile;
                filesToAttach[2] = BenchmarkOutputFile;

                String explenationLine = String.Format("Follwoing are the differneces to expected output. The input used to test is attached to this email at file \"{0}\". Your output is attached at file \"{1}\". Expected output is attached at file \"{2}\". Please fix program and upload project again to Moodle. Detailed differences between your output and the expected one are:\n {3}", randomFileName, studentOutputFileName, benchmarkOutputFileName, rr.errorsAsSingleString());
                stud.Send_Gmail("Your last submission was not correct. It run but did not give exactly the desired output", explenationLine, filesToAttach);
            }
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
            for (int i = kelet_repetitions - 2; i >= 0; i--)
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
