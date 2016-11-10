using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HWs_Generator
{
    public class JAVA0 : HW0
    {
        public JAVA0()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\" + this.GetType().Name;

            pattern_dir = Students_All_Hws_dirs + @"\Patterns_docs";
            pattern_file_copy = @"JAVA0_pattern_Copy.docx";
            pattern_file_orig = @"JAVA0_pattern_Orig.docx";
            exampleRectangleSize = new Size(300, 300);
        }

        public override void createRandomInputFile(int id, string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                //sw.WriteLine();
            }
        }

        public override bool BuildProject(string path, out string resulting_file_path)
        {
            return Compiler.BuildJavaZippedProject(path, out resulting_file_path);
        }

        public override void Create_HW(Object[] args, bool real_input)
        {
            int id = (int)args[0];
            Student stud = Students.students_dic[id];

            Console.WriteLine("Hello World");
            Console.WriteLine("{0}", stud.email);

            System.Threading.Thread.Sleep(1000);

            if (real_input == false)
            {
                Process lol = Process.GetCurrentProcess();
                IntPtr ptr = lol.MainWindowHandle;
                Rect ConsoleRect = new Rect();
                GetWindowRect(ptr, ref ConsoleRect);

                // Set the bitmap object to the size of the screen
                Size screenCaptureSize = new Size(300, 300);
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
                Console.Clear();
                Create_DocFile(args);
            }

        }

        public override void Create_DocFile(Object[] args)
        {
            int id = (int)args[0];

            String orig_file_path = pattern_dir + "//" + pattern_file_orig;
            //ADDING A NEW DOCUMENT TO THE APPLICATION
            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;

            Microsoft.Office.Interop.Word.Document wordDoc = oWord.Documents.Open(orig_file_path);
            String student_full_name = Students.students_dic[id].first_name + " " + Students.students_dic[id].last_name;
            Worder.Replace_in_doc(wordDoc, "AAAA", student_full_name);
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

        public override Object[] LoadArgs(int id)
        {
            return get_random_args(id);
        }

        public override RunResults Test_HW(Object[] args, String resulting_exe_path)
        {
            RunResults rr = new RunResults();
            FileInfo resultingClassFile = new FileInfo(resulting_exe_path);
            String randomInputFilesFolder = resultingClassFile.DirectoryName + "//GeneratedInput";
            if (!Directory.Exists(randomInputFilesFolder)) Directory.CreateDirectory(randomInputFilesFolder);

            String classNameOnly = resultingClassFile.Name.Replace(resultingClassFile.Extension, "");

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
            p.StartInfo.FileName = @"java";
            p.StartInfo.Arguments = "-cp .. " + classNameOnly;
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
            ProcessStartInfo psi = new ProcessStartInfo("java");
            psi.Arguments = "-cp .. " + classNameOnly; 
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

        public override object[] get_random_args(int id)
        {
            Object[] args = new Object[1];
            args[0] = id;
            return args;
        }



    }
}
