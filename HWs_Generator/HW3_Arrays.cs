using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using Microsoft.Office.Interop.Word;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HWs_Generator
{
    public class HW3 : HW2
    {
        public HW3()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW3";
            exampleRectangleSize = new Size(500, 900);
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[4];
            args[0] = id;
            args[1] = r.Next(5, 10); // Q1 - number of students
            args[2] = r.Next(10, 20); // Q2 -- topNum
            args[3] = r.Next(0, 3); // Q3 - type
            return args;

        }

        public override void Create_HW(Object[] args, bool real_input)
        {
            int id = (int)args[0];
            String[] funcsToExecute;
            List<Creators> afterRandom = new List<Creators>();
            if (real_input == false)
            {
                //CreateDocFunc
                afterRandom = new List<Creators>();
                List<Creators> creators = new List<Creators>();
                creators.Add(new Creators(Create_Q1, Create_Q1_doc, Create_Q1_Input));
                creators.Add(new Creators(Create_Q2, Create_Q2_doc, Create_Q2_Input));
                creators.Add(new Creators(Create_Q3, Create_Q3_doc, Create_Q3_Input));
                //creators.Add(new Creators(Create_Q4, Create_Q4_doc, Create_Q4_Input));

                while (creators.Count > 0)
                {
                    int rndIdx = 0;
                    Debug.WriteLine("rndIdx=" + rndIdx);
                    afterRandom.Add(creators[rndIdx]);
                    creators.RemoveAt(rndIdx);
                }

                funcsToExecute = saveOrderFunctions(id, afterRandom);

            }
            else
            {
                funcsToExecute = loadOrderFunctions(id);
            }


            //Debug.WriteLine(afterRandom[0].questFunc.Method.Name)
            for (int i = 0; i < funcsToExecute.Length; i++)
            {
                Console.WriteLine("**********{0}", i + 1);
                // Get the ItsMagic method and invoke with a parameter value of 100
                MethodInfo magicMethod = this.GetType().GetMethod(funcsToExecute[i]);
                magicMethod.Invoke(this, new object[] { args, real_input });
            }

            if (real_input == false)
            {
                GetConsoleRectImage(args);

                Create_DocFile_By_Creators(args, afterRandom);

            }
        }

        public override RunResults Test_HW(Object[] args, String resulting_exe_path)
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

            if (!p.WaitForExit(10000))
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


            String[] tokenizer = new String[3];
            for (int i = 1; i <= 3; i++)
            {
                tokenizer[i - 1] = new string('*', 10) + i.ToString();
            }
            String[] student_tokens = studentText.Split(tokenizer, StringSplitOptions.RemoveEmptyEntries);
            String[] benchmark_tokens = benchmarkText.Split(tokenizer, StringSplitOptions.RemoveEmptyEntries);
            int studCounter = 0;
            RunResults test1_rr = new RunResults();
            if (studentText.Contains(tokenizer[0])){
                test1_rr = Test_Q1(args, student_tokens[studCounter], benchmark_tokens[0]);
                if (test1_rr.grade < 70)
                {
                    int grades_to_add = 70 - test1_rr.grade;
                    test1_rr.grade += grades_to_add;
                    test1_rr.error_lines.Add(String.Format("Q1 - Adding {0} points to ensure at most 30 pts are taken off Q1", grades_to_add));
                }
                studCounter++;
            }
            else
            {
                test1_rr.grade -= 30;
                test1_rr.error_lines.Add(String.Format("Could not locate tokenizer \"{0}\"", tokenizer[0]));
            }

            RunResults test2_rr = new RunResults();
            if (studentText.Contains(tokenizer[1]))
            {
                test2_rr = Test_Q2(args, student_tokens[studCounter], benchmark_tokens[1]);
                if (test2_rr.grade < 70)
                {
                    int grades_to_add = 70 - test2_rr.grade;
                    test2_rr.grade += grades_to_add;
                    test2_rr.error_lines.Add(String.Format("Q2 - Adding {0} points to ensure at most 30 pts are taken off Q2", grades_to_add));
                }
                studCounter++;

            }
            else
            {
                test2_rr.grade -= 30;
                test2_rr.error_lines.Add(String.Format("Could not locate tokenizer \"{0}\"", tokenizer[1]));
            }

            RunResults test3_rr = new RunResults();
            if (studentText.Contains(tokenizer[2]))
            {
                test3_rr = Test_Q3(args, student_tokens[studCounter], benchmark_tokens[2]);
                if (test3_rr.grade < 70)
                {
                    int grades_to_add = 70 - test3_rr.grade;
                    test3_rr.grade += grades_to_add;
                    test3_rr.error_lines.Add(String.Format("Q3 - Adding {0} points to ensure at most 30 pts are taken off Q3", grades_to_add));
                }
                studCounter++;
            }
            else
            {
                test3_rr.grade -= 30;
                test3_rr.error_lines.Add(String.Format("Could not locate tokenizer \"{0}\"", tokenizer[2]));
            }



            rr = rr + test1_rr + test2_rr + test3_rr;
            if (rr.grade < 100)
            {
                rr.filesToAttach.Add(studentOutputFile);
                rr.filesToAttach.Add(BenchmarkOutputFile);
                rr.filesToAttach.Add(randomInputFile);
            }
            return rr;
        }

        public override void Create_DocFile_By_Creators(Object[] args, List<Creators> creators)
        {
            int id = (int)(args[0]);

            String student_full_name = Students.students_dic[id].first_name + " " + Students.students_dic[id].last_name;


            //            String orig_file_path = pattern_dir + "//" + pattern_file_orig;
            //ADDING A NEW DOCUMENT TO THE APPLICATION
            Application oWord = new Application();
            oWord.Visible = false;
            Document wordDoc = oWord.Documents.Add();

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "שלום " + student_full_name;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "ש\"ב 3 נועדו לתרגל אתכם על מערכים כפי שנלמדו בהרצאה ובתרגול. כרגיל, עליכם לייצר בדיוק את הפלט המצופה כדי שהבודק האוטומטי לא יכשיל אתכם. ושוב כרגיל, דוגמא לפלט המצופה מופיעה בסוף המסמך הזה.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "לפני כל סעיף אבקש להדפיס שורה של 10 כוכביות ומספר הסעיף. לדוגמא, לפני הביצוע של סעיף 3 יש להדפיס את השורה:";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = "3**********";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "אם לא ברור, ניתן להסתכל בדוגמת הפלט הנדרש בסוף המסמך.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Underline = WdUnderline.wdUnderlineSingle;

//            par1.Range.Text = String.Format("תאריך הגשה אחרון - 11/12/2016 בשעה 23:55");
//            par1.Range.InsertParagraphAfter();
            par1.Range.Underline = WdUnderline.wdUnderlineNone;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עליך ליכתוב תוכנית אשר:";
            par1.Range.InsertParagraphAfter();

            for (int i = 0; i < 2; i++)
            {
                creators[i].docFunc(args, wordDoc, (i + 1));
            }

            par1.Range.InsertBreak(WdBreakType.wdPageBreak);

            for (int i = 2; i < creators.Count; i++)
            {
                creators[i].docFunc(args, wordDoc, (i + 1));
            }

            par1.Range.Text = "בעמוד הבא מופיעה דוגמא לפלט המצופה מהתוכנית שלך. זיכרו כי בדוגמא זו, השורות הכחולות מציינות שורות קלט מה-Console שהוכנסו על ידי מריץ התוכנית. השורות הלבנות מסמנות שורות פלט שנכתבו על ידי התוכנית אל ה-Console.";
            par1.Range.InsertParagraphAfter();

            par1.Range.InsertBreak(WdBreakType.wdPageBreak);
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            wordDoc.Application.Selection.Collapse();
            InlineShape bigExamplePicture = Worder.Replace_to_picture(wordDoc, "XXXX", Students_Hws_dirs + "\\" + id.ToString() + ".png");

            object fileName = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
            object missing = Type.Missing;
            wordDoc.SaveAs(ref fileName,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Close();
            oWord.Quit();
            return;
        }



        public virtual RunResults Test_Q1(Object[] args, String studentText, String benchmarkText)
        {
            RunResults rr = new RunResults();
            String[] tokenizer = { "\n" };
            String[] lines = studentText.Split(tokenizer, StringSplitOptions.RemoveEmptyEntries);
            int firstLineIdx = LevenshteinDistance.GetIndexOfClosest(lines, "Students grades are:");
            int secondLineIdx = LevenshteinDistance.GetIndexOfClosest(lines, "Students grades (reverse) are:");

            if (!(secondLineIdx > firstLineIdx) || (lines.Length - secondLineIdx <= 1))
            {
                rr.grade -= 20;
                rr.error_lines.Add("Could not find proper marker lines \"Students grades are:\" and \"Students grades (reverse) are:\"");
                return rr;
            }
            List<int> allInts = new List<int>();
            for (int i = secondLineIdx + 1; i < lines.Length; i++)
            {
                int line1Idx = secondLineIdx - (i - secondLineIdx);
                if (lines[i] != lines[line1Idx])
                {
                    rr.grade -= 15;
                    rr.error_lines.Add(String.Format("Could not match line #{0}=\"{1}\" to line #{2}=\"{3}\"",i,lines[i],line1Idx,lines[line1Idx]));
                    return rr;
                }
                int x = int.Parse(lines[i]);
                if (!allInts.Contains(x)) allInts.Add(x);
            }
            if (allInts.Count < 3)
            {
                rr.grade -= 15;
                rr.error_lines.Add(String.Format("Not enough random grades. Less then 3 grades observed. Are you sure you randomized the grades?"));
                return rr;
            }
            return rr;
        }

        public virtual RunResults Test_Q2(Object[] args, String studentText, String benchmarkText)
        {
            RunResults rr = new RunResults();

            SideBySideDiffBuilder diffBuilder = new SideBySideDiffBuilder(new Differ());
            var model = diffBuilder.BuildDiffModel(benchmarkText ?? string.Empty, studentText ?? string.Empty);
            List<String> comparisonErrors = new List<string>();
            for (int i = 0; i < model.NewText.Lines.Count; i++)
            {
                DiffPiece dp = model.NewText.Lines[i];
                switch (dp.Type)
                {
                    case ChangeType.Unchanged:
                        continue;
                    case ChangeType.Modified:
                        // check if minor diff

                        bool minorDiff = (dp.Text.ToLower().Trim() == model.OldText.Lines[i].Text.ToLower().Trim());
                        if (minorDiff)
                        {
                            rr.error_lines.Add(String.Format("Q2 - Minor Diff at line # {0}. Minus 1 pts.", (int)dp.Position));
                            rr.grade -= 1;
                        }
                        else
                        {
                            rr.error_lines.Add(String.Format("Q2 - Major Diff at line # {0}. Minus 5 pts.", (int)dp.Position));
                            rr.grade -= 5;
                        }
                        
                        rr.error_lines.Add(String.Format("Q2 -  Correct line is \"{0}\"", model.OldText.Lines[i].Text));
                        rr.error_lines.Add(String.Format("Q2 -      Your Line is \"{0}\"", dp.Text));
                        break;
                    case ChangeType.Inserted:
                        if (dp.Text == String.Empty)
                        {
                            rr.grade -= 2;
                            rr.error_lines.Add(String.Format("Q2 - Extra empty line at line # {0}. Minus 2 pts.", (int)dp.Position));
                        }
                        else if (dp.Text.Trim() == String.Empty)
                        {
                            rr.grade -= 4;
                            rr.error_lines.Add(String.Format("Q2 - Extra line of blanks at line # {0}. Minus 4 pts.", (int)dp.Position));
                        }
                        else
                        {
                            rr.grade -= 7;
                            rr.error_lines.Add(String.Format("Q2 - Extra line at line # {0}. Minus 7 pts.", (int)dp.Position));
                            rr.error_lines.Add(String.Format("       Your Line is \"{0}\"", dp.Text));
                        }
                        break;
                    case ChangeType.Deleted:
                    case ChangeType.Imaginary:
                        rr.grade -= 5;
                        rr.error_lines.Add(String.Format("Q2 - Missing line at line # {0}. Minus 5 pts.", i + 1));
                        rr.error_lines.Add(String.Format("      expected Line is \"{0}\"", model.OldText.Lines[i].Text));
                        break;
                }
            }
            return rr;
        }

        public virtual RunResults Test_Q3(Object[] args, String studentText, String benchmarkText)
        {
            RunResults rr = new RunResults();

            SideBySideDiffBuilder diffBuilder = new SideBySideDiffBuilder(new Differ());
            var model = diffBuilder.BuildDiffModel(benchmarkText ?? string.Empty, studentText ?? string.Empty);
            List<String> comparisonErrors = new List<string>();
            for (int i = 0; i < model.NewText.Lines.Count; i++)
            {
                DiffPiece dp = model.NewText.Lines[i];
                switch (dp.Type)
                {
                    case ChangeType.Unchanged:
                        continue;
                    case ChangeType.Modified:
                        // check if minor diff

                        bool minorDiff = (dp.Text.ToLower().Trim() == model.OldText.Lines[i].Text.ToLower().Trim());
                        if (minorDiff)
                        {
                            rr.error_lines.Add(String.Format("Q3 - Minor Diff at line # {0}. Minus 1 pts.", (int)dp.Position));
                            rr.grade -= 1;
                        }
                        else
                        {
                            rr.error_lines.Add(String.Format("Q3 - Major Diff at line # {0}. Minus 5 pts.", (int)dp.Position));
                            rr.grade -= 5;
                        }

                        rr.error_lines.Add(String.Format("Q3 -  Correct line is \"{0}\"", model.OldText.Lines[i].Text));
                        rr.error_lines.Add(String.Format("Q3 -      Your Line is \"{0}\"", dp.Text));
                        break;
                    case ChangeType.Inserted:
                        if (dp.Text == String.Empty)
                        {
                            rr.grade -= 2;
                            rr.error_lines.Add(String.Format("Q3 - Extra empty line at line # {0}. Minus 2 pts.", (int)dp.Position));
                        }
                        else if (dp.Text.Trim() == String.Empty)
                        {
                            rr.grade -= 4;
                            rr.error_lines.Add(String.Format("Q3 - Extra line of blanks at line # {0}. Minus 4 pts.", (int)dp.Position));
                        }
                        else
                        {
                            rr.grade -= 7;
                            rr.error_lines.Add(String.Format("Q3 - Extra line at line # {0}. Minus 7 pts.", (int)dp.Position));
                            rr.error_lines.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                        }
                        break;
                    case ChangeType.Deleted:
                    case ChangeType.Imaginary:
                        rr.grade -= 5;
                        rr.error_lines.Add(String.Format("Q3 - Missing line at line # {0}. Minus 5 pts.", i + 1));
                        rr.error_lines.Add(String.Format("     expected Line is \"{0}\"", model.OldText.Lines[i].Text));
                        break;
                }
            }
            return rr;
        }

        public new void Create_Q1(Object[] args, bool real_input)
        {
            int students_num = (int)args[1];
            int[] grades_array = new int[students_num];
            for (int i = 0; i < students_num; i++)
            {
                grades_array[i] = r.Next(0, 101);
            }
            Console.WriteLine("Students grades are:");
            for (int i = 0; i < students_num; i++)
            {
                Console.WriteLine(grades_array[i]);
            }
            Console.WriteLine("Students grades (reverse) are:");
            for (int i = students_num - 1; i >= 0; i--)
            {
                Console.WriteLine(grades_array[i]);
            }

        }

        public new String[] Create_Q1_Input(Object[] args)
        {
            String[] res = new String[0];
            return res;
        }

        public new void Create_Q1_doc(Object[] args, Document wordDoc, int seif)
        {
            int students_num = (int)args[1];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) תגריל ציונים עבור {1} סטודנטים ותדפיס אותם. כלומר, על התוכנית להדפיס מערך של מספרים שלמים שאותם היא תגריל. לאחר מכן התוכנית תדפיס את אותם המספרים שהוגרלו רק בסדר הפוך.:", seif, students_num);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("     באופן מפורט, התוכנית תגריל ציון (מספר בין 0 ל-100) לכל אחד מ-{0} הסטודנטים ותכניס את הציון שהוגרל למערך בגודל מתאים", students_num);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("      לפני הדפסת הציונים בפעם הראשונה התוכנית תדפיס את השורה:");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("Students grades are:");
            par1.Range.InsertParagraphAfter();



            par1.Range.Text = String.Format("     אח\"כ התוכנית תדפיס את המספרים שהוגרלו");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("      לפני הדפסת הציונים בסדר ההפוך התוכנית תדפיס את השורה:");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("Students grades (reverse) are:");
            par1.Range.InsertParagraphAfter();


            par1.Range.Text = String.Format("     לבסוף התוכנית תדפיס את המספרים שהוגרלו בסדר הפוך");
            par1.Range.InsertParagraphAfter();

            StudentsLib.Worder.English_Format_By_Search(wordDoc, "Students grades are:");
            StudentsLib.Worder.English_Format_By_Search(wordDoc, "Students grades (reverse) are:");
        }

        public new void Create_Q2(Object[] args, bool real_input)
        {
            int topNum = (int)args[2];
            if (!real_input) topNum = 4;
            int[] non_real_examples = { 2, 0, 2, 3, 2, -1 };
            bool[] flags = new bool[topNum + 1];
            int num;
            int counter = 0;
            do
            {
                int temp = -1;
                if (!real_input) temp = non_real_examples[counter];
                num = int.Parse(get_input_string(real_input, temp.ToString()));
                if (num < 0) break;
                if (flags[num]) Console.WriteLine("Number already entered",num);
                else Console.WriteLine("New Number");
                flags[num] = true;
                counter++;
            } while (num >= 0);

        }

        public new String[] Create_Q2_Input(Object[] args)
        {
            int topNum = (int)args[2];
            String[] res = new string[31];
            for (int i = 0; i < 30; i++)
            {
                res[i] = r.Next(0, topNum + 1).ToString();
            }
            res[30] = (-1).ToString();
            return res;
        }

        public new void Create_Q2_doc(Object[] args, Document wordDoc, int seif)
        {
            int topNum = (int)args[2];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0})  קולטת מספרים שלמים בתחום 0-{1}, לכל אחד מהמספרים שמוכנסים התוכנית מדפיסה את ההודעה New Number אם המספר שהוכנס עוד לא הוכנס קודם, אחרת (המספר שהוכנס כבר הוכנס מקודם) התוכנית תדפיס את ההודעה Number already entered. הלולאה תפסיק כאשר מריץ התוכנית יכניס מספר שלילי. עבור המספר השלילי התוכנית לא תדפיס שום הודעה ורק תפסיק את הלולאה שקולטת את המספרים. דוגמא להרצת הסעיף מופיעה בסוף מסמך זה.:", seif, topNum);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("רמז - רצוי להשתמש במערך של משתנים מסוג bool כך שהערך בתא ה-i במערך ייצג את העובדה שהמספר i כבר הוקלד בעבר.");
            par1.Range.InsertParagraphAfter();
        }


        public new void Create_Q3(Object[] args, bool real_input)
        {
            int size = 5;
            int type = (int)args[3];

            int[] id_array = new int[size];
            Object[] objects_array = new Object[id_array.Length];
            int[] non_real_ids = { 103, 999, 721, 014, 543 };
            int[] non_real_grades = { 90, 85, 75, 85, 100 };
            String[] printers = { "grade", "height", "name" };

            double[] non_real_hights = { 172.1 , 155.5, 192, 177.77, 123.95 };
            String[] non_real_names = { "David", "Avi", "Shimon", "Jacky", "Yossi" };
            for (int i = 0; i < id_array.Length; i++)
            {
                Console.WriteLine("Type in id");
                id_array[i] = int.Parse(get_input_integer(real_input, non_real_ids[i]));
                Console.WriteLine("Type in {0}",printers[type]);
                switch (type)
                {
                    case 0:
                        objects_array[i] = int.Parse(get_input_integer(real_input, non_real_grades[i]));
                        break;
                    case 1:
                        objects_array[i] = double.Parse(get_input_string(real_input, non_real_hights[i].ToString()));
                        break;
                    case 2:
                        objects_array[i] = get_input_string(real_input, non_real_names[i]);
                        break;
                }

            }
            int[] sorted_ids = new int[id_array.Length];
            id_array.CopyTo(sorted_ids, 0);
            Array.Sort(sorted_ids);
            for (int i = 0; i < id_array.Length; i++)
            {
                int id = sorted_ids[i];
                int orig_index = 0;
                while (id_array[orig_index] != id) orig_index++;
                Console.WriteLine("id:{0}, {2}:{1}",id,objects_array[orig_index].ToString(), printers[type]);
            }
        }

        public new String[] Create_Q3_Input(Object[] args)
                {
                    String[] res = new string[10];
            int type = (int)args[3];
            List<int> ids = new List<int>();
            for (int id = 10; id < 200; id++) ids.Add(id);
            for (int i = 0; i < 5; i++)
            {
                int idx = r.Next(0, ids.Count);
                res[2 * i] = ids[idx].ToString();
                ids.RemoveAt(idx);
                switch (type)
                {
                    case 0:
                        res[2 * i + 1] = r.Next(20, 101).ToString();
                        break;
                    case 1:
                        double height = Math.Round(r.NextDouble() * 220, 2);
                        res[2 * i + 1] = height.ToString();
                        break;
                    case 2:
                        res[2 * i + 1] = getRandomString();
                        break;
                }
            }
            int small_num = r.Next(10234, 998654);
                    res[0] = small_num.ToString();
                    int big_num = small_num + r.Next(30, 50);
                    res[1] = big_num.ToString();
                    return res;
                }

                public new void Create_Q3_doc(Object[] args, Document wordDoc, int seif)
                {
                    int type = (int)args[3];
            String[] typename = { "ציון","גובה", "שם" };
            String[] typenames = { "ציונים", "גבהים", "שמות" };
            String[] typetypes = { "int", "double", "String" };
            Paragraph par1 = wordDoc.Paragraphs.Add();
                    par1.Range.Font.Name = "Ariel";
                    par1.Range.Font.Size = 12;
                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) בשאלה זו נקלוט מהמריץ מספרי ת.ז. של 5 סטודנטים וכן את ה{1} שלהם. לאחר מכן נדפיס את ה-ת.ז. וה{1} ממוינים לפי ת.ז מנמוך לגבוהה.:", seif,typename[type]);
                    par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("כדי לעשות זאת הצהירו על 2 מערכים. הראשון של int בגודל 5 כדי לזכור את הת.ז. ומערך שני של {0} בגודל 5 כדי לזכור את ה{1}..",typetypes[type],typenames[type]);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("בשלב הבא עליכם לקלוט את ה-ת.ז וה{0} מהמשתמש. מומלץ לעשות זאת ע\"י לולאה. נא להסתכל בדוגמת ההרצה שבסוף המסמך כדי להבין את פרטי הקלט-פלט של קליטת הת.ז. וה{0}.",typenames[type]);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("לאחר מכן יש להעתיק את תוכן מערך הת.ז. למערך חדש ולמיין את המערך החדש (()Array.Sort). כעת בעצם ניתן להדפיס את הת.ז. ממוינות מקטן לגדול.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("בזמן הדפסת הת.ז. ממוינות יש להדפיס גם את ה{0} שהוכנס עבור הת.ז. השונות. כדי לעשות זאת כדאי לחפש כל ת.ז. שמדפיסים במערך הת.ז. המקורי וכך בעצם לגלות את האינדקס המקורי של הת.ז.",typename[type]);
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("לאחר שגיליתם את האינדקס המקורי - ניתן בקלות לגשת לאותו אינדקס במערך ה{0} כדי לדעת איזה {1} להדפיס עם הת.ז. הנוכחית.",typenames[type],typename[type]);
            par1.Range.InsertParagraphAfter();
                }
/*
                public new void Create_Q4(Object[] args, bool real_input)
                {
                    double number;
                    String input_if_not_real = null;

                    int i = 0;
                    do
                    {
                        Console.WriteLine("Type non-integer number");
                        if (!real_input)
                        {
                            String[] inputs = { "9", "12.0", "-34", "45.5" };
                            input_if_not_real = inputs[i];
                        }
                        String number_string = get_input_string(real_input, input_if_not_real);
                        number = double.Parse(number_string);
                        i++;
                    } while ((int)number == (int)(Math.Ceiling(number)));
                    Console.WriteLine("Finally entered a non-integer number {0}", number);
                }

                public new String[] Create_Q4_Input(Object[] args)
                {
                    String[] inputs = { "9", "12.0", "-34", "45.00", "101", "-137.0", "-66.0", "32.0", "126.5" };
                    int num_of_lines = r.Next(5, 10);
                    String[] res = new string[num_of_lines];
                    for (int i = 0; i < num_of_lines; i++)
                    {
                        res[i] = inputs[i];
                    }
                    res[num_of_lines - 1] = inputs[inputs.Length - 1];
                    return res;
                }

                public new void Create_Q4_doc(Object[] args, Document wordDoc, int seif)
                {
                    int digiter = (int)args[3];
                    Paragraph par1 = wordDoc.Paragraphs.Add();
                    par1.Range.Font.Name = "Ariel";
                    par1.Range.Font.Size = 12;
                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    par1.Range.Text = String.Format("{0}) מנסה לקלוט מהמשתמש מספר לא שלם:", seif);
                    par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
                    par1.Range.InsertParagraphAfter();
                    par1.Range.Text = String.Format("     היא עושה זאת על ידי (שוב ושוב) הדפסת השורה");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    par1.Range.Text = String.Format("Type non-integer number");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    par1.Range.Text = String.Format("     לאחר מכן קולטת שורה מהמשתמש. אם המשתמש הכניס מספר שלם (חיובי או שלילי) התוכנית ממשיכה לבקש מהמשתמש מספר חדש. רק כאשר המשתמש מכניס באמת מספר לא שלם (עם ערך לא אפס אחרי הנקודה העשרונית) התוכנית מפסיקה את הלולאה ומדפיסה את השורה.");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    par1.Range.Text = String.Format("המספר הלא שלם שהוכנס Finally entered a non-integer number ");
                    par1.Range.InsertParagraphAfter();
                    par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                }
        */

    }
}
