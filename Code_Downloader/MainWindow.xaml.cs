using MSHTML;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using StudentsLib;
using System.IO.Compression;
using HWs_Generator;
using System.Reflection;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;
using DiffPlex;

namespace Code_Downloader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        public MainWindow()
        {
            InitializeComponent();
            WebBrowser1.Navigate("http://el1.netanya.ac.il/login/index.php");
            Students students = new Students();
            typeToLinkDict = new Dictionary<Type, string>();
            typeToLinkDict[typeof(HW0)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143049&action=grading";
            typeToLinkDict[typeof(HW1)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143050&action=grading";
            typeToLinkDict[typeof(HW2)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=142830&action=grading";
            typeToLinkDict[typeof(HW3)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143026&action=grading";

        }
        int currectIdx = 0;

        public static Dictionary<Type, String> typeToLinkDict;
        private void WebBrowser1_Navigated(object sender, NavigationEventArgs e)
        {
            Debug.WriteLine("Yofi");
        }

        private void WebBrowser1_LoadCompleted(object sender, NavigationEventArgs e)
        {
            HTMLDocument doc = WebBrowser1.Document as HTMLDocument;
            Debug.WriteLine("Completed:url=" + doc.url);
            if (doc.url.EndsWith(@"login/index.php"))
            {
                doc.getElementById("username").setAttribute("value", "levytami");
                doc.getElementById("password").setAttribute("value", "yardena12");
                doc.getElementById("loginbtn").click();
            }

            if (doc.url.EndsWith(@"/mod/assign/view.php"))
            {
                currectIdx++;
                if (currectIdx >= typeToLinkDict.Count)
                {
                    this.Close();
                    return;
                }
                Type t = typeToLinkDict.Keys.ToArray()[currectIdx];
                WebBrowser1.Navigate(typeToLinkDict[t]);
            }
            if (doc.url.EndsWith(@"http://el1.netanya.ac.il/"))
            {
                //WebBrowser1.Navigate(@"http://el1.netanya.ac.il/mod/assign/view.php?id=142263&action=grading");
                WebBrowser1.Navigate(typeToLinkDict[typeof(HW0)]);               
            }
            if (doc.url.EndsWith(@"action=grading"))
            {
                // get correct HW by url
                Type[] types = typeToLinkDict.Keys.ToArray();
                for (int j = 0; j < types.Length; j++)
                {
                    String hw_name = types[j].Name;
                    String hw_url = typeToLinkDict[types[j]];
                    String last_20_letters = hw_url.Substring(hw_url.Length - 20);
                    if (doc.url.EndsWith(last_20_letters))
                    {
                        Debug.WriteLine("Grading...");
                        HTMLTable allTable = (doc.getElementsByClassName("flexible generaltable generalbox")).item(0);
                        for (int r = 0; r < allTable.rows.length; r++)
                        {
                            HTMLTableRow tableRow = allTable.rows.item(r);
                            if (tableRow.className == null) continue;
                            if (tableRow.className.Contains("emptyrow")) continue;
                            Debug.WriteLine("Row #" + r);

                            HTMLTableCell tc_name = tableRow.cells.item(2); // name cell
                            String name = tc_name.innerText.Trim();
                            Debug.WriteLine("name=" + name);

                            HTMLTableCell tc_email = tableRow.cells.item(3); // email cell
                            String email = tc_email.innerText.Trim();
                            Debug.WriteLine("email=" + email);

                            HTMLTableCell tc_last_update = tableRow.cells.item(7); // last upload time
                            String last_Update_Time = tc_last_update.innerText.Trim();
                            Debug.WriteLine("last update time=" + last_Update_Time);

                            if (last_Update_Time.Length < 10) continue;

                            // check if there is a directry for this hw
                            String hw_path = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\" + hw_name;
                            if (!Directory.Exists(hw_path)) Directory.CreateDirectory(hw_path);

                            // check if there is a directry for this id
                            Student stud = Students.students_dic.Where(z => z.Value.email == email).FirstOrDefault().Value;
                            if (stud == null) continue;

                            String folderPath = hw_path +@"\" + stud.id;
                            if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                            String time_String = last_Update_Time.Replace(" ", "");
                            time_String = time_String.Replace(",", "_");
                            time_String = time_String.Replace("/", "_");
                            time_String = time_String.Replace("\\", "_");
                            time_String = time_String.Replace(":", "_");
                            String filePath = folderPath + "\\" + time_String + ".zip";

                            if (File.Exists(filePath))
                            {
                                Debug.WriteLine("File {0} already exists!!!");
                                continue;
                            }

                            HTMLTableCell tc_file = tableRow.cells.item(8); // last upload time
                            HTMLAnchorElement link_to_file = tc_file.getElementsByTagName("a").item(0);

                            CookieContainer cc = new CookieContainer();
                            String[] cookies = doc.cookie.Split(';');
                            foreach (String cokie in cookies)
                            {
                                String entireCookie = cokie.Trim();
                                string cname = entireCookie.Split('=')[0].Trim();
                                if (cname != "MoodleSession") continue;
                                string cvalue = entireCookie.Substring(cname.Length + 1);
                                string path = "/";
                                string domain = "el1.netanya.ac.il"; //change to your domain name
                                cc.Add(new Cookie(cname.Trim(), cvalue.Trim(), path, domain));
                            }
                            CookieAwareWebClient wbc = new CookieAwareWebClient();
                            wbc.CookieContainer = cc;
                            wbc.DownloadFile(link_to_file.href, filePath);

                            HTMLTableCell tc_grade = tableRow.cells.item(5); // grade table cell
                            HTMLInputTextElement grade_box = tc_grade.getElementsByTagName("input").item(0);
                            HTMLTableCell tc_remarks = tableRow.cells.item(11); // grade cell
                            HTMLInputTextElement remarks_box = tc_remarks.getElementsByTagName("textarea").item(0);


                            String resulting_exe_path;
                            if (!Compiler.BuildZippedProject(filePath, out resulting_exe_path))
                            {
                                grade_box.setAttribute("value", "30");
                                remarks_box.setAttribute("value", Compiler.errorReason);
                                stud.Send_Gmail(String.Format("Your last submission of {0} failed to build!!",hw_name), "Hi - " + stud.first_name + "\nSorry but the last project you uploaded to Moodle failed to build. Compilation error was:\n" + Compiler.errorReason + "\n\n\n. Please check your code and upload again!", new List<String>());
                                continue;
                            }


                            String randomInputFilesFolder = new FileInfo(resulting_exe_path).DirectoryName + "//GeneratedInput";
                            if (!Directory.Exists(randomInputFilesFolder)) Directory.CreateDirectory(randomInputFilesFolder);


                            // Start testing....
                            Type type = types[j];
                            ConstructorInfo ctor = type.GetConstructor(Type.EmptyTypes);
                            HW0 hw = (HW0)ctor.Invoke(new object[] { });

                            // 1) get the specific student values...
                            Object[] args = hw.LoadArgs(stud.id);

                            RunResults rr = hw.Test_HW(args, resulting_exe_path);
                            grade_box.setAttribute("value", rr.grade.ToString());

                            if (rr.grade == 100)
                            {
                                remarks_box.setAttribute("value", "Perfect");
                                stud.Send_Gmail(String.Format("Your last submission of {0} was perfect!!",hw_name), "Good job - " + stud.first_name, rr.filesToAttach);
                            }
                            else
                            {
                                remarks_box.setAttribute("value", rr.errorsAsSingleString());
                                stud.Send_Gmail(String.Format("Your last submission of {0} was not correct. It run but did not give exactly the desired output", hw_name), rr.errorsAsSingleString(), rr.filesToAttach);
                            }

                            /*
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

                                                        string output = p.StandardOutput.ReadToEnd();
                                                        if (!p.WaitForExit(10000))
                                                        {
                                                            filesToAttach[0] = randomInputFile;
                                                            filesToAttach[1] = filesToAttach[2] = String.Empty;
                                                            grade_box.setAttribute("value", "50");
                                                            remarks_box.setAttribute("value", "Running your program did not complete in 10 seconds. Minus 50 pts. The input I tried to feed to your program is attached to the email sent to you.");
                                                            String email_body = String.Format("Hi - " + stud.first_name + "\nSorry but the last project you uploaded to Moodle failed to run (hoever, it did compile succesfully). The input I tried to feed to your program is attached to this email at file \"{0}\".\n\n\n. Please check your code and upload again to Moodle!", randomFileName);
                                                            stud.Send_Gmail("Your last submission failed to run.", email_body, filesToAttach);
                                                            continue;
                                                        }
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
                                                        int errorGrades = 0;
                                                        List<String> comparisonErrors = new List<string>();
                                                        for (int i = 0; i < model.NewText.Lines.Count; i++)
                                                        {
                                                            DiffPiece dp = model.NewText.Lines[i];
                                                            switch (dp.Type)
                                                            {
                                                                case ChangeType.Unchanged:
                                                                    continue;
                                                                case ChangeType.Modified:
                                                                    errorGrades += 5;
                                                                    comparisonErrors.Add(String.Format("Diff at line # {0}. Minus 5 pts.", (int)dp.Position));
                                                                    comparisonErrors.Add(String.Format("  Correct line is \"{0}\"", model.OldText.Lines[i].Text));
                                                                    comparisonErrors.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                                                                    break;
                                                                case ChangeType.Inserted:
                                                                    if (dp.Text == String.Empty)
                                                                    {
                                                                        errorGrades += 5;
                                                                        comparisonErrors.Add(String.Format("Extra empty line at line # {0}. Minus 5 pts.", (int)dp.Position));
                                                                    }
                                                                    else if (dp.Text.Trim() == String.Empty)
                                                                    {
                                                                        errorGrades += 7;
                                                                        comparisonErrors.Add(String.Format("Extra line of blanks at line # {0}. Minus 7 pts.", (int)dp.Position));
                                                                    }
                                                                    else
                                                                    {
                                                                        errorGrades += 10;
                                                                        comparisonErrors.Add(String.Format("Extra line at line # {0}. Minus 10 pts.", (int)dp.Position));
                                                                        comparisonErrors.Add(String.Format("     Your Line is \"{0}\"", dp.Text));
                                                                    }
                                                                    break;
                                                                case ChangeType.Deleted:
                                                                case ChangeType.Imaginary:
                                                                    errorGrades += 10;
                                                                    comparisonErrors.Add(String.Format("Missing line at line # {0}. Minus 10 pts.", i + 1));
                                                                    comparisonErrors.Add(String.Format("     expected Line is \"{0}\"", model.OldText.Lines[i].Text));
                                                                    break;
                                                            }
                                                        }

                                                        if (errorGrades == 0)
                                                        {
                                                            grade_box.setAttribute("value", "100");
                                                            remarks_box.setAttribute("value", "OK");
                                                            stud.Send_Gmail("Your last submission was perfect!!", "Good job - " + stud.first_name, filesToAttach);
                                                        }
                                                        else
                                                        {
                                                            String comparisonErrorSingleString = String.Empty;
                                                            foreach (String comarison_line in comparisonErrors) comparisonErrorSingleString += (comarison_line + "\n");
                                                            grade_box.setAttribute("value", (100-errorGrades).ToString());
                                                            remarks_box.setAttribute("value", comparisonErrorSingleString);

                                                            filesToAttach[0] = randomInputFile;
                                                            filesToAttach[1] = studentOutputFile;
                                                            filesToAttach[2] = BenchmarkOutputFile;

                                                            String explenationLine = String.Format("Follwoing are the differneces to expected output. The input used to test is attached to this email at file \"{0}\". Your output is attached at file \"{1}\". Expected output is attached at file \"{2}\". Please fix program and upload project again to Moodle. Detailed differences between your output and the expected one are:\n {3}", randomFileName, studentOutputFileName, benchmarkOutputFileName,comparisonErrorSingleString);
                                                            stud.Send_Gmail("Your last submission was not correct. It run but did not give exactly the desired output", explenationLine, filesToAttach);
                                                        }
                            */

                        } // for (int r = 0; r < allTable.rows.length; r++)

                        doc.getElementById("id_savequickgrades").click();

                    }
                }
            } // if (doc.url.EndsWith(@"action=grading"))


        }
    }
}
