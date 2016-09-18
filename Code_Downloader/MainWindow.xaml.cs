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
            
        }

        private void WebBrowser1_Navigated(object sender, NavigationEventArgs e)
        {
            Debug.WriteLine("Yofi");
        }

        private void WebBrowser1_LoadCompleted(object sender, NavigationEventArgs e)
        {
            Debug.WriteLine("Completed");
            HTMLDocument doc = WebBrowser1.Document as HTMLDocument;
            if (doc.url.EndsWith(@"login/index.php"))
            {
                doc.getElementById("username").setAttribute("value", "levytami");
                doc.getElementById("password").setAttribute("value", "yardena12");
                doc.getElementById("loginbtn").click();
            }
            if (doc.url.EndsWith(@"http://el1.netanya.ac.il/"))
            {
                WebBrowser1.Navigate(@"http://el1.netanya.ac.il/mod/assign/view.php?id=142263&action=grading");
            }
            if (doc.url.EndsWith(@"id=142263&action=grading"))
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

                    // check if there is a directry for this id
                    Student stud = Students.students_dic.Where(z => z.Value.email == email).FirstOrDefault().Value;

                    String folderPath = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions\HW0\" + stud.id;
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
                        grade_box.setAttribute("value", "40");
                        remarks_box.setAttribute("value", Compiler.errorReason);
                        continue;
                    }

                    String randomInputFilesFolder = new FileInfo(resulting_exe_path).DirectoryName + "//GeneratedInput";
                    if (!Directory.Exists(randomInputFilesFolder)) Directory.CreateDirectory(randomInputFilesFolder);


                    // Start testing....
                    HW0 hw0 = new HW0();
                    // 1) get the specific student values...
                    int[] args = hw0.LoadArgs(stud.id);

                    // 2) decide how many try tests
                    int num_of_tries = hw0.Num_Of_Test_Tries;
                    // foreach test - 
                    bool comparisonOK = true;
                    String[] filesToAttach = new String[3];
                    String compareFailedSummary = String.Empty;
                    for (int testNum = 1;testNum <= num_of_tries; testNum++)
                    {
                        // create random input file
                        String randomInputFile = randomInputFilesFolder + "//test" + testNum + ".txt";
                        if (File.Exists(randomInputFile))
                        {
                            File.Delete(randomInputFile);
                            System.Threading.Thread.Sleep(500);
                        }

                        hw0.createRandomInputFile(randomInputFile);

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
                        p.WaitForExit();
                        String studentOutputFile = randomInputFilesFolder + "//student_output" + testNum + ".txt";
                        File.WriteAllText(studentOutputFile, output);

                        // run through official HW to get output
                        TextReader oldInput = Console.In;
                        TextWriter oldOutput = Console.Out;
                        String BenchmarkOutputFile = randomInputFilesFolder + "//my_output" + testNum + ".txt";
                        using (StreamWriter sw = new StreamWriter(BenchmarkOutputFile, false))
                        {
                            Console.SetIn(new StreamReader(randomInputFile));
                            Console.SetOut(sw);
                            hw0.Create_HW(args, true);
                        }
                        Console.SetIn(oldInput);
                        Console.SetOut(oldOutput);
                        // compare and give feedback

                        String[] studentOutputLines = File.ReadAllLines(studentOutputFile);
                        String[] benchmarkOutputLines = File.ReadAllLines(BenchmarkOutputFile);

                        comparisonOK = true;
                        if (studentOutputLines.Length != benchmarkOutputLines.Length)
                        {
                            comparisonOK = false;
                            compareFailedSummary = String.Format("Failed on test number {0} on comparison. Student output file has {1} lines while benchmark has {2} lines\n", testNum, studentOutputLines.Length,benchmarkOutputLines.Length);
                            filesToAttach[0] = randomInputFile;
                            filesToAttach[1] = studentOutputFile;
                            filesToAttach[2] = BenchmarkOutputFile;
                            break;
                        }
                        for (int lineNum = 0; lineNum < benchmarkOutputLines.Length; lineNum++)
                        {
                            if (studentOutputLines[lineNum] != benchmarkOutputLines[lineNum])
                            {
                                comparisonOK = false;
                                compareFailedSummary = String.Format("Failed on test number {0} on comparison on line {1}\n",testNum,lineNum);
                                compareFailedSummary += "Student   line=" + studentOutputLines[lineNum] + "\n";
                                compareFailedSummary += "Benchmark line=" + benchmarkOutputLines[lineNum] + "\n";
                                break;
                            }
                        }
                        if (comparisonOK == false)
                        {
                            filesToAttach[0] = randomInputFile;
                            filesToAttach[1] = studentOutputFile;
                            filesToAttach[2] = BenchmarkOutputFile;
                            break;
                        }
                    } // for (int testNum = 0;testNum < num_of_tries; testNum++)
                    if (comparisonOK == false)
                    {
                        int grade = 60;
                        grade_box.setAttribute("value", "60");
                        remarks_box.setAttribute("value", compareFailedSummary);
                        stud.Send_Gmail("Your last submission was not correct", compareFailedSummary,filesToAttach);
                    }
                    else
                    {
                        grade_box.setAttribute("value", "100");
                        remarks_box.setAttribute("value", "OK");
                    }
                } // for (int r = 0; r < allTable.rows.length; r++)

                doc.getElementById("id_savequickgrades").click();
            } // if (doc.url.EndsWith(@"id=142263&action=grading"))


        }
    }
}
