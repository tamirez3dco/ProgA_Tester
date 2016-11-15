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

        String ClassName;
        String hw_entire_class_path;
        public MainWindow()
        {
            InitializeComponent();
            WebBrowser1.Navigate("http://el1.netanya.ac.il/login/index.php");
            String[] commandLineArgs = Environment.GetCommandLineArgs();
            String excel_file_path = commandLineArgs[1];
            ClassName = Environment.GetCommandLineArgs()[1];
            this.WindowState = WindowState.Minimized;
            Students students;
            typeToLinkDict = new Dictionary<Type, string>();
            switch (ClassName)
            {
                case "ProgrammingA_2017_Summer":
                    HW0.Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs_Summer";
                    students = new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\students_name_id_Shana_B.xlsx");
                    hw_entire_class_path = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions_Summer";
                    typeToLinkDict[typeof(HW0)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143049&action=grading";
                    typeToLinkDict[typeof(HW1)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143050&action=grading";
                    typeToLinkDict[typeof(HW2)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=142830&action=grading";
                    typeToLinkDict[typeof(HW3)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143026&action=grading";
                    typeToLinkDict[typeof(HW4)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=144124&action=grading";                   
                    break;
                case "EDP_2017":
                    GUI1.Students_All_Hws_dirs = @"D:\Tamir\Netanya_Desktop_App\2017\Students_HWs";                    
                    students = new Students(@"D:\Tamir\Netanya_Desktop_App\2017\Shana_B_2017.xlsx");
                    hw_entire_class_path = @"D:\Tamir\Netanya_Desktop_App\2017\Students_Submissions";
                    typeToLinkDict[typeof(GUI1)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=143956&action=grading";
                    break;
                case "ProgrammingA_2017":
                    students = new Students(@"D:\Tamir\Netanya_ProgrammingA\2017\Programming_A_Semester_A_2017.xlsx");
                    HW0.Students_All_Hws_dirs = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_HWs";
                    hw_entire_class_path = @"D:\Tamir\Netanya_ProgrammingA\2017\Students_Submissions";
                    typeToLinkDict[typeof(HW0)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=144329&action=grading";
                    break;
                case "Java1_2017_Highschool":
                    students = new Students(@"D:\Tamir\Netanya_Java_1\2017\Highschool\Highschool_Class.xlsx");
                    HW0.Students_All_Hws_dirs = @"D:\Tamir\Netanya_Java_1\2017\Highschool\Students_HWs";
                    hw_entire_class_path = @"D:\Tamir\Netanya_Java_1\2017\Highschool\Students_Submissions";
                    typeToLinkDict[typeof(JAVA0)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=146793&action=grading";
                    break;
                case "Java1_2017":
                    students = new Students(@"D:\Tamir\Netanya_Java_1\2017\SemesterA\SemesterA.xlsx");
                    HW0.Students_All_Hws_dirs = @"D:\Tamir\Netanya_Java_1\2017\SemesterA\Students_HWs";
                    hw_entire_class_path = @"D:\Tamir\Netanya_Java_1\2017\SemesterA\Students_Submissions";
                    typeToLinkDict[typeof(JAVA0)] = @"http://el1.netanya.ac.il/mod/assign/view.php?id=147049&action=grading";
                    break;

            }


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
                WebBrowser1.Navigate(typeToLinkDict[typeToLinkDict.Keys.First()]);
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
                            if (tc_email == null) continue;
                            if (tc_email.innerText == null) continue;
                            String email = tc_email.innerText.Trim();
                            Debug.WriteLine("email=" + email);
                            if (email == null) continue;

                            HTMLTableCell tc_last_update = tableRow.cells.item(7); // last upload time
                            String last_Update_Time = tc_last_update.innerText.Trim();
                            Debug.WriteLine("last update time=" + last_Update_Time);

                            if (last_Update_Time.Length < 10) continue;

                            // check if there is a directry for this hw
                            String hw_path = hw_entire_class_path + "/" + hw_name;
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

                            // Start testing....
                            Type type = types[j];
                            ConstructorInfo ctor = type.GetConstructor(Type.EmptyTypes);
                            HW0 hw = (HW0)ctor.Invoke(new object[] { });

                            String resulting_exe_path;
                            if (!hw.BuildProject(filePath, out resulting_exe_path))
                            {
                                grade_box.setAttribute("value", "30");
                                remarks_box.setAttribute("value", Compiler.errorReason);
                                stud.Send_Gmail(String.Format("Your last submission of {0} failed to build!!",hw_name), "Hi - " + stud.first_name + "\nSorry but the last project you uploaded to Moodle failed to build. Compilation error was:\n" + Compiler.errorReason + "\n\n\n. Please check your code and upload again!", new List<String>());
                                continue;
                            }


                            String randomInputFilesFolder = new FileInfo(resulting_exe_path).DirectoryName + "//GeneratedInput";
                            if (!Directory.Exists(randomInputFilesFolder)) Directory.CreateDirectory(randomInputFilesFolder);



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


                        } // for (int r = 0; r < allTable.rows.length; r++)

                        doc.getElementById("id_savequickgrades").click();

                    }
                }
            } // if (doc.url.EndsWith(@"action=grading"))


        }
    }
}
