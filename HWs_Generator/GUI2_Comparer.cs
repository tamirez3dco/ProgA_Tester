using StudentsLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HWs_Generator
{
    //ToDo: How to handle students exceptions on forms ?
    //ToDo: Insert text deletion from end...

    public partial class GUI2_Comparer : Form
    {
        TEST_PHASES phase = TEST_PHASES.BEFORE_TEST;
        Form benchmark_form;
        Form students_form;
        TextBox s_tb,b_tb;
        object[] args;
        bool hide_dis_TextBox;
        bool hide_dis_chopButton;
        bool hide_dis_comboBox;
        bool use_pictureBox;
        public GUI2_Comparer(Form s, Form b, object[] _args, RunResults _rr)
        {
            InitializeComponent();
            benchmark_form = b;
            students_form = s;
            rr = _rr;
            args = _args;
            hide_dis_chopButton = (bool)args[(int)GUI2_ARGS.HIDE_DIS_CHOP_BUTTON];
            hide_dis_TextBox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_TEXTBOX];
            hide_dis_comboBox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_COMBOBOX];
            use_pictureBox = (bool)args[(int)GUI2_ARGS.USE_PICTUREBOX];
        }

        private void GUI2_Comparer_Load(object sender, EventArgs e)
        {
            try
            {
                students_form.Show();
            }
            catch (Exception exc)
            {
                MessageBox.Show("What the fuck" + exc.Message);
            }
            
            students_form.StartPosition = FormStartPosition.Manual;
            students_form.DesktopLocation = new Point(100, 100);
            students_form.Text = ((int)args[0]).ToString();
            benchmark_form.Show();
            benchmark_form.StartPosition = FormStartPosition.Manual;
            benchmark_form.DesktopLocation = new Point(600, 100);
            benchmark_form.Text = "BENCHMARK";
            this.WindowState = FormWindowState.Minimized;
            timer1.Tick += Timer1_Test_On_Show;
            timer1.Interval = 1000;
            timer1.Start();
        }

        private List<Control> getControlsByType(Form f, Type type, bool screenVisibilty)
        {
            List<Control> res = new List<Control>();
            foreach (Control c in f.Controls)
            {
                if (c.GetType() != type) continue;
                if (c.Visible || !screenVisibilty) res.Add(c);
            }
            return res;
        }

        private Control getSingleVisibleControlByType(Form f, Type t)
        {
            List<Control> list = getControlsByType(f, t, true);
            if (list.Count > 1)
            {
                int grade_cost = 20;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Found more then one ({1}) visible Controls of type {2}. Minus {0} points",
                    grade_cost, list.Count, t.Name));
                return null;
            }
            if (list.Count < 1) return null;
            return list[0]; 
        }
        private void CloseAll()
        {

            students_form.Close();
            benchmark_form.Close();

            this.Close();
        }

        bool CompareCB_Items(ComboBox.ObjectCollection sc, ComboBox.ObjectCollection bc)
        {
            bool res = true;
            if (sc.Count != bc.Count)
            {
                int grade_lost = 10;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Expected number of items in ComboBox = {1} != {2} = items found. Minus {0} points", grade_lost, bc.Count, sc.Count));
                res = false;
            }
            foreach (String s in bc)
            {
                if (!sc.Contains(s))
                {
                    rr.grade -= 5;
                    rr.error_lines.Add(String.Format("Missing item {1} in ComboBox. Minus {0} points", 5, s));
                    res = false;
                }
            }
            foreach (String s in sc)
            {
                if (!bc.Contains(s))
                {
                    rr.grade -= 5;
                    rr.error_lines.Add(String.Format("Redundant item {1} in ComboBox. Minus {0} points", 5, s));
                    res = false;
                }
            }
            return res;
        }
        public static List<Control> screenControlsByNotEmptyText(List<Control> list)
        {
            List<Control> res = new List<Control>();
            foreach (Control c in list) if (c.Text.Trim() != String.Empty) res.Add(c);
            return res;
        }
        public static string getAllControlsText(List<Control> list)
        {
            String res = String.Empty;
            foreach (Control c in list) res += c.Text;
            return res;
        }
        DateTime testStart;
        int timeFromTest()
        {
            return (int)((DateTime.Now - testStart).TotalMilliseconds);
        }
        private String getTestPhaseDesc()
        {
            String res = " " + phase.ToString();
            if (phase != TEST_PHASES.BEFORE_TEST)
            {
                res += ", " + timeFromTest() + " milliseconds into the test";
            }
            return res;
        }
        String[] controlsToText(List<Control> list)
        {
            String[] res = new String[list.Count];
            for (int i = 0; i < list.Count; i++) res[i] = list[i].Text;
            return res;
        }
        bool compareLabels()
        {
            bool res = true;
            List<Control> studentsLabels = getControlsByType(students_form, typeof(Label), true);
            List<Control> stud_vis_labels = screenControlsByNotEmptyText(studentsLabels);
            List<Control> benchmarkLabels = getControlsByType(benchmark_form, typeof(Label), true);
            List<Control> bench_vis_labels = screenControlsByNotEmptyText(benchmarkLabels);
            Debug.WriteLine(getTestPhaseDesc());
            String allMyLabels = getAllControlsText(bench_vis_labels);
            Debug.WriteLine("My=" + allMyLabels);

            if (allMyLabels.Contains("solved"))
            {
                phase = TEST_PHASES.TEST_SOLVED;
            }
            Debug.WriteLine("His=" + getAllControlsText(stud_vis_labels));
            if (bench_vis_labels.Count != stud_vis_labels.Count)
            {
                int grade_cost = 15;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Labels mismatch at phase {1}. Number of visible non empty labels expected={2}. Number of visible non empty labels found={3}. Total labels text expected = \"{4}\". Total labels text found=\"{5}\". Minus {0} points.",
                   grade_cost, getTestPhaseDesc(), bench_vis_labels.Count, stud_vis_labels.Count,getAllControlsText(bench_vis_labels), getAllControlsText(stud_vis_labels)));
                return false;
            }

            Dictionary<String, String> dic = new Dictionary<String, String>();
            String[] studs_strings = controlsToText(stud_vis_labels);

            foreach (Control c in bench_vis_labels)
            {
                int indexClosest = LevenshteinDistance.GetIndexOfClosest(studs_strings, c.Text);
                dic[c.Text.Trim()] = studs_strings[indexClosest].Trim();
            }
            
            foreach (String b_str in dic.Keys)
            {
                if (b_str.Replace(" ","") != dic[b_str].Replace(" ",""))
                {
                    String b_str1 = b_str.ToLower();
                    String s_str1 = dic[b_str].ToLower();
                    b_str1 = b_str1.Replace("your time is","").Trim();
                    s_str1 = s_str1.Replace("your time is", "").Trim();
                    b_str1 = b_str1.Replace(":", "").Trim();
                    s_str1 = s_str1.Replace(":", "").Trim();
                    b_str1 = b_str1.Replace("seconds", "").Trim();
                    s_str1 = s_str1.Replace("seconds", "").Trim();

                    int i1, i2;
                    bool b1 = int.TryParse(b_str1, out i1);
                    bool b2 = int.TryParse(s_str1, out i2);
                    if (b1 && b2)
                    {
                        if (Math.Abs(i1 - i2) <= 1) continue;
                    }

                    int firstCharOff = 0;
                    while (firstCharOff < Math.Min(b_str.Length, dic[b_str].Length) &&
                        b_str[firstCharOff] == dic[b_str][firstCharOff])
                    {
                        firstCharOff++;
                    }
                        

                    int grade_cost = 15;
                    rr.grade -= grade_cost;
                    rr.error_lines.Add(String.Format("Labels mismatch at phase {1}. expected=\"{2}\".  found=\"{3}\". (Inconcistency started at char # {6}. Total labels text expected = \"{4}\". Total labels text found=\"{5}\". Minus {0} points.",
                       grade_cost, getTestPhaseDesc(), b_str, dic[b_str], getAllControlsText(bench_vis_labels), getAllControlsText(stud_vis_labels),firstCharOff));
                    res = false;
                }
            }
            return res;
        }
        private void prepareCompare()
        {
            timer1.Tick += Timer1_Inside_Riddle_CompareAll;
            TimeSpan ts = DateTime.Now - this.testStart;
            //timer1.Interval = 500 - ts.Milliseconds + ts.Milliseconds > 400 ? 1000 : 0;
            timer1.Interval = r.Next(100, 400);
            timer1.Start();
        }
        public RunResults rr;
        ComboBox b_cb, s_cb;
        private void Timer1_Test_On_Show(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_Test_On_Show");
            timer1.Stop();
            timer1.Tick -= Timer1_Test_On_Show;

            bool closeAll = false;
            s_cb = (ComboBox)getSingleVisibleControlByType(students_form, typeof(ComboBox));
            if (s_cb == null)
            {
                rr.error_lines.Add(String.Format("Could not locate ComboBox on first show"));
                CloseAll();
                return;
            }
            b_cb = (ComboBox)getSingleVisibleControlByType(benchmark_form, typeof(ComboBox));
            if (!CompareCB_Items(s_cb.Items, b_cb.Items)) closeAll = true;
            if (!compareLabels()) closeAll = true;
            if (s_cb.Text != b_cb.Text)
            {
                int grade_lost = 10;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Incorrect text on ComboBox. At phase {1}, expected text=\"{2}\", found text=\"{3}\". Minus {0} points",
                    grade_lost,getTestPhaseDesc(),b_cb.Text,s_cb.Text));
            }

            if (closeAll) CloseAll();

            timer1.Tick += Timer1_Test_SelectFlag;
            timer1.Interval = 250;
            timer1.Start();
        }

        String item;
        Random r = new Random();
        private void Timer1_Test_SelectFlag(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_Test_SelectFlag");

            timer1.Stop();
            timer1.Tick -= Timer1_Test_SelectFlag;

            do
            {
                item = (String)b_cb.Items[r.Next(0, b_cb.Items.Count)];
                if (b_cb.SelectedItem == null) break;
            } while (item == b_cb.SelectedItem.ToString());

            //DateTime.Now
            b_cb.SelectedItem = item;
            s_cb.SelectedItem = item;

            imagesCompared = false;

            phase = TEST_PHASES.TEST_STARTED_CORRECT;
            testStart = DateTime.Now;
            prepareCompare();

        }

        //private bool testTextBox


        public bool compareFormOnly()
        {
            bool res = true;
            if ((students_form.BackgroundImage == null) && (benchmark_form.BackgroundImage != null))
            {
                int grade_lost = 20;
                rr.error_lines.Add(String.Format("Found unexpected image on form background ! At phase {1}. Minus {0} points",
                    grade_lost, getTestPhaseDesc()));
                return false;
            }
            if ((students_form.BackgroundImage != null) && (benchmark_form.BackgroundImage == null))
            {
                int grade_lost = 20;
                rr.error_lines.Add(String.Format("Missing image on form background ! At phase {1}. Minus {0} points",
                    grade_lost, getTestPhaseDesc()));
                return false;
            }
            if (students_form.ControlBox != benchmark_form.ControlBox)
            {
                int grade_lost = 10;
                rr.error_lines.Add(String.Format("ControlBox error.At phase {1}, Expected form.ControlBoxto be {2} but found {3}. Minus {0} points",
                    grade_lost, getTestPhaseDesc(), benchmark_form.ControlBox, students_form.ControlBox));
                return false;
            }
            return true;

        }


        bool imagesCompared = false;
        bool comparisonFinished = false;
        bool imagesComparison = false;
        public bool CompareAll()
        {
            bool res = true;
            Debug.WriteLine("111");
            if (!compareFormOnly()) return false;
            Debug.WriteLine("222");
            if (!imagesCompared) compareImages();
            if (comparisonFinished)
            {
                if (imagesComparison == false)
                {
                    return false;
                }
            }
            if (!compareLabels())
            {
                Debug.WriteLine("333");

                return false;
            }
            Debug.WriteLine("444");

            foreach (Control c in benchmark_form.Controls)
            {
                String name = c.GetType().Name;
                Debug.WriteLine("555"+name);

                if (c.GetType() == typeof(Label)) continue;
                Debug.WriteLine("666");

                Control stud_c = getSingleVisibleControlByType(students_form, c.GetType());
                Debug.WriteLine("777");

                if (c.GetType() == typeof(TextBox))
                {
                    Debug.WriteLine("888");

                    b_tb = (TextBox)c;
                    Debug.WriteLine("b_tb=" + b_tb);
                    if (b_tb.Visible && b_tb.Enabled)
                    {
                        if (b_tb.BackColor == Color.Red) phase = TEST_PHASES.TEST_STARTED_WRONG;
                        else phase = TEST_PHASES.TEST_STARTED_CORRECT;
                    }
                    s_tb = (TextBox)stud_c;
                }
                if (c.Visible)
                {
                    if (stud_c == null)
                    {
                        int grade_lost = 15;
                        rr.error_lines.Add(String.Format("Could not locate an (expected to be visible) {2} with text \"{3}\" ! At phase {1}. Minus {0} points",
                            grade_lost, getTestPhaseDesc(), name, c.Text));
                        res = false;
                    }
                    else
                    {
                        if (c.Enabled != stud_c.Enabled)
                        {
                            int grade_lost = 5;
                            rr.grade -= grade_lost;
                            rr.error_lines.Add(String.Format("The control {2} was found unexpectedly with Enabled={3}. Expected Enabled={4} ! At phase {1}. Minus {0} points",
                                grade_lost, getTestPhaseDesc(), name, stud_c.Enabled, c.Enabled));
                            res = false;
                        }
                        if (c.Text != stud_c.Text)
                        {
                            int grade_lost = 5;
                            rr.grade -= grade_lost;
                            rr.error_lines.Add(String.Format("The control {2} was found unexpectedly with Text={3}. Expected Text={4} ! At phase {1}. Minus {0} points",
                                grade_lost, getTestPhaseDesc(), name, stud_c.Text, c.Text));
                            res = false;
                        }
                        if (c.BackColor.ToArgb() != stud_c.BackColor.ToArgb())
                        {
                            int grade_lost = 5;
                            rr.grade -= grade_lost;
                            rr.error_lines.Add(String.Format("The control {2} was found unexpectedly with BackColor={3}. Expected BackColor={4} ! At phase {1}. Minus {0} points",
                                grade_lost, getTestPhaseDesc(), name, stud_c.BackColor.ToArgb(), c.BackColor.ToArgb()));
                            res = false;
                        }
                    }
                }
                else // Control is invisible in Benchmark - just check visibility...
                {
                    if (stud_c != null)
                    {
                        if (phase == TEST_PHASES.TEST_SOLVED)
                        {
                            int grade_lost = 0;
                            rr.error_lines.Add(String.Format("The control {2} was found unexpectedly VISIBLE! At phase {1}. However, due to wrong picture in Hws doc -> no points deducted.Minus {0} points",
                                grade_lost, getTestPhaseDesc(), name));
                        }
                        else
                        {
                            int grade_lost = 15;
                            rr.error_lines.Add(String.Format("The control {2} was found unexpectedly VISIBLE! At phase {1}. Minus {0} points",
                                grade_lost, getTestPhaseDesc(), name));
                            res = false;

                        }
                    }
                }
            }

            return res;
        }

        private bool compareImages()
        {
            //Image i_s, i_b;
            if (use_pictureBox)
            {
                PictureBox s_pb = (PictureBox)getSingleVisibleControlByType(students_form, typeof(PictureBox));
                if (s_pb == null)
                {
                    int grade_lost = 35;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add("Expected pictureBox was not found.");
                    return false;
                }
                i_s = s_pb.Image;
                if (i_s == null)
                {
                    i_s = s_pb.BackgroundImage;
                    if (i_s == null)
                    {
                        int grade_lost = 35;
                        rr.grade -= grade_lost;
                        rr.error_lines.Add("Expected image was not found in pictureBox.");
                        return false;

                    }
                }
                PictureBox b_pb = (PictureBox)getSingleVisibleControlByType(benchmark_form, typeof(PictureBox));
                //i_b = b_pb.Image;
            }
            else
            {
                i_s = new Bitmap(students_form.BackgroundImage);
                //i_b = benchmark_form.BackgroundImage;
            }

            Thread t = new Thread(images_compare_threaded);
            t.Start();
            imagesCompared = true;
            return true;

            FileInfo origFile = new FileInfo(@"../../Flags/" + item + ".png");
            Image origImage = Bitmap.FromFile(origFile.FullName);
            double similarity = StudentsLib.Imaging.getSimilarity(new Bitmap(origImage) , new Bitmap(i_s));
            if (similarity > 5)
            {
                i_s.Save("imageFound.png");
                FileInfo fin = new FileInfo("imageFound.png");
                rr.filesToAttach.Add(fin.FullName);
                rr.filesToAttach.Add(origFile.FullName);

                int grade_cost = 20;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("After clicking on item \"{1}\" expected different image then found. Expected image is attached in file \"{2}\", image found attached in the file \"{3}\". Minus {0} points.", 
                    grade_cost, item, origFile.Name, fin.Name));
                return false;
            }
            imagesCompared = true;
            return true;
        }

        Image i_s;
        private void images_compare_threaded()
        {
            
            FileInfo origFile = new FileInfo(@"../../Flags/" + item + ".png");
            Image origImage = Bitmap.FromFile(origFile.FullName);
            double similarity = StudentsLib.Imaging.getSimilarity(new Bitmap(origImage), new Bitmap(i_s));
            if (similarity > 5)
            {
                i_s.Save("imageFound.png");
                FileInfo fin = new FileInfo("imageFound.png");
                rr.filesToAttach.Add(fin.FullName);
                rr.filesToAttach.Add(origFile.FullName);

                int grade_cost = 20;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("After clicking on item \"{1}\" expected different image then found. Expected image is attached in file \"{2}\", image found attached in the file \"{3}\". Minus {0} points.",
                    grade_cost, item, origFile.Name, fin.Name));
                imagesComparison = false;
            }
            imagesComparison = true;
            comparisonFinished = true;
        }

        private void Timer1_Inside_Riddle_CompareAll(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_Inside_Riddle_CompareAll");
            timer1.Stop();
            timer1.Tick -= Timer1_Inside_Riddle_CompareAll;

            //Thread thread = new Thread(new ThreadStart(copyFormsToFiles));
            //thread.Start();
            copyFormsToFiles();

            if (!CompareAll())
            {
                rr.error_lines.Add("Test breaked at phase" + getTestPhaseDesc()+" after solving succesfully " + solved_riddles+" riddles");
                if (timeFromTest() < 5000)
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Test stopped very early. Under 5 seconds. Minus {0} points", grade_lost));
                }
                rr.filesToAttach.Add("student_form.jpg");
                rr.filesToAttach.Add("benchmark_form.jpg");
                CloseAll();
                return;
            }

            timer1.Interval = r.Next(100, 400);
            timer1.Tick += Timer1_DecideNext;
            timer1.Start();
        }

        private void copyFormsToFiles()
        {
            Bitmap s_b = new Bitmap(students_form.Width, students_form.Height);
            students_form.DrawToBitmap(s_b, new System.Drawing.Rectangle(System.Drawing.Point.Empty, s_b.Size));
            s_b.Save("student_form.jpg");

            Bitmap b_b = new Bitmap(benchmark_form.Width, benchmark_form.Height);
            benchmark_form.DrawToBitmap(b_b, new System.Drawing.Rectangle(System.Drawing.Point.Empty, b_b.Size));
            b_b.Save("benchmark_form.jpg");
        }

        private char getRandomChar()
        {
            return (char)('a' + r.Next(0, 'z' - 'a' + 1));
        }

        private void Timer1_AddNewLetter(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_AddNewLetter");

            timer1.Stop();
            timer1.Tick -= Timer1_AddNewLetter;

            char nextLetter = getRandomChar();
            String currentString = b_tb.Text;
            if (r.Next(0, 10) > 0 && currentString.Length < item.Length)
            {
                char nextCorrectLetter = item[currentString.Length];
                nextLetter = nextCorrectLetter;
            }
            if (r.Next(0, 2) == 0)
            {
                nextLetter = nextLetter.ToString().ToUpper()[0];
            }


            currentString += nextLetter;
            s_tb.Text += nextLetter;
            b_tb.Text += nextLetter;

            prepareCompare();
        }
        private void Timer1_ClickChop(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_ClickChop");

            timer1.Stop();
            timer1.Tick -= Timer1_ClickChop;

            Button s_chopB = (Button)getSingleVisibleControlByType(students_form,typeof(Button));
            Button b_chopB = (Button)getSingleVisibleControlByType(benchmark_form, typeof(Button));

            s_chopB.PerformClick();
            b_chopB.PerformClick();

            prepareCompare();
        }
        int solved_riddles = 0;
        private void Timer1_DecideNext(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_DecideNext");

            timer1.Stop();
            timer1.Tick -= Timer1_DecideNext;

            timer1.Interval = 200;

            if (phase == TEST_PHASES.TEST_SOLVED)
            {
                solved_riddles++;
                if (solved_riddles == 2)
                {
                    CloseAll();
                    return;
                }
                timer1.Tick += Timer1_Test_SelectFlag;
                timer1.Start();
                return;
            }

            if (hint_mouse_down)
            {
                timer1.Tick += Timer1_UnClickHint;
                timer1.Start();
                return;
            }

            String currentString = b_tb.Text;
            Debug.WriteLine("AAA:" + currentString);
            if (item.ToLower().StartsWith(currentString.ToLower())) // correct
            {
                Debug.WriteLine("BBB");
                switch (r.Next(0, 4))
                {
                    case 0:
                        Debug.WriteLine("DDD");

                        timer1.Tick += Timer1_ClickHint;
                        break;
                    default:

                        Debug.WriteLine("EEE");

                        timer1.Tick += Timer1_AddNewLetter;
                        break;
                }
            }
            else
            {
                Debug.WriteLine("CCC");
                switch (r.Next(0, 4))
                {
                    case 0:
                        Debug.WriteLine("GGG");

                        timer1.Tick += Timer1_ClickHint;
                        break;
                    case 1:
                        Debug.WriteLine("HHH");

                        timer1.Tick += Timer1_ClickChop;
                        break;
                    default:
                        timer1.Tick += Timer1_AddNewLetter;
                        break;
                }
            }
           
            timer1.Start();
        }

        public void do_event_control(String event_name, Control c, Form f)
        {
            EventInfo evClick = c.GetType().GetEvent(event_name);
            FieldInfo eventClick = typeof(Control).GetField("Event" + event_name, BindingFlags.NonPublic | BindingFlags.Static);
            object secret = eventClick.GetValue(null);
            // Retrieve the click event
            PropertyInfo eventsProp = typeof(Component).GetProperty("Events", BindingFlags.NonPublic | BindingFlags.Instance);
            EventHandlerList events = (EventHandlerList)eventsProp.GetValue(c, null);
            Delegate click = events[secret];
            if (click == null) return;
            MethodInfo click_method = click.GetMethodInfo();
            ParameterInfo[] click_params = click_method.GetParameters();

            MouseEventArgs ea = new MouseEventArgs(MouseButtons.Left, 1, 1, 1, 0);
            Object[] click_objects = { c, ea };
            click_method.Invoke(f, click_objects);
            //                MessageBox.Show("2");

        }
        bool hint_mouse_down = false;

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CompareAll();
        }

        private void Timer1_ClickHint(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_ClickHint");

            timer1.Stop();
            timer1.Tick -= Timer1_ClickHint;

            if (use_pictureBox)
            {
                PictureBox b_pb = (PictureBox)getSingleVisibleControlByType(benchmark_form, typeof(PictureBox));
                PictureBox s_pb = (PictureBox)getSingleVisibleControlByType(students_form, typeof(PictureBox));
                do_event_control("MouseDown", b_pb, benchmark_form);
                do_event_control("MouseDown", s_pb, students_form);
            }
            else
            {
                do_event_control("MouseDown", benchmark_form, benchmark_form);
                do_event_control("MouseDown", students_form, students_form);
            }
            hint_mouse_down = true;

            prepareCompare();
        }

        private void Timer1_UnClickHint(object sender, EventArgs e)
        {
            Debug.WriteLine("Timer1_UnClickHint");
            timer1.Stop();
            timer1.Tick -= Timer1_UnClickHint;

            if (use_pictureBox)
            {
                PictureBox b_pb = (PictureBox)getSingleVisibleControlByType(benchmark_form, typeof(PictureBox));
                PictureBox s_pb = (PictureBox)getSingleVisibleControlByType(students_form, typeof(PictureBox));
                do_event_control("MouseUp", b_pb, benchmark_form);
                do_event_control("MouseUp", s_pb, students_form);
            }
            else
            {
                do_event_control("MouseUp", benchmark_form, benchmark_form);
                do_event_control("MouseUp", students_form, students_form);
            }
            hint_mouse_down = false;

            prepareCompare();
        }
    }

    public enum TEST_PHASES
    {
        BEFORE_TEST,
        TEST_STARTED_CORRECT,
        TEST_STARTED_WRONG,
        TEST_SOLVED,
    }
    public enum GUI2_ARGS
    {
        ID,
        HIDE_DIS_CHOP_BUTTON,
        HIDE_DIS_TEXTBOX,
        HIDE_DIS_COMBOBOX,
        USE_PICTUREBOX
    }
}
