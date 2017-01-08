using StudentsLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HWs_Generator
{
    public partial class GUI3_MegaButton_Comparer : Form
    {
        Assembly benchApp, studApp;
        object[] args;
        RunResults rr;
        public GUI3_MegaButton_Comparer()
        {
            InitializeComponent();
        }

        private List<Control> GetAllVisibleControlsByType(Control container, Type t)
        {
            List<Control> res = new List<Control>();
            foreach (Control c in container.Controls)
            {
                Type tt = c.GetType();
                bool cVis = c.Visible;

                if (c.GetType().Equals(t) && c.Visible) res.Add(c);
                if (c.Controls.Count > 0)
                {
                    res.AddRange(GetAllVisibleControlsByType(c, t));
                }
                    
            }

            res.OrderBy(o => o.Location.Y);
            return res;
        }

        /*
                private List<Control> GetAllGateButtons(Control container, List<Control> list, Type t)
                {
                    foreach (Control c in container.Controls)
                    {
                        if (c.GetType().Equals(t)) list.Add(c);
                        if (c.Controls.Count > 0)
                            list = GetAllGateButtons(c, list, t);
                    }


                    return list;
                }
        */
        int timeValue = 6;
        List<Control> benchGateButtons, studGateButtons;
        Form stud_mbcForm,bench_mbcForm;
        Control studCtrl,benchCtrl;
        Type stud_GateType, stud_MegaType, bench_GateType, bench_MegaType;
        TextBox benchTB, studTB;
        int benchMegaClicks, studMegaClicks, benchMegaFlushes;
        PropertyInfo benchTimeProp, studTimeProp;
        public bool GetReady()
        {
            Cursor.Position = new Point(1000, 1000);

            try
            {
                // myMegaChecker
                bench_GateType = GUI3_GateButton_Comparer.getClosestTypeByNameProximity(benchApp, "GateButton");
                bench_MegaType = GUI3_GateButton_Comparer.getClosestTypeByNameProximity(benchApp, "MegaButton");
                ConstructorInfo benchBanai = bench_MegaType.GetConstructor(new Type[] { typeof(object[]) });
                benchCtrl = (Control)benchBanai.Invoke(new object[] { args });
                benchTimeProp = bench_MegaType.GetProperty("Time");
                benchTimeProp.SetValue(benchCtrl, timeValue);
                benchCtrl.Dock = DockStyle.Fill;
                bench_mbcForm = new Form();
                bench_mbcForm.Controls.Add(benchCtrl);


                MethodInfo benchMi = this.GetType().GetMethod("benchMegaClicked", BindingFlags.NonPublic | BindingFlags.Instance);
                EventInfo benchEvent = bench_MegaType.GetEvent("MegaClick");
                Type t1 = benchEvent.EventHandlerType;
                Delegate handler = Delegate.CreateDelegate(t1, this, benchMi);
                benchEvent.AddEventHandler(benchCtrl, handler);

                MethodInfo benchFlushMi = this.GetType().GetMethod("benchMegaFlushed", BindingFlags.NonPublic | BindingFlags.Instance);
                EventInfo benchFlushEvent = bench_MegaType.GetEvent("MegaFlushed");
                Type t11 = benchFlushEvent.EventHandlerType;
                Delegate Flushhandler = Delegate.CreateDelegate(t11, this, benchFlushMi);
                benchFlushEvent.AddEventHandler(benchCtrl, Flushhandler);

                bench_mbcForm.Text = "Benchmark";
                bench_mbcForm.Show();

                benchTB = (TextBox)(GetAllVisibleControlsByType(benchCtrl, typeof(TextBox))[0]);
                benchGateButtons = GetAllVisibleControlsByType(benchCtrl, bench_GateType);
                foreach (Control c in benchGateButtons)
                {
                    c.BackColor = Color.White;
                    c.ForeColor = Color.Black;
                    c.Padding = c.Margin = new Padding(0, 0, 0, 0);
                    //                MethodInfo mouseLeave = bench_GateType.GetMethod("OnMouseLeave", BindingFlags.NonPublic | BindingFlags.Instance);
                    //                Object[] pars = { new EventArgs() };
                    //                mouseLeave.Invoke(c, pars);

                }

                // stud Mega Checker
                stud_GateType = GUI3_GateButton_Comparer.getClosestTypeByNameProximity(studApp, "GateButton");
                stud_MegaType = GUI3_GateButton_Comparer.getClosestTypeByNameProximity(studApp, "MegaButton");
                ConstructorInfo studBanai = stud_MegaType.GetConstructor(new Type[0]);
                studCtrl = (Control)studBanai.Invoke(new object[0]);
                studCtrl.Dock = DockStyle.Fill;
                stud_mbcForm = new Form();
                stud_mbcForm.Controls.Add(studCtrl);
                stud_mbcForm.Text = "Student - " + (int)args[0];
                stud_mbcForm.SetDesktopLocation(bench_mbcForm.DesktopLocation.X + bench_mbcForm.Width + 10, bench_mbcForm.DesktopLocation.Y);
                stud_mbcForm.StartPosition = FormStartPosition.Manual;
                studTimeProp = stud_MegaType.GetProperty("Time");
                if (studTimeProp == null)
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Could not find Property \"Time\" in your MegaButton. Minus {0} points", gradeLost));
                    return false;
                }
                studTimeProp.SetValue(studCtrl, timeValue);
                studGateButtons = new List<Control>();

                MethodInfo studMi = this.GetType().GetMethod("studMegaClicked", BindingFlags.NonPublic | BindingFlags.Instance);
                EventInfo studEvent = stud_MegaType.GetEvent("MegaClick");
                Type t2 = studEvent.EventHandlerType;
                Delegate handler2 = Delegate.CreateDelegate(t2, this, studMi);
                studEvent.AddEventHandler(studCtrl, handler2);


                stud_mbcForm.Show();

                List<Control> optiobalTextBoxes = GetAllVisibleControlsByType(studCtrl, typeof(TextBox));
                if (optiobalTextBoxes.Count < 1)
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Could not find a visible TextBox in your MegaButton. Minus {0} points", gradeLost));
                    return false;
                }
                studTB = (TextBox)optiobalTextBoxes[0];
                studGateButtons = GetAllVisibleControlsByType(studCtrl, stud_GateType);
                foreach (Control c in studGateButtons)
                {
                    c.BackColor = Color.White;
                    c.ForeColor = Color.Black;
                    c.Padding = c.Margin = new Padding(0, 0, 0, 0);
                }
                before = DateTime.Now;
                operations.Add("Starting test with Time property=" + timeValue);

                Cursor.Position = new Point(1000, 1000);
                return true;

            }
            catch (Exception e)
            {
                int gradeLost = 30;
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Some terrible exception happened while testing your app. Exception is:{1}. Therefore, could not continue with test. Minus {0} points.", gradeLost, e.Message));
                return false;
            }
        }

        private void benchMegaClicked(object sender, EventArgs e)
        {
            Debug.WriteLine("Horrey");
            addToOperationsList("MegaButton Clicked!!!");
            benchMegaClicks++;
        }

        private void benchMegaFlushed(object sender, EventArgs e)
        {
            Debug.WriteLine("Flushed");
            addToOperationsList("MegaButton Flushed!!!");
            benchMegaFlushes++;
        }

        private void studMegaClicked(object sender, EventArgs e)
        {
            studMegaClicks++;
        }

        private void GUI3_MegaButton_Comparer_Load(object sender, EventArgs e)
        {
            
        }

        Random r = new Random();
        private void timer1_Resize(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Tick -= timer1_Resize;

            int newWidth = r.Next(250, 450);
            int newHeight = r.Next(100, 300);

            stud_mbcForm.Size = new Size(newWidth, newHeight);
            bench_mbcForm.Size = new Size(newWidth, newHeight);

            addToOperationsList(String.Format("Resized form to {0} x {1}",newWidth, newHeight));


            timer1.Tick += timer1_CheckAll;
            timer1.Start();
        }
        List<String> operations = new List<string>();
        private void timer1_ChangeText(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Tick -= timer1_ChangeText;
            String randomText = HW0.getRandomString();
            studTB.Text = benchTB.Text = randomText;
            addToOperationsList(String.Format("Changed text feild text to {0}",randomText));

            timer1.Tick += timer1_CheckAll;
            timer1.Start();
        }

        private void timer1_ClickRandomGateButton(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Tick -= timer1_ClickRandomGateButton;

            int randomButton = r.Next(0, studGateButtons.Count);
            Button b1 = (Button)studGateButtons[randomButton];
            b1.PerformClick();

            Button b2 = (Button)benchGateButtons[randomButton];
            b2.PerformClick();

            addToOperationsList(String.Format("Clicked on GateButton {0}", randomButton + 1));


            timer1.Tick += timer1_CheckAll;
            timer1.Start();
        }
        DateTime before;
        private void addToOperationsList(String str)
        {
            TimeSpan ts = DateTime.Now - before;
            operations.Add(String.Format("{0}:{1}.{2} - {3}",ts.Minutes,ts.Seconds,ts.Milliseconds,str));
        }

        int stepsCounter = 0;
        int step = 0;
        private void timer1_CheckAll(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Tick -= timer1_CheckAll;


            saveForms();

            bool err = false;
            Debug.WriteLine("StepCounter="+stepsCounter);
            if (!checkGateButtons())
            {
                Debug.WriteLine("checkGateButtons failed");
                err = true;
            }
            if (!checkLabels())
            {
                Debug.WriteLine("checkLabels failed");
                err = true;
            }
            if (!checkTextBoxes())
            {
                Debug.WriteLine("checkTextBoxes failed");
                err = true;
            }
            if (!checkMegaClicks())
            {
                Debug.WriteLine("checkMegaClicks failed");
                err = true;
            }

            if (err)
            {
                using (StreamWriter sw = new StreamWriter("operations.txt"))
                {
                    for(int i = 0; i < operations.Count; i++)
                    {
                        sw.WriteLine(operations[i]);
                    }
                }
                rr.filesToAttach.Add("MegaButtonForm_Benchmark" + stepsCounter + ".png");
                rr.filesToAttach.Add("MegaButtonForm_Student" + stepsCounter + ".png");
                rr.filesToAttach.Add("operations.txt");
                CloseAll();
                return;
            }

            if (stepsCounter++ > 20 && benchMegaFlushes > 0 && benchMegaClicks > 0)
            {
                CloseAll();
                return;
            }
            switch (r.Next(0, 5))
            {
                case 0:
                    timer1.Tick += timer1_Resize;
                    break;
                case 1:
                    timer1.Tick += timer1_ChangeText;
                    break;
                case 2:
                case 3:
                case 4:
                    timer1.Tick += timer1_ClickRandomGateButton;
                    break;
            }

            timer1.Start();
        }

        private bool checkTextBoxes()
        {
            return studTB.Text.Trim() == benchTB.Text.Trim();
        }

        private bool checkMegaClicks()
        {
            return studMegaClicks == benchMegaClicks;
        }

        private void saveForms()
        {
            Bitmap benchBmp = new Bitmap(bench_mbcForm.Width, bench_mbcForm.Height);
            bench_mbcForm.DrawToBitmap(benchBmp, new Rectangle(0, 0, bench_mbcForm.Width, bench_mbcForm.Height));
            benchBmp.Save("MegaButtonForm_Benchmark" + stepsCounter + ".png", ImageFormat.Png);

            Bitmap studBmp = new Bitmap(stud_mbcForm.Width, stud_mbcForm.Height);
            stud_mbcForm.DrawToBitmap(studBmp, new Rectangle(0, 0, stud_mbcForm.Width, stud_mbcForm.Height));
            studBmp.Save("MegaButtonForm_Student" + stepsCounter + ".png", ImageFormat.Png);

        }

        private void CloseAll()
        {
            stud_mbcForm.Close();
            bench_mbcForm.Close();
            this.Close();
        }

        private bool checkLabels()
        {
            Control studLabel = null, benchLabel = null;
            List<Control> benchLabels = GetAllVisibleControlsByType(benchCtrl, typeof(Label));
            if (benchLabels.Count > 0) benchLabel = benchLabels[0];
            if (benchLabel != null)
            {
                if (benchLabel.Text.Trim() == "") benchLabel = null;
            }
            List<Control> studLabels = GetAllVisibleControlsByType(studCtrl, typeof(Label));
            if (studLabels.Count > 0) studLabel = studLabels[0];
            if (studLabel != null)
            {
                if (studLabel.Text.Trim() == "") studLabel = null;
            }

            if (studLabel == null && benchLabel == null) return true;
            if (benchLabel == null)
            {
                if (studLabel.Text.Trim() == "0") return true;
                else
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Found unexpected counter label visible with text \"{1}\". Minus {0} points", gradeLost, studLabel.Text));
                    return false;
                }
            }
            if (studLabel == null)
            {
                if (benchLabel.Text.Trim() == "0") return true;
                else
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Could not find an expected visible counter label with text \"{1}\". Minus {0} points", gradeLost, benchLabel.Text));
                    return false;
                }
            }

            int benchCounter = int.Parse(benchLabel.Text);
            int studCounter;
            if (!int.TryParse(studLabel.Text,out studCounter))
            {
                int gradeLost = 20;
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Could not parse your label Text \"{1}\". Minus {0} points", gradeLost, studLabel.Text));
                return false;
            }

            if (Math.Abs(benchCounter - studCounter) > 1)
            {
                int gradeLost = 20;
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Mismath counters. Expected {1} but found {2}. Minus {0} points", gradeLost, benchCounter, studCounter));
                return false;
            }

            return true;
        }
        private bool checkGateButtons()
        {
            studGateButtons = GetAllVisibleControlsByType(studCtrl, stud_GateType);
            benchGateButtons = GetAllVisibleControlsByType(benchCtrl, bench_GateType);

            if (studGateButtons.Count != benchGateButtons.Count)
            {
                int gradeLost = 20;
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Wrong number of Visible GateButtons. Expected{1} but found {2}. Minus {0} points", gradeLost,benchGateButtons.Count, studGateButtons.Count));
                return false;
            }
            bool res = true;

            for (int i = 0; i < studGateButtons.Count; i++)
            {
                Control studB = studGateButtons[i];
                Control benchB = benchGateButtons[i];
                
                if (Math.Abs(benchB.Width - studB.Width) > 2)
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Wrong Width for GateButton {1}. Expected {2} but found {3}. Minus {0} points", gradeLost, i+1 ,benchB.Width , studB.Width));
                    res = false;
                }

                if (Math.Abs(benchB.Height - studB.Height) > 2)
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Wrong Height for GateButton {1}. Expected {2} but found {3}. Minus {0} points", gradeLost, i + 1, benchB.Height, studB.Height));
                    res = false;
                }

                if (benchB.Text.Trim() != studB.Text.Trim())
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Wrong Text for GateButton {1}. Expected {2} but found {3}. Minus {0} points", gradeLost, i + 1, benchB.Text, studB.Text));
                    res = false;
                }

                if ((benchB.BackColor == Color.Yellow && studB.BackColor != Color.Yellow) ||
                    (benchB.BackColor != Color.Yellow && studB.BackColor == Color.Yellow))
                {
                    int gradeLost = 20;
                    rr.grade -= gradeLost;
                    rr.error_lines.Add(String.Format("Wrong BackColor for GateButton {1}. Expected {2} but found {3}. Minus {0} points", gradeLost, i + 1, benchB.BackColor, studB.BackColor));
                    res = false;
                }

            }

            return res;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Tick -= timer1_Tick;

            if (!GetReady())
            {
                //MessageBox.Show("GetReady failed!!!");
                CloseAll();
                return;
            }

            timer1.Tick += timer1_Resize;
            timer1.Start(); 

        }

        public GUI3_MegaButton_Comparer(Assembly benchApp, Assembly studApp, object[] args, RunResults rr) : this()
        {
            this.benchApp = benchApp;
            this.studApp = studApp;
            this.args = args;
            this.rr = rr;
        }
    }
}
