﻿using Microsoft.Office.Interop.Word;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace HWs_Generator
{
    public class GUI2 : GUI1
    {
        [DllImport("user32.dll")]
        static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_CLOSE = 0xF060;


        [DllImport("gdi32.dll")]
        static extern uint GetBkColor(IntPtr hdc);

        public enum GUI2_ARGS
        {
            ID,
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[Enum.GetNames(typeof(GUI2_ARGS)).Length];
            args[(int)GUI2_ARGS.ID] = id;
            return args;

        }

        public override RunResults test_Hw_by_assembly(object[] args, FileInfo executable)
        {
            int stud_id = (int)args[0];
            Student stud = Students.students_dic[stud_id];
            RunResults rr = new RunResults();
            Assembly studentApp = Assembly.LoadFile(executable.FullName);
            Type[] appTypes = studentApp.GetTypes();
            //studentApp.get
            if (appTypes.Length < 1)
            {
                rr.grade = 30;
                rr.error_lines.Add("No classes in code");
                return rr;
            }

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


            if (son_form == null)
            {
                rr.grade = 30;
                rr.error_lines.Add("No Form derivitive available in code");
                return rr;
            }

            Type[] constructor_param_types = { };
            ConstructorInfo desired_constructor = son_form.GetConstructor(constructor_param_types);

            if (desired_constructor == null)
            {
                int grade_lost = 50;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Could not find the empty constructor. Minus {0} points.", grade_lost));
                return rr;
            }

            Directory.SetCurrentDirectory(executable.Directory.FullName);
            //Application.cu
            Object[] constructor_params = { };
            form_to_run = (Form)desired_constructor.Invoke(constructor_params);


            ThreadStart ts = new ThreadStart(run_form_to_run);
            Thread th = new Thread(ts);
            th.Start();

            int tries = 10;
            while (!form_to_run.Visible) Thread.Sleep(1000);

            if (!form_to_run.Visible)
            {
                int grade_lost = 50;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Form was never opened. Minus {0} points.", grade_lost));
                return rr;
            }

            if (form_to_run.BackColor != SystemColors.Control)
            {
                int grade_cost = 25;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Wrong Background Color on initial state on form. Expected {0} but found {1}. Minus {2} points.", "SystemColors.Control", form_to_run.BackColor.ToString(), grade_cost));
                form_to_run.Close();
                return rr;
            }

            // check that labels are not seen or non existent or that text is empty
            String labelsText = getAllLabeShowingText();
            if (labelsText != String.Empty)
            {
                int grade_lost = 30;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("When form started found text \"{0}\" in Labels that were supposed to be not showing. Minus {1} points.", labelsText, grade_lost));
            }
            // check that combo box shows "Choose a country...";
            List<Control> comboBoxes = ScreenControlsByType(typeof(ComboBox));
            if (comboBoxes.Count < 1)
            {
                int grade_cost = 35;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Could not find any ComboBox in your Form. Minus {0} points.", grade_cost));
                form_to_run.Close();
                return rr;
            }
            if (comboBoxes.Count > 1)
            {
                int grade_cost = 35;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Found more then one ComboBox in your Form. Minus {0} points.", grade_cost));
                form_to_run.Close();
                return rr;
            }
            ComboBox cb = (ComboBox)comboBoxes[0];
            if (!cb.Visible)
            {
                int grade_cost = 35;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Your ComboBox is not Visible when Form launches. Minus {0} points.", grade_cost));
                form_to_run.Close();
                return rr;
            }
            if (!cb.Enabled)
            {
                int grade_cost = 35;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Your ComboBox is not Enabled when Form launches. Minus {0} points.", grade_cost));
                form_to_run.Close();
                return rr;
            }
            // check items in combobox
            ComboBox.ObjectCollection items = cb.Items;
            // get flags from file system
            String baseStr = new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            String flagsPath = baseStr + @"\Flags";
            DirectoryInfo flagsDin = new DirectoryInfo(flagsPath);
            FileInfo[] files = flagsDin.GetFiles("*.png");
            if (files.Length != items.Count)
            {
                int grade_cost = 25;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Number of .png files in Flags directory = {1} != {2} = number of items in ComboBox. Minus {0} points.", grade_cost, files.Length , items.Count));
                form_to_run.Close();
                return rr;
            }
            foreach (FileInfo f in files)
            {
                String name = f.Name.Substring(0, f.Name.Length - f.Extension.Length);
                if (!items.Contains(name)){
                    int grade_cost = 25;
                    rr.grade -= grade_cost;
                    rr.error_lines.Add(String.Format("Your ComboBox did not have the expected item \"{0}\". Minus {1} points.", name,grade_cost));
                    form_to_run.Close();
                    return rr;
                }
            }

            // check state of all relevant components...
            // here only the ComboBox should be visible
            foreach (Control c in form_to_run.Controls)
            {
                if (c == cb)
                {
                    String expectedCBText = "Choose a country...";
                    if (c.Text != expectedCBText)
                    {
                        int grade_cost = 10;
                        rr.grade -= grade_cost;
                        rr.error_lines.Add(String.Format("Your ComboBox did not have the expected text \"{1}\". Instead it was showing \"{2}\". Minus {0} points.", grade_cost, expectedCBText, c.Text));
                    }
                }
                else
                {
                    if (isReallyVisible(c))
                    {
                        int grade_cost = 15;
                        rr.grade -= grade_cost;
                        rr.error_lines.Add(String.Format("At Form first show you had an unexpected visible Control of type \"{1}\".. Minus {0} points.", grade_cost, c.GetType().ToString()));
                    }
                }
            }
            /*
                        Debug.WriteLine("3333");
                        Button b = (Button)ScreenControlsByText(form_to_run.Controls, random_start.ToString());
                        if (b == null)
                        {
                            int grade_lost = 30;
                            rr.grade -= grade_lost;
                            rr.error_lines.Add(String.Format("No \"counter\" button with text=random_start={0} found !!! Minus {1} points.", random_start, grade_lost));
                            form_to_run.Close();
                            return rr;
                        }

                        b.BackColor = SystemColors.Control;


                        Control hidder_disabler = null;
                        if ((int)args[(int)GUI1_ARGS.EXTRA_BUTTON_FORM] == 0)
                        {
                            String button_text;
                            if ((int)args[(int)GUI1_ARGS.EXTRA_DISABLE_HIDE] == 0) button_text = "Eraser";
                            else button_text = "Disabler";
                            hidder_disabler = ScreenControlsByText(form_to_run.Controls, button_text);
                            if (hidder_disabler == null)
                            {
                                int grade_lost = 30;
                                rr.grade -= grade_lost;
                                rr.error_lines.Add(String.Format("Could not find {0} button. Minus {1} points.", button_text, grade_lost));
                                form_to_run.Close();
                                return rr;
                            }
                        }
                        else
                        {
                            hidder_disabler = form_to_run;
                        }

                        Dictionary<Color, int> colorDicts = new Dictionary<Color, int>();
                        for (int i = random_start, clicks = 0; i > 0; i--)
                        {


                            // Some crazy shit code to Close down some MessageBox over ununderstandable exception...
                            IntPtr window = FindWindow(null, "Microsoft .NET Framework");
                            if (window != IntPtr.Zero)
                            {
                                MessageBox.Show("walla");
                                Debug.WriteLine("Window found, closing...");
                                SendMessage((int)window, WM_SYSCOMMAND, SC_CLOSE, 0);
                            }

                            if (!form_to_run.Visible)
                            {
                                int grade_lost = 20;
                                rr.grade -= grade_lost;
                                rr.error_lines.Add(String.Format("Form closed unexpectedly after {0} clicks. Minus {1} points.", clicks, grade_lost));
                                return rr;
                            }

                            Color colorBefore;
                            if ((int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND] == 0)
                            {
                                colorBefore = form_to_run.BackColor;
                            }
                            else
                            {
                                colorBefore = b.BackColor;
                            }
                            //MessageBox.Show("1");
                            click_control(b);
                            mouseClick_control(b);
                            //MessageBox.Show("2");


                            clicks++;
                            if (clicks == random_start) break;

                            if (!form_to_run.Visible)
                            {
                                int grade_lost = 20;
                                rr.grade -= grade_lost;
                                rr.error_lines.Add(String.Format("Form closed unexpectedly after {0} clicks. Minus {1} points.", clicks, grade_lost));
                                return rr;
                            }


                            Console.WriteLine("random_start={0}, clicks={1}, b.Text={2}, i={3} Visible={4}", random_start, clicks, b.Text, i, form_to_run.Visible);
                            if (b.Text.Trim() != (i - 1).ToString().Trim())
                            {
                                int grade_lost = 30;
                                rr.grade -= grade_lost;
                                rr.error_lines.Add(String.Format("\"counter\" button with wrong text after {0} clicks. Expected : {1} but found {2}. Minus {3} points.", clicks, random_start - clicks, b.Text, grade_lost));
                                Console.WriteLine(rr.error_lines.Last());
                                form_to_run.Close();
                                return rr;
                            }
                            int counter_from_button = int.Parse(b.Text);
                            int last_color_start_count = (int)args[(int)GUI1_ARGS.LAST_COLOR_STARTER];
                            if (counter_from_button == last_color_start_count)
                            {
                                Color[] temp2 = { Color.DarkBlue, Color.Yellow, Color.Violet };
                                Color benchmark = temp2[(int)args[(int)GUI1_ARGS.LAST_COLOR]];
                                Color color_found;
                                String control_name = "button";
                                if ((int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND] == 0)
                                {
                                    color_found = b.BackColor;
                                }
                                else
                                {
                                    control_name = "Form";
                                    color_found = form_to_run.BackColor;
                                }
                                if (!(benchmark == color_found))
                                {
                                    int grade_lost = 20;
                                    rr.grade -= grade_lost;
                                    rr.error_lines.Add(String.Format("When reaching counter={0} (after {1} clicks). {2} background color did not change to {3}. Found background to be {4}. Minus {5} points.",
                                        b.Text, clicks, control_name, benchmark.Name, color_found.ToString(), grade_lost));
                                    Console.WriteLine(rr.error_lines.Last());
                                }

                            }


                            Color colorAfter;
                            String changer;
                            if ((int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND] == 0)
                            {
                                changer = "Form";
                                colorAfter = form_to_run.BackColor;
                            }
                            else
                            {
                                changer = "Button";
                                colorAfter = b.BackColor;
                            }

                            Console.WriteLine("changer={0}, colorBefore={1}, ColorAfter={2}, b.Back={3}, form.Back={4}", changer, colorBefore.ToString(), colorAfter.ToString(), b.BackColor.ToString(), form_to_run.BackColor.ToString());

                            Debug.WriteLine("b.color=" + b.BackColor);
                            if ((i - 1) % 10 == 9)
                            {
                                if (colorBefore == colorAfter)
                                {
                                    int grade_lost = 30;
                                    rr.grade -= grade_lost;
                                    rr.error_lines.Add(String.Format("{0} background Color did not change when counter got to {1}.", changer, b.Text, grade_lost));
                                    Console.WriteLine(rr.error_lines.Last());
                                    form_to_run.Close();
                                    return rr;
                                }
                                if (!colorDicts.ContainsKey(colorAfter)) colorDicts[colorAfter] = 0;
                                colorDicts[colorAfter]++;
                                if (colorDicts[colorAfter] > 2)
                                {
                                    int grade_lost = 20;
                                    rr.grade -= grade_lost;
                                    rr.error_lines.Add(String.Format("Same color ({3}) appeared in background of {0} already for 3'rd time - Not random like!!! when counter got to {1}. Minus {2} points.", changer, b.Text, grade_lost, colorAfter.ToString()));
                                    Console.WriteLine(rr.error_lines.Last());
                                    form_to_run.Close();
                                    return rr;
                                }
                            }
                            else
                            {
                                if (colorBefore != colorAfter)
                                {
                                    int grade_lost = 30;
                                    rr.grade -= grade_lost;
                                    rr.error_lines.Add(String.Format("{0} background Color changed unexpectedly when counter got to {1}.", changer, b.Text, grade_lost));
                                    Console.WriteLine(rr.error_lines.Last());
                                    form_to_run.Close();
                                    return rr;
                                }
                            }


                            bool test_seif_9 = (r.Next(0, 10) < 6);
                            if (test_seif_9)
                            {
                                if (hidder_disabler == form_to_run)
                                {
                                    click_control(form_to_run);
                                    mouseClick_control(form_to_run);

                                    if ((int)args[(int)GUI1_ARGS.EXTRA_DISABLE_HIDE] == 0)
                                    {
                                        if (b.Visible == true)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button did not disaapear as expected. Minus {0} points.", grade_lost));
                                            Console.WriteLine(rr.error_lines.Last());
                                            form_to_run.Close();
                                            return rr;
                                        }
                                        if (b.Enabled == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button was unexpectedly Disabled. Minus {0} points.", grade_lost));
                                            Console.WriteLine(rr.error_lines.Last());
                                            form_to_run.Close();
                                            return rr;
                                        }

                                        //MessageBox.Show("1");
                                        click_control(form_to_run);
                                        mouseClick_control(form_to_run);

                                        //MessageBox.Show("2");

                                        if (b.Visible == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button did not resaapear as expected. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }
                                        if (b.Enabled == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button was not ReEnabled. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }

                                    }
                                    else
                                    {
                                        if (b.Visible == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button disaapeared unexpectedly. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }
                                        if (b.Enabled == true)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button did not Disable as expected. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }

                                        click_control(form_to_run);
                                        mouseClick_control(form_to_run);

                                        MySleep(1000);

                                        if (b.Visible == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button did not resaapear as expected. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }
                                        if (b.Enabled == false)
                                        {
                                            int grade_lost = 20;
                                            rr.grade -= grade_lost;
                                            rr.error_lines.Add(String.Format("Counter button was not ReEnabled. Minus {0} points.", grade_lost));
                                            form_to_run.Close();
                                            return rr;
                                        }
                                    }

                                }
                                else
                                {
                                    ((Button)hidder_disabler).PerformClick();
                                }
                            }
                        }


                        MySleep(1000);
                        Debug.WriteLine(form_to_run.Visible.ToString());
                        if (form_to_run.Visible)
                        {
                            int grade_lost = 20;
                            rr.grade -= grade_lost;
                            rr.error_lines.Add(String.Format("Form did ont close as expected although counter reached 0. Minus {0} points.", grade_lost));
                        }

            */
            form_to_run.Close();
            return rr;
        }

        // really visible means Visible and 
        // whose bitmap pixels are not all the Form background
        private bool isReallyVisible(Control ctrl)
        {
            if (!ctrl.Visible) return false;
            Bitmap bmp = new Bitmap(ctrl.Width, ctrl.Height);
            ctrl.DrawToBitmap(bmp, ctrl.ClientRectangle);
            bmp.Save(ctrl.GetType().ToString() + ".png");
            if (ctrl.GetType() == typeof(PictureBox))
            {
                MessageBox.Show("Baasa");
            }
            for (int r = 0; r < bmp.Height; r++)
            {
                for (int c = 0; c < bmp.Width; c++)
                {
                    Color pixel = bmp.GetPixel(c, r);
                    if (pixel != form_to_run.BackColor)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private string getAllLabeShowingText()
        {
            String res = String.Empty;
            List<Control> ctrls = ScreenControlsByType(typeof(Label));
            foreach (Control c in ctrls)
            {
                if (c.Visible == false) continue;
                res += c.Text.Trim();
            }
            return res;
        }

        public override void Create_DocFile_By_Creators(object[] args, List<Creators> creators)
        {
            int id = (int)(args[0]);

            Student stud = Students.students_dic[id];
            String student_full_name = stud.first_name + " " + stud.last_name;


            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            Document wordDoc = oWord.Documents.Add();

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "שלום " + student_full_name;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "ש\"ב 2 נועדו לתרגל אתכם על שימוש בפקדים שלמדנו בהרצאה ובתרגול. על הפתרון שלכם לעמוד בדיוק בדרישות כדי שהבודק האוטומטי לא יכשיל אתכם.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "הפרויקט שתגישו יהיה כמובן Windows Forms Application כמו שהיה בש\"ב 1. ";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "כרגיל, יש להכין את הפרויקט, לכווץ ולהעלות אותו לאתר הקורס. ושוב, כרגיל - עם שאלות וכאלה תפנו אליי או אל אמיר.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "הבודק האוטמטי אמור לענות לכם עם ציון בתוך דקות ספורות מההגשה. אם לא חזרה תשובה או לא ברורה התשובה או כל שאלה - תודיעו לי שאוכל לבדוק מה \"נתקע\".";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "המלצתי האישית היא לבדוק (ואם צריך לתקן) את תוכניתכם לאחר ביצוע של כל אחד מהסעיפים הבאים:";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "1) שנו את הכותרת (Title) של הטופס ל-email שלכם (לפני השינוי הוא בטח יהיה Form1).";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "2) הוסיפו כפתור לטופס. אין דרישה מיוחדת לגבי המיקום של הכפתור או לגבי השם שתתנו לו (אם לא תשנו את שמו - הוא יהיה button1). בשלב זה גם אין חשיבות לטקסט שיופיע על הכפתור (ואם לא תשנו אותו הוא גם יהיה button1 - אבל בהמשך התרגיל הזה נשנה את הטקסט שעל הכפתור.).";
            par1.Range.InsertParagraphAfter();

            Button b = new Button();
            b.Text = "button1";
            b.Location = new System.Drawing.Point(75, 75);
            pictures_form = new Form();
            pictures_form.Size = new Size(300, 200);
            pictures_form.Text = stud.email;
            pictures_form.Controls.Add(b);

            /*
                        ThreadStart ts = new ThreadStart(run_picture_form);
                        Thread t = new Thread(ts);
                        t.Start();
            */
            pictures_form.Show();
            MySleep(1000);

            par1.Range.Text = ", .בשלב הזה הטופס אמור להראות דומה לתמונה הבאה (כמו שאמרתי, אין חשיבות לגודל הטופס ולמיקום הכפתור כל עוד רואים אותו כמובן!)";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);

            //par1.Range.InsertParagraphAfter();
            par1.Range.Text = "3) הוסיפו בנאי (constructor) נוסף לטופס שלכם. אם לא שיניתם את שם הטופס - ההוספה צריכה להיות בקובץ Form1.cs . הוסיפו למחלקת הטופס בנאי (בנוסף לבנאי הריק ש-Visual Studio ייצר עבורכם) גם בנאי שמקבל פרמטר יחיד מסוג int.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "4) שנו את הפונקציה Main שבקובץ Program.cs כך שתקרא לבנאי החדש שלכם (במקום לבנאי הריק שנקרא עכשיו). שילחו לבנאי שמצפה לפרמטר int איזשהוא מספר אקראי בתחום שבין 20 ל 50.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "5) בתוך הבנאי החדש שהוספתם - כיתבו קוד שמשנה את הטקסט שעל הכפתור למספר שנשלח לבנאי. כלומר על הכפתור בטופס יופיע המספר שאותו הגרלתם בפונקציה Main שבקובץ Program.cs .";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עכשיו הטופס שלכם ייראה ככה אחרי ההפעלה (זיכרו כי המספר בכפתור הוא אקראי ובכל הפעלה של התוכנית יופיע מספר אחר)...";
            par1.Range.InsertParagraphAfter();

            int random_start = r.Next(30, 51);
            while (random_start % 10 == 9 || random_start % 10 == 0)
            {
                random_start = r.Next(30, 51);
            }
            b.Text = random_start.ToString();


            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            add_form_picture(wordDoc, pictures_form);

            par1.Range.Text = "6) הוסיפו קוד לטופס כך שבכל לחיצה על הכפתור המספר שמופיע עליו ירד ב-1. כאשר המספר מגיע ל-0, על הטופס להיסגר. אתם יכולים לסגור את הטופס בעזרת הפונקציה ()Close של המחלקה Form. - כלומר, פשוט ע\"י כתיבת ()Close;";
            par1.Range.InsertParagraphAfter();

            String[] temp1 = { "הטופס", "הכפתור" };
            String bgrd1 = temp1[(int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND]];
            String bgrd2 = temp1[1 - (int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND]];

            par1.Range.Text = String.Format("7) הוסיפו קוד לטופס כך שאם אחרי לחיצה על הכפתור המספר על הכפתור מסתיים ב-9 (לדוגמא 9,19,29,39 וכו) {0} ישנה את צבע הרקע שלו לצבע אקראי כלשהוא. שימו לב שכדי ליצור צבע אקראי עליכם רק להגריל ערכים בתחום 0-255 למרכיבי האדום\\ירוק\\כחול של הצבע ולהשתמש בפונקציה Color.FromArgb. הרקע יישאר בצבעו החדש עד הפעם הבאה שהמספר על הכפתור יסתיים ב-9.", bgrd1);
            par1.Range.InsertParagraphAfter();


            int num_of_clicks = random_start % 10 + 1;
            par1.Range.Text = String.Format("כך שאחרי {0} הקלקות על הכפתור - הטופס יכול להראות בערך ככה: (זיכרו כי רקע של {1} התחלף לצבע אקראי)", num_of_clicks, bgrd1);
            par1.Range.InsertParagraphAfter();
            b.Text = (random_start - num_of_clicks).ToString();
            if ((int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND] == 0)
            {
                pictures_form.BackColor = Color.Orange;
                b.BackColor = SystemColors.Control;
            }
            else
            {
                b.BackColor = Color.Orange;
            }

            MySleep(2000);

            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            add_form_picture(wordDoc, pictures_form);


            Color[] temp2 = { Color.DarkBlue, Color.Yellow, Color.Violet };
            Color clr = temp2[(int)args[(int)GUI1_ARGS.LAST_COLOR]];
            String color_name = "Color." + clr.Name;
            int starter = (int)args[(int)GUI1_ARGS.LAST_COLOR_STARTER];
            par1.Range.Text = String.Format("8) הוסיפו קוד לטופס כך שכאשר המספר על הכפתור ירד ל-{0},{1} יישנה את צבעו ל-{2}. שימו לב שמעבר לכך לא צפויים שינויי צבע נוספים עד שהטופס צפוי להיסגר.", starter, bgrd2, color_name);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("כך שכאשר המספר על הכפתור יגיע ל-{0} הטופס ייראה בערך ככה. (שימו לב כי בשלב הזה {1} שינה את צבעו כבר מספר פעמים - לפי סעיף 7).", starter, bgrd1);
            par1.Range.InsertParagraphAfter();

            Color anotherRandomColor = Color.Green;
            b.Text = starter.ToString();
            if ((int)args[(int)GUI1_ARGS.CHANGE_FORM_BUTTON_BACKGROUND] == 0)
            {
                b.BackColor = clr;
                pictures_form.BackColor = anotherRandomColor;
            }
            else
            {
                b.BackColor = anotherRandomColor;
                pictures_form.BackColor = clr;
            }

            MySleep(2000);
            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            add_form_picture(wordDoc, pictures_form);

            pictures_form.BackColor = b.BackColor = SystemColors.Control;
            b.Text = random_start.ToString();
            MySleep(2000);

            if ((int)args[(int)GUI1_ARGS.EXTRA_BUTTON_FORM] == 0)
            {
                if ((int)args[(int)GUI1_ARGS.EXTRA_DISABLE_HIDE] == 0)
                {
                    par1.Range.Text = String.Format("9) הוסיפו כפתור נוסף איפושהוא בטופס. (שוב-אין חשיבות לגודלו\\מיקומו\\שמו). על הטקסט בכפתור להיות \"Eraser\" בכל פעם שלוחצים על הכפתור החדש - יש להעלים את הכפתור הראשון. בלחיצה הבאה על כפתור \"Eraser\" יש להחזיר את הכפתור הראשון להופעה. ושוב כל כלחיצה על כפתור \"Eraser\" מעליה או מחזירה את הכפתור הראשון. לידיעתכם - העלמה\\הופעה של Control ניתנים לביצוע ע\"י התכונה Visible של ה-Control.");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.Text = String.Format("אז יחד עם הכפתור החדש הטופס יכול להיראות בהתחלה:");
                    par1.Range.InsertParagraphAfter();


                    Button eraser_button = new Button();
                    eraser_button.Location = new System.Drawing.Point(200, 130);
                    eraser_button.Text = "Eraser";

                    pictures_form.Controls.Add(eraser_button);


                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי לחיצה על כפתור \"Eraser\" הוא ייראה כך");
                    par1.Range.InsertParagraphAfter();


                    b.Visible = false;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי עוד לחיצה על כפתור \"Eraser\" הוא שוב ייראה כך");
                    par1.Range.InsertParagraphAfter();

                    b.Visible = true;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);
                }
                else
                {
                    par1.Range.Text = String.Format("9) הוסיפו כפתור נוסף איפושהוא בטופס. (שוב-אין חשיבות לגודלו\\מיקומו\\שמו). על הטקסט בכפתור להיות \"Disabler\" בכל פעם שלוחצים על הכפתור החדש - יש לשנות את הכפתור הראשון למצב - Disabled. בלחיצה הבאה על כפתור \"Disabler\" יש להחזיר את הכפתור הראשון למצב - Enabled. ושוב  כל לחיצה על כפתור \"Disabler\" הופכת את המצב של הכפתור הראשון מ-Enabled ל-Disabled וההפך. ניתן לעשות זאת ע\"י שליטה על התכונה  Enabled של הכפתור הראשון..");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.Text = String.Format("אז יחד עם הכפתור החדש הטופס יכול להיראות בהתחלה:");
                    par1.Range.InsertParagraphAfter();

                    Button disabler_button = new Button();
                    disabler_button.Location = new System.Drawing.Point(200, 130);
                    disabler_button.Text = "Disabler";

                    pictures_form.Controls.Add(disabler_button);

                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי לחיצה על כפתור \"Disabler\" הוא ייראה כך");
                    par1.Range.InsertParagraphAfter();


                    b.Enabled = false;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי עוד לחיצה על כפתור \"Disabler\" הטופס שוב ייראה כך");
                    par1.Range.InsertParagraphAfter();

                    b.Enabled = true;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);
                }
                par1.Range.Text = "כמובן שהכפתור החדש נשאר פעיל לאורך כל חיי התוכנית (ולא רק במצב ההתחלתי כפי שהדגמתי בציורים)";
                par1.Range.InsertParagraphAfter();

            }
            else
            {
                if ((int)args[(int)GUI1_ARGS.EXTRA_DISABLE_HIDE] == 0)
                {
                    par1.Range.Text = String.Format("9) הוסיפו קוד לטופס כך שבכל פעם שלוחצים על הטופס (לא על הכפתור ! - על הטופס מחוץ לכפתור) - יש להעלים את הכפתור עם המספרים. בלחיצה הבאה על הטופס יש להחזיר את הכפתור להופעה. ושוב - לחיצה אחת מעלימה והלחיצה הבאה מחזירה וכן הלאה... לידיעתכם - העלמה\\הופעה של Control ניתנים לביצוע ע\"י התכונה Visible של ה-Control.");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.Text = String.Format("כך שאחרי לחיצה על הטופס (מהמצב ההתחלתי) הטופס ייראה");
                    par1.Range.InsertParagraphAfter();


                    b.Visible = false;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי עוד לחיצה על הטופס הוא שוב ייראה כך");
                    par1.Range.InsertParagraphAfter();

                    b.Visible = true;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);
                }
                else
                {
                    par1.Range.Text = String.Format("9) הוסיפו קוד לטופס כך שבכל פעם שלוחצים על הטופס (לא על הכפתור ! - על הטופס מחוץ לכפתור) - יש לשנות את הכפתור עם המספרים למצב Disabled. בלחיצה הבאה על הטופס יש להחזיר את הכפתור למצב Enabled. ושוב - אם נלחץ על הטופס הכפתור ייעלם ואם שוב נלחץ - הכפתור יחזור וכן הלאה. ניתן לעשות זאת ע\"י שליטה על התכונה  Enabled של הכפתור.");
                    par1.Range.InsertParagraphAfter();

                    par1.Range.Text = String.Format("כך שאחרי לחיצה על הטופס (מהמצב ההתחלתי) הטופס ייראה");
                    par1.Range.InsertParagraphAfter();

                    b.Enabled = false;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);

                    par1.Range.Text = String.Format("ואחרי עוד לחיצה על הטופס הוא שוב ייראה כך");
                    par1.Range.InsertParagraphAfter();

                    b.Enabled = true;
                    MySleep(2000);

                    par1.Range.Text = "XXXX";
                    par1.Range.InsertParagraphAfter();
                    add_form_picture(wordDoc, pictures_form);
                }

            }

            par1.Range.Text = "זהו.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "סיימתם כבר ?";
            par1.Range.InsertParagraphAfter();

            pictures_form.Close();
            MySleep(2000);


            object fileName = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
            object missing = Type.Missing;
            wordDoc.SaveAs(ref fileName,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Close();
            oWord.Quit();



        }
    }
}