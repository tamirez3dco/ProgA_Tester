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
    //ToDo: Make time labels forgivable for +-1 or check how to make accurate
    // 
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
            HIDE_DIS_CHOP_BUTTON,
            HIDE_DIS_TEXTBOX,
            HIDE_DIS_COMBOBOX,
            USE_PICTUREBOX
        }

        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[Enum.GetNames(typeof(GUI2_ARGS)).Length];
            args[(int)GUI2_ARGS.ID] = id;
            args[(int)GUI2_ARGS.HIDE_DIS_CHOP_BUTTON] = Convert.ToBoolean(r.Next(0,2));
            args[(int)GUI2_ARGS.HIDE_DIS_COMBOBOX] = Convert.ToBoolean(r.Next(0, 2));
            args[(int)GUI2_ARGS.HIDE_DIS_TEXTBOX] = Convert.ToBoolean(r.Next(0, 2));
            args[(int)GUI2_ARGS.USE_PICTUREBOX] = Convert.ToBoolean(r.Next(0, 2));
            return args;

        }


        public override void Create_DocFile_By_Creators(Object[] args, List<Creators> creators)
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

            bool usePictureBox = (bool)args[(int)GUI2_ARGS.USE_PICTUREBOX];
            String str1 = usePictureBox?"PictureBox,":"";

            par1.Range.Text = String.Format("ש\"ב 2 נועדו לתרגל אתכם על כתיבת GUI שכולל כמה מהפקדים שפגשתם, לדוגמא -  ComboBox, {0} TextBox, Label, Button  וכמובן Form. ", str1);
           //..על הפתרון שלכם לעמוד בדיוק(ואני מתכוון בדיוק - כמעט כל סטייה ברווח או אות קטנה\\גדולה נחשבת סטייה) בדרישות כדי שהבודק האוטומטי לא יכשיל אתכם.בהמשך התיאור מצורפים צילומי מסך של התוכנה המצופה מכם.התוכנה שלכם יכולה להיות שונה במיקומי הפקדים ובגודל ה - Font אבל לא מעבר לזה.
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "הפעם אני שולח לכם פרויקט מותחל. הפרויקט הוא בעצם סתם פרויקט WindowsFormApplication שהוספתי לו תיקייה שמכילה תמונות של דגלים של מדינות. עליכם להשלים את הפרויקט ולפתח את הקוד שלו כך שיענה לדרישות המפורטות.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "כרגיל, יש להכין את הפרויקט, לכווץ ולהעלות אותו לאתר הקורס. ושוב, כרגיל - עם שאלות על הש\"ב הללו תפנו אליי. בשאלות כלליות לגבי C# תיפנו אליי או אל אמיר.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "הבודק האוטמטי אמור לענות לכם עם ציון בתוך דקות ספורות מההגשה (האמת שבתרגיל זה התשובה יכולה לקחת קצת יותר דקות - בגלל השימוש בTimer הבדיקה לוקחת יותר זמן אבל לא הרבה דקות). אם לא חזרה תשובה או לא ברורה התשובה או כל שאלה - תודיעו לי שאוכל לבדוק מה \"נתקע\".";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "המלצתי האישית היא לבדוק (ואם צריך לתקן) את תוכניתכם לאחר ביצוע של כל אחד מהסעיפים הבאים:";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "0) שנו את תכונת הטופס ControlBox ל-false כך שהמשתמש לא יוכל להגדיל\\להקטין\\לסגור הטופס מהפינה הימנית העליונה של הטופס. את  ";
            par1.Range.InsertParagraphAfter();


            par1.Range.Text = "1) הוסיפו ComboBox לטופס Form1. שנו את תכונת ה-Text של ה-ComboBox ל-\"...Choose a country\"";
            par1.Range.InsertParagraphAfter();


            ComboBox cb = new ComboBox();
            cb.Text = "Choose a country...";
            cb.Location = new System.Drawing.Point(275, 75);
            pictures_form = new Form();
            pictures_form.ControlBox = false;
            pictures_form.Text = "Form1";
            pictures_form.Size = new Size(450, 350);
            pictures_form.Controls.Add(cb);

            /*
                        ThreadStart ts = new ThreadStart(run_picture_form);
                        Thread t = new Thread(ts);
                        t.Start();
            */
            pictures_form.Show();
            MySleep(1000);

            par1.Range.Text = "בשלב הזה הטופס אמור להראות דומה לתמונה הבאה (כמו שאמרתי, אין חשיבות לגודל הטופס ולמיקום ה ComboBox  כל עוד רואים אותו כמובן.!)";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);


            par1.Range.Text = "2) את הפריטים שיציג ה-ComboBox אתם תוסיפו באמצעות קוד שיסרוק את תיקיית Flags שהוספתי לכם ולכל תמונה שתימצאו שם תוסיפו את הפריט המתאים. כלומר, אם בתיקיה Flags ששלחתי לכם בפרויקט מופיעים הקבצים:";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "Israel.png";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "United States.png";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "Brasil.png";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "אז כאשר תיפתחו את ה-ComboBox הטופס צריך להיראות בערך ככה:";
            par1.Range.InsertParagraphAfter();

            cb.Items.Add("Israel");
            cb.Items.Add("United States");
            cb.Items.Add("Brasil");


            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_Desktop_App\2017\Patterns_docs\GUI2-combobox_open.png");


            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "מספר נקודות לגבי הפריטים:";
            par1.Range.InsertParagraphAfter();

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "על התוכנה שלכם לסרוק את התיקיה Flags בזמן הריצה ולפי הקבצים שהיא מוצאת להכניס איברים ל-ComboBox. הרעיון הוא שבזמן התכנות עוד לא ידוע לכם בדיוק אילו קבצים יהיו בתיקיה. אני אבדוק את זה ע\"י הכנסת קבצי דגלים אחרים לתיקיה בזמן הבדיקה.";

            Worder.Underline_in_doc(wordDoc, "בזמן הריצה");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "חשוב מאד! - הגישה אל התיקייה Flags ואל הקבצים שבה חייבת להיות יחסית למיקום של קובץ ה-executable שלכם (כדי שהבודק יוכל למצוא את הקבצים הללו). כלומר, כיון שה-executable שלכם נמצא בתיקיה bin//Debug ומכיון ששמתי לכם את התיקיה Flags בתוך התיקיה המרכזית של הפרויקט - הגישה אל התיקיה Flags תהייה ע\"י הנתיב (path)";

            par1.Range.InsertParagraphAfter();
            String text = "/../..\"Flags\"";
            par1.Range.Text = text;
            Worder.English_Format_By_Search(wordDoc, text);

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "והגישה(לדוגמא) אל הקובץ Brasil.png תהייה דרך הנתיב";
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            //par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;


            par1.Range.InsertParagraphAfter();
            String text2 = "/../..\"Flags/Brasil.png\"";
            par1.Range.Text = text2;
            Worder.English_Format_By_Search(wordDoc, text2);

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "אין חשיבות לסדר האיברים ב-ComboBox.";
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "האיברים ב-ComboBox מכילים רק את שם המדינה (ללא הסיומת של הקובץ שמכיל את התמונה).";

            par1.Range.InsertParagraphAfter();
            cb.SelectedItem = "Brasil";

            par1.Range.InsertParagraphAfter();
            PictureBox pb = new PictureBox();
            if (usePictureBox)
            {
                par1.Range.Text = "2) בואו נשתמש בתמונות. הוסיפו PictureBox לטופס. כוונו את תכונת ה-SizeMode ל-StretchImage. בהתחלה (כשהטופס רק עלה ולפני שהמשתמש בחר מדינה כשלהיא מה-ComboBox) ה-PictureBox יהיה ריק כי לא תציגו בו אף תמונה. אולם לאחר שהמשתמש בחר מדינה כלשהיא על ה-PictureBox להציג את התמונה של הדגל המתאים שקיבלתם בתיקיה Flags. כך ש(לדוגמא) אם המשתמש בחר ב-ComboBox ב-Brasil אז על הטופס להיראות כמו בתמונה הבאה:";
                pb.Location = new System.Drawing.Point(40, 40);
                pb.Size = new Size(150, 150);
                pb.SizeMode = PictureBoxSizeMode.StretchImage;
                pb.Image = Image.FromFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI2_Mine\GUI2_Mine\Flags\Brasil.png");
                pictures_form.Controls.Add(pb);
                
            }
            else
            {
                par1.Range.Text = "2) בואו נשתמש בתמונות. לאחר שהמשתמש בחר מדינה כלשהיא על ה-ConboBox עליכם להציג את התמונה של הדגל המתאים שקיבלתם בתיקיה Flags כתמונת הרקע של הטופס. שנו את התכונה BackgroundImage לתמונה הרצויה וכוונו את התכונה BackgroundImageLayout לערך Stretch כדי שהתמונה תיכנס במלואה לתוך הטופס. כך ש(לדוגמא) אם המשתמש בחר ב-ComboBox ב-Brasil אז על הטופס להיראות כמו בתמונה הבאה:";
                pictures_form.BackgroundImageLayout = ImageLayout.Stretch;
                pictures_form.BackgroundImage = Image.FromFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI2_Mine\GUI2_Mine\Flags\Brasil.png");
            }
            MySleep(1000);
            //MessageBox.Show("111");
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);


            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "בשלב הזה תוכלו לבדוק שאם תיבחרו מדינות אחרות ב-ConboBox, אז גם התמונה תתחלף. יותר מאוחר בתרגיל לא תוכלו לבצע את הבדיקה הזו כי ה-ComboBox לא יהיה מאופשר או שיהיה חבוי.";

            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();
            String hebrewName = "חבוי";
            String propName = "Visible";
            String propValueDis = "false";
            String propValueEn = "true";
            bool hide_dis_textBox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_TEXTBOX];
            if (hide_dis_textBox)
            {
                hebrewName = "לא מאופשר (disabled)";
                propName = "Enabled";
            }
            String comboStateInRiddle = "חבוי";
            String comboStatePropName = "Visible";
            bool hide_dis_combobox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_COMBOBOX];
            if (hide_dis_combobox)
            {
                comboStateInRiddle = "לא מאופשר";
                comboStatePropName = "Enabled";
            }

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("3) עכשיו נהפוך את הטופס שלנו ל-חידה. נוסיף לטופס TextBox שבו המשתמש ייצטרך להקליד את שם המדינה. לפני שהמשתמש בחר מדינה ב-ComboBox ה-TextBox יהיה {0} כלומר - תכונת ה-{1} תהיה {2}. לאחר שהמשתמש בחר מדינה והתמונה מוצגת נאפשר לו לכתוב ב-TextBox ע\"י שינוי התכונה {1} לערך {3}. מרגע שהמשתמש בחר המדינה כלשהיא ב-ComboBox והחידה התחילה (היפעלנו את ה-TextBox) יש להעביר את ה-ComboBox למצב {4} (כלומר לשנות את תכונת {5} של ה-ComboBox לערך false) כדי שהמשתמש לא יוכל להחליף חידה עד שלא סיים לפתור את זאת שכבר בחר.",
                hebrewName, propName, propValueDis, propValueEn, comboStateInRiddle, comboStatePropName);

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("עכשיו הטופס שלכם (אחרי בחירה של מדינה ב-ComboBox)  ייראה בערך כך:");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            if (hide_dis_combobox) cb.Enabled = false;
            else cb.Visible = false;

            TextBox tb = new TextBox();
            tb.Location = new System.Drawing.Point(100,210);
            tb.Width = 250;
            tb.Text = String.Empty;
            pictures_form.Controls.Add(tb);
            Thread.Sleep(1000);

            add_form_picture(wordDoc, pictures_form);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("4) הוסיפו Event_Handler ל-TextBox שייטפל בארוע של TextChanged. ב-Event Handler שתכתבו עליכם להחליט האם מה שכתב המשתמש ב-TextBox הוא התחלה של שם המדינה (כפי שהוא מופיע ב-ComboBox).");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("ההשוואה לא צריכה לקחת בחשבון האם המשתמש כתב באותיות קטנות או גדולות או במעורבב קטנות וגדולות.");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("אם המשתמש כתב התחלה שגויה - יש לשנות את הרקע של ה-TextBox לאדום. הרקע יישאר אדום כל עוד המשתמש לא תיקן את הכתוב ב-TextBox להתחלה של שם המדינה.");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("לדוגמא, אם המשתמש בחר ב-CheckBox במדינה Brasil והכניס ל-TextBox את ההתחלה bRA - ה-TextBox יישאר לבן כי זו באמת ההתחלה של המילה Brasil בהתעלמות מאותיות קטנות\\גדולות. אבל אם המשתמש הוסיף את האות U ועכשיו ב-TextBox מופיע המילה bRAU, אז הרקע של ה-TextBox יהפוך לאדום, כמו בתמונות הבאות:");

            tb.Text = "bRA";
            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);

            tb.BackColor = Color.Red;
            tb.Text = "bRAU";

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);


            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("5) נוסיף מונה זמן שיודיע למשתמש כמה זמן עבר מתחילת החידה (תחילת החידה נחשבת מהזמן שבו המשתמש בחר באיזושהיא מדינה ב-ComboBox). לצורך תזדקקו ל-Timer שיעזור לכם לספור את הזמן ולפקד מסוג Label שבו תדווחו על הזמן שעבר. אני מצפה שבכל שנייה שעוברת תשנו את הכתוב ב-Label. לדוגמא, אחרי 5 שניות מתחילת החידה יופיע ב-Label הכיתוב הבא: Your time is:5 seconds. כמובן שלפני שהחידה \"התחילה\" ה-Label שמציג את הזמן לא צריך להופיע כלל (כנראה ע\"י קביעת תכונת Visible שלו ל-false).");


            Label labelTimer = new Label();
            labelTimer.Location = new System.Drawing.Point(10, 10);
            labelTimer.Text = "Your time is:5 seconds";
            //labelTimer.Width = 300;
            labelTimer.AutoSize = true;
            pictures_form.Controls.Add(labelTimer);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);


            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("ספירת הזמן לא עוצרת גם כשהרקע של ה-TextBox הוא אדום וגם כאשר הרקע הוא לבן.המונה ימשיך לספור את הזמן עד שהמשתמש יסיים לכתוב את המילה הנדרשת. מה שמביא אותנו לסעיף הבא.");

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("6) כאשר סוף סוף המשתמש סיים לכתוב את המילה נכון נעצור את מונה הזמן ובנוסף נודיע לו ע\"י Label חדש שהוא סיים לפתור את החידה וכמה טוב הוא עשה. ההודעה שנכתוב לו ב-Label החדש תהיה תלויה בכמה זמן לקח לו לפתור את החידה. ההודעה ב-Label תהיה מנוסחת באופן הבא:");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format(" ,(מילת תואר) you solved the word (המילה שנפתרה) in (הזמן שעבר) seconds");
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("כאשר (הזמן שעבר) מציין את מספר השניות שעברו מתחילת החידה");
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("כאשר (המילה שנפתרה) מציין את שם המדינה שהמשתמש בחר ב-ComboBox");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("כאשר (מילת תואר) תהייה:");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("Horrey - אם המשתמש סיים בתוך פחות מ-10 שניות");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("OK - אם המשתמש סיים ב10 עד 19 שניות");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("Baasa - אם המשתמש סיים ב 20 שניות או יותר");

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("לדוגמא, אם המשתמש סיים את החידה Brasil ב-24 שניות, הטופס ייראה בערך כך:");

            labelTimer.Text = "Your time is:24 seconds";
            pictures_form.Controls.Add(labelTimer);

            Label labelSolved = new Label();
            labelSolved.AutoSize = true;
            labelSolved.Location = new System.Drawing.Point(10, 270);
            labelSolved.Text = "Baasa, you solved the word Brasil in 24 seconds";
            labelSolved.Width = 1500;
            tb.Text = "bRaSiL";
            tb.BackColor = Color.White;
            pictures_form.Controls.Add(labelSolved);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);

            String whereMouseDown = "Form";
            if (usePictureBox) whereMouseDown = "PictureBox";
            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("7) רמז ראשון למשתמש - הוסיפו קוד שכאשר מבצעים MouseDown על ה-{0} ה-TextBox יציג את המילה הנכונה (זו שכתובה ב-ComboBox) וכאשר המשתמש יבצע MouseUp על ה-{0} תחזור המילה שהופיעה ב-TextBox לפני ה-MouseDown.", whereMouseDown);

            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("חשוב מאד מאד מאד !!! - לחיצת MouseDown שכזו לא צריכה לגרום לחידה להיפתר. לכן עליכם לדאוג שבמקרה של MouseDown, אפילו שהמילה הנכונה כתובה ב-TextBox עליכם להתעלם ממנה ולא להגיד למשתמש שהוא פתר נכון את החידה. אחת האפשרויות לעשות זאת (יש כמה וכמה כאלה) היא להוריד את ה-Event Handler שמטפל ב-TextChanged של ה-TextBox כאשר רוצים התעלמות שכזו. הורדת ה-Event Handler מתבצעת ע\"י:", whereMouseDown);

            par1.Range.InsertParagraphAfter();
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
            par1.Range.Text = String.Format("textBox1.TextChanged -= textBox1_TextChanged;");

            par1.Range.InsertParagraphAfter();
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("וכמובן להוסיף את ה-Event Handler מחדש אחרי שהחזרתם את המילה הנוכחית של המשתמש בסיום ה-MouseUp על ה-{0}.",whereMouseDown);
            //  Horrey, you solved the word " + comboBox1.SelectedItem.ToString() + " in " + seconds_passed + " seconds"

            bool hide_dis_chop_button = (bool)args[(int)GUI2_ARGS.HIDE_DIS_CHOP_BUTTON];
            String chop_state_correct = "חבוי";
            String propChop = "Visible";
            if (hide_dis_chop_button)
            {
                chop_state_correct = "לא מאופשר";
                propChop = "Enabled";
            }
            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("8) רמז שני למשתמש - הוסיפו כפתור חדש שהטקסט עליו יהיה \"Chop To Correct\". כל עוד המשתמש לא טועה (כל עוד הרקע של ה-TextBox הוא לבן), הכפתור החדש יהיה במצב {0}, כלומר תכונת {1} של הכפתור החדש תקבל ערך false.",chop_state_correct,propChop);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("הוסיפו קוד שדואג שבכל פעם שהמשתמש טועה (הרקע של ה-TextBox הוא אדום) הכפתור החדש עם הטקסט  \"Chop To Correct\") מופיע ומאופשר (Enabled). כאשר המשתמש לוחץ על הכפתור הזה, המילה שב-TextBox מתקצצת (יורדות אותיות מהסוף שלה) עד שהיא נהיית נכונה (עד שהיא נהיית באמת התחלה של המילה שב-ComboBox). לדוגמא, בתמונה הבאה מופיע הטופס במצב של טעות של המשתמש עם הכפתור החדש:");

            Button chopButton = new Button();
            chopButton.Location = new System.Drawing.Point(100, 240);
            chopButton.Text = "Chop To Correct";
            chopButton.Width = 100;
            tb.Text = "bRaUOSiL";
            tb.BackColor = Color.Red;
            labelSolved.Visible = false;
            labelTimer.Text = "Your time is:12 seconds";
            pictures_form.Controls.Add(chopButton);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("לאחר שהמשתמש ילחץ על הכפתור המילה שב-TextBox תתקצץ ל-bRa, הרקע של ה-TextBox יחזור להיות לבן וכפתור ה-Chop To Correct יחזור להיות {0} כמו בתמונה הבאה:",chop_state_correct);

            if (hide_dis_chop_button) chopButton.Enabled = false;
            else chopButton.Visible = false;
            tb.Text = "bRa";
            tb.BackColor = Color.White;
            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            Thread.Sleep(1000);
            add_form_picture(wordDoc, pictures_form);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("לסיכום מספר הבהרות:");

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("1) כאשר החידה פועלת (המונה פועל) על ה-TextBox להיות כמובן גלוי ומאופשר. כאשר המשתמש סיים לפתור את החידה - על ה-TextBox לחזור למצב {0}.", hebrewName);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("2) כאשר החידה פועלת (המונה פועל) והמשתמש שוגה (הרקע של ה-TextBox הוא אדום) על כפתור ה-Chop To Correct להיות במצב גלוי ומאופשר (Visible and Enabled). בכל מצב אחר עליו להיות במצב {0}.", chop_state_correct );

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("3) כאשר החידה פועלת (המונה פועל) על ה-ComboBox להיות במצב {0}. רק כאשר המתשתמש סיים לפתור את החידה ה-ComboBox חוזר להיות גלוי ומאופשר כדי לאפשר למשתמש להתחיל בחידה חדשה.", comboStateInRiddle);

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("4) כאשר המשתמש בוחר להתחיל חידה חדשה (לאחר שסיים לפתור את הקודמת) יש להעלים את ה-Label מסעיף 6 שמבשר על סיום החידה (כי התחלנו חידה חדשה) ויש לאפס את מונה הזמן ל-0 לפני שמתחילים שוב את ספירת הזמן.");

            par1.Range.InsertParagraphAfter();
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("זהו, מספיק, לא?");


            object fileName = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
            object missing = Type.Missing;
            wordDoc.SaveAs(ref fileName,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Close();
            oWord.Quit();
        }

        public static PictureBox myPb;
        public static Form myForm;

        public override RunResults test_Hw_by_assembly(object[] args, FileInfo executable)
        {
            RunResults rr = new RunResults();

            try
            {
                String chopString = "Chop To Correct";
                int stud_id = (int)args[0];
                bool hide_dis_chop_button = (bool)args[(int)GUI2_ARGS.HIDE_DIS_CHOP_BUTTON];
                bool hide_dis_textBox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_TEXTBOX];
                bool hide_dis_comboBox = (bool)args[(int)GUI2_ARGS.HIDE_DIS_COMBOBOX];
                Student stud = Students.students_dic[stud_id];

                Assembly studentApp = Assembly.LoadFile(executable.FullName);
                Type[] appTypes = studentApp.GetTypes();

                Directory.SetCurrentDirectory(executable.Directory.FullName);

                // get flags from file system
                String baseStr = new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
                String flagsPath = baseStr + @"\Flags";
                DirectoryInfo flagsDin = new DirectoryInfo(flagsPath);
                foreach (FileInfo f in flagsDin.GetFiles("*.png"))
                {
                    f.Delete();
                }
                do
                {
                    DirectoryInfo allOptionalFlags = new DirectoryInfo(@"D:\Tamir\Netanya_Desktop_App\2017\Patterns_docs\Flags");
                    foreach (FileInfo f in allOptionalFlags.GetFiles("*.png"))
                    {
                        if (r.Next(0, 3) == 0) f.CopyTo(flagsDin.FullName + @"\" + f.Name);
                    }

                } while (flagsDin.GetFiles("*.png").Length < 3);

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


                Object[] constructor_params = { };
                form_to_run = (Form)desired_constructor.Invoke(constructor_params);
                GUI2.myForm = form_to_run;

                // get my form
                Assembly myApp = Assembly.LoadFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI2_Mine\GUI2_Mine\bin\Debug\GUI2_Mine.exe");
                Type myFormType = myApp.GetType("GUI2_Mine.Form1");
                Type[] myConsTypes = { args.GetType() };
                ConstructorInfo my_constructor = myFormType.GetConstructor(myConsTypes);
                Object[] myParams = { args };
                Form myForm = (Form)my_constructor.Invoke(myParams);

                GUI2_Comparer comp_form = new GUI2_Comparer(form_to_run, myForm, args, rr);
                comp_form.ShowDialog();

                return rr;

            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                int gradeLost = 40;
                Logger.Log("Got excpetion on checking. " + exc.Message,this.GetType().Name);
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Recieved the following exception when trying to check your work:{0}",exc.Message));
                return rr;
            }
        }

        private void add_stack_frame(RunResults rr)
        {
            throw new Exception("RRR");
        }

        String timeStarter = "Your time is:";
        private bool check_time_labels(string labelsText, TimeSpan ts_till_now, ref RunResults rr)
        {
            add_stack_frame(rr);
            if (!labelsText.StartsWith(timeStarter))
            {
                int grade_cost = 25;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Expected time label to start with \"{1}\". Instead found label with text=\"{2}\". Minus {0} points.", grade_cost, timeStarter, labelsText));
                return false;
            }
            String lowerRemaining = labelsText.Replace(timeStarter, String.Empty).ToLower();
            if (lowerRemaining.StartsWith(" "))
            {
                int grade_cost = 25;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Incorrect time label format. Unexpected \" \" after the \"{1}\" in your label text=\"{2}\". Minus {0} points.", grade_cost, timeStarter, labelsText));
                return false;
            }
            if (!lowerRemaining.Trim().EndsWith("seconds"))
            {
                int grade_cost = 20;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Expected time label to end with \"{1}\". Instead found label with text=\"{2}\". Minus {0} points.", grade_cost, "seconds", labelsText));
                return false;
            }
            String onlyNumber = lowerRemaining.Replace("seconds", String.Empty).Trim();
            int number;
            if (!int.TryParse(onlyNumber,out number))
            {
                int grade_cost = 20;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Could not locate time number inside labels text=\"{1}\". Minus {0} points.", grade_cost, labelsText));
                return false;
            }
            if (Math.Abs(ts_till_now.TotalSeconds - number) > 2)
            {
                int grade_cost = 25;
                rr.grade -= grade_cost;
                rr.error_lines.Add(String.Format("Timing problem - expected to find time in label = +- {1}, istead found time label to havetext=\"{2}\". Minus {0} points.", grade_cost, (int)ts_till_now.TotalSeconds, labelsText));
                return false;
            }
            return true;
        }

        private void combined_click(Control c)
        {
            click_control(c);
            mouseClick_control(c);
        }

        private char getRandomChar()
        {
            return (char)('a' + r.Next(0, 'z' - 'a' + 1));
        }

        private List<Control> getVisibleControlsByType(Type type)
        {
            List<Control> res = new List<Control>();
            List<Control> optionals = ScreenControlsByType(type);
            foreach(Control c in optionals)
            {
                if (c.Visible) res.Add(c);
            }
            return res;
        }

        private List<Control> getEnabledControlsByType(Type type)
        {
            List<Control> res = new List<Control>();
            List<Control> optionals = getVisibleControlsByType(type);
            foreach (Control c in optionals)
            {
                if (c.Enabled) res.Add(c);
            }
            return res;
        }

        // really visible means Visible and 
        // whose bitmap pixels are not all the Form background
        private bool isReallyVisible(Control ctrl)
        {
            if (!ctrl.Visible) return false;
            Bitmap bmp = new Bitmap(ctrl.Width, ctrl.Height);
            ctrl.DrawToBitmap(bmp, ctrl.ClientRectangle);
            bmp.Save(ctrl.GetType().ToString() + ".png");
/*
            if (ctrl.GetType() == typeof(PictureBox))
            {
                MessageBox.Show("Baasa");
            }
*/
            for (int r = 0; r < bmp.Height; r++)
            {
                for (int c = 0; c < bmp.Width; c++)
                {
                    Color pixel = bmp.GetPixel(c, r);
                    if (pixel.ToArgb() != form_to_run.BackColor.ToArgb())
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

    }
}
