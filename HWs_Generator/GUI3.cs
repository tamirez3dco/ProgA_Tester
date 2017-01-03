using Microsoft.Office.Interop.Word;
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
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace HWs_Generator
{
    public class GUI3 : GUI2
    {
        public override RunResults test_Hw_by_assembly(object[] args, FileInfo executable)
        {
            RunResults rr = new RunResults();

            try
            {
                int stud_id = (int)args[0];
                Student stud = Students.students_dic[stud_id];

                Assembly studentApp = Assembly.LoadFile(executable.FullName);
                Directory.SetCurrentDirectory(executable.Directory.FullName);


                // get my form
                Assembly myApp = Assembly.LoadFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\GUI3_Mine.exe");
                GUI3_GateButton_Comparer comp_form = new GUI3_GateButton_Comparer(myApp,args, rr);
                comp_form.ShowDialog();
                //File.Move(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\results.bin","benchmark.bin");
                IFormatter formatter = new BinaryFormatter();
                List<GUI3_GateButton_Comparer.GuiResults> ress;
                using (Stream stream = new FileStream(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\results.bin",
                    FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    ress = (List<GUI3_GateButton_Comparer.GuiResults>)formatter.Deserialize(stream);
                }

                // student control
                Assembly studApp = Assembly.LoadFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI3_Mine\GUI3_Mine\bin\Debug\GUI3_Mine.exe");
                GUI3_GateButton_Comparer stud_form = new GUI3_GateButton_Comparer(studentApp, ress, rr);
                stud_form.ShowDialog();
                File.Move("results.bin", "student.bin");

                return rr;

            }
            catch (Exception exc)
            {
                int gradeLost = 40;
                Logger.Log("Got excpetion on checking. " + exc.Message, this.GetType().Name);
                rr.grade -= gradeLost;
                rr.error_lines.Add(String.Format("Recieved the following exception when trying to check your work:{0}", exc.Message));
                return rr;
            }

        }

        public override void Create_DocFile(object[] args)
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

//            bool usePictureBox = (bool)args[(int)GUI2_ARGS.USE_PICTUREBOX];
//            String str1 = usePictureBox ? "PictureBox," : "";


            par1.Range.Text = String.Format(".ש\"ב 3 נועדו לתרגל אתכם על ירושת פקדים ושימוש ב-TableLayoutPanel.");
            par1.Range.InsertParagraphAfter();

//            par1.Range.Text = "הפעם אני שולח לכם פרויקט מותחל. הפרויקט הוא בעצם סתם פרויקט WindowsFormApplication שהוספתי לו תיקייה שמכילה תמונות של דגלים של מדינות. עליכם להשלים את הפרויקט ולפתח את הקוד שלו כך שיענה לדרישות המפורטות.";
//            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "כרגיל, יש להכין את הפרויקט, לכווץ ולהעלות אותו לאתר הקורס. ושוב, כרגיל - עם שאלות על הש\"ב הללו תפנו אליי. בשאלות כלליות לגבי C# תיפנו אליי או אל אמיר.";
            par1.Range.InsertParagraphAfter();

//            par1.Range.Text = "הבודק האוטמטי אמור לענות לכם עם ציון בתוך דקות ספורות מההגשה (האמת שבתרגיל זה התשובה יכולה לקחת קצת יותר דקות - בגלל השימוש בTimer הבדיקה לוקחת יותר זמן אבל לא הרבה דקות). אם לא חזרה תשובה או לא ברורה התשובה או כל שאלה - תודיעו לי שאוכל לבדוק מה \"נתקע\".";
//            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "המלצתי האישית היא לבדוק (ואם צריך לתקן) את תוכניתכם לאחר ביצוע של כל אחד מהסעיפים הבאים:";
            par1.Range.InsertParagraphAfter();

            SIDE side = (SIDE)args[(int)GUI3_ARGS.GATE_BUTTON_SIDE];
            Color disColor = (Color)args[(int)GUI3_ARGS.GATE_DIS_COLOR];
            par1.Range.Text = String.Format("1) כיתבו מחלקה חדשה בשם GateButton שתירש מ-Button.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("התכונה שתאפיין GateButton מכפתור רגיל היא של-GateButton מותר להיכנס רק מ{0}. כניסה מכל צד אחר אל הכפתור תגרום לצביעת הכפתור בצבע {1} ({2}) ולאחר כניסה \"אסורה\" שכזו לחיצות עליו לא יפעילו את ה-EventHandlers שרשומים לכפתור. שימו לב : בשום מקרה לא נהפוך את הכפתור ל-Disabled או לחבוי (Visible=false) כיון שאז לא נוכל יותר לקבל events מהכפתור וזה לא מתאים לנו לתרגיל הזה.", hebrew(side), hebrew(disColor), "Color." + disColor.Name);
            par1.Range.InsertParagraphAfter();

            int gate_width = (int)args[(int)GUI3_ARGS.GATE_PIX_WIDTH];
            par1.Range.Text = String.Format("כדי לסמן למשתמש שמותר להיכנס לכפתור רק מצד {0} אנו נצייר קווים בעובי של {1} פיקסלים בכל הצדדים האחרים של הכפתור בצבע {2}. לדוגמא, כפתור שכזה (שעוד לא נכנסנו אליו עם העכבר או שנכנסנו מהצד הנכון, כלומר מ{0}) ייראה:", hebrew(side), gate_width, hebrew(disColor));
            par1.Range.InsertParagraphAfter();


            pictures_form = new Form();
            
            SpecialButton spButton = new SpecialButton();
            spButton.Text = "specialButton1";
            spButton.Gate = side;
            spButton.Gate_color = disColor;
            spButton.Gate_width = gate_width;
            spButton.Size = new Size(150, 70);
            spButton.Location = new System.Drawing.Point(100, 100);

            pictures_form.Controls.Add(spButton);

            pictures_form.Show();
            MySleep(1000);

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, spButton);

            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("לעומת זאת, אם ניכנס לכפתור מאחד מהצדדים ה\"אסורים\" הכפתור יישתנה להראות:", hebrew(side), gate_width, hebrew(disColor));
            par1.Range.InsertParagraphAfter();

            spButton.myDisable = true;
            MySleep(2000);

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, spButton);

            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("ממצב צבוע זה, אם העכבר ייצא מתחומי הכפתור - נחזיר את הכפתור חזרה למצבו הגלוי כפי שנראה בתמונה הראשונה.", hebrew(side), gate_width, hebrew(disColor));
            par1.Range.InsertParagraphAfter();

            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("המלצתי לכם היא לבדוק את נכונות הכפתור שלכם על טופס ריק משלכם, כולל שינויי הצורה שהיתבקשו, שלחיצות על הכפתור במצבו המלא אכן לא ייקראו לEvent Handler. אני בוודאי אבדוק לכם את הכפתור לפני שאמשיך הלאה לבדוק את שאלה 2 ששם תשתמשו בכפתור שכתבתם.", hebrew(side), gate_width, hebrew(disColor));
            par1.Range.InsertParagraphAfter();

            wordDoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

            par1.Range.Text = String.Format("2) בשאלה זו נכתוב UserControl שיכיל מספר GateButtons (כאלו מהסעיף הקודם) TextBox ו-Label אחד (וכנראה גם Timer).");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("נקרא ל-UserControl החדש שניצור MegaButton והוא בעצם יתפקד כמו כפתור - רק שכדי ללחוץ על ה-MegaButton יהיה צורך ללחוץ על כל הכפתורים שב-MegaButton וזאת במגבלת זמן.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("כאשר המשתמש יילחץ על כל כפתורי ה-GateButton במסגרת הזמן שמוקצב לכך, ה-MegaButton יוציא ארוע חדש שייקרא MegaClick ומי שמועוניין (לדוגמא הטופס שיכיל את ה-MegaButton) יוכל לקבל את הארוע ולכתוב EventHandler עבור הארוע MegaClick כמו שאנו עושים לארוע Click בכפתור רגיל.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("מבחינת תצוגה, ה-MegaButton יהיה בעל רקע Color.DarkGray. בנוסף, הוא יהיה מסודר בטבלה על ידי שימוש ב-TableLayoutPanel. לטבלה יהיו {0} שורות ו-{1} עמודות. ה-TableLayoutPanel ייתפוס את כל גודל ה-MegaButton וזאת על ידי שימוש בתכונה Dock של ה-TableLAyoutPanel. בנוסף, כל הפקדים שב-MegaButton ייתפסו, כל פקד, את כל המקום (התא\\תאים) המוקצה\\מוקצים לו ב-TableLayoutPanel וזאת על ידי שימוש בתכונת Dock של הפקדים עצמם.",4,3);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("להלן דוגמא לאיך נראה ה-MegaButton");
            par1.Range.InsertParagraphAfter();

            MegaButton megaButton = new MegaButton(args);
            spButton.Size = new Size(300, 300);
            spButton.Location = new System.Drawing.Point(50, 200);
            pictures_form.Controls.Add(megaButton);

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, megaButton);

            par1.Range.Text = String.Format("השורה הראשונה תיהייה בגובה של 20 פיקסלים ותכיל TextBox שייתפוס את כל רוחב השורה. ניתן לעשות זאת על ידי הנחת ה-TextBox בתא השמאלי ביותר ושימוש בתכונת ColumnSpan של ה-TextBox כדי לציין כמה עמודות על ה-TextBox לתפוס בתוך ה-TableLayoutPanel.");
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("בשורות הבאות יופיע כפתור מהכפתורים המיוחדים (GateButton) של הסעיף הקודם, כפתור אחד בכל שורה ובכל עמודה בדיוק לפי הדוגמא שהופיעה בציור. כל השורות של הGateButtons תהיינה שוות בגובהן. כאמור כל כפתור GateButton יופיע בעמודה נפרדת ב-TableLayoutPanel. העמודות יהיו שוות ברוחבן. להבהרה - גובה ורוחב השורות והעמודות של ה-GateButton יגדלו ויקטנו כאשר ה-MegaButton כולו ייגדל ויקטן ולכן לא ניתן להגדיר אותם מראש בפיקסלים (בניגוד לגובה של השורה הראשונה שמכילה את ה-TextBox שגובהה 20 פיקסלים תמיד).", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("הטקסט על הכפתורי GateButton יהיה הטקסט שמופיע ב-TextBox ועוד מספר השורה של הכפתור - כפי שמופיע בציור. כמובן שאם המשתמש משנה את הטקסט שב-TextBox (תוך כדי ריצה) על הטקסט שבכפתורים להתעדכן בהתאם.", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("ל-UserControl MegaButton תהייה תכונה ציבורית Time שתזכור את פרק הזמן בשניות שבו צריך המשתמש ללחוץ על כל כפתורי ה-GateButton בכדי להפעיל את הארוע MegaClick.", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("כאשר המשתמש לוחץ על כל אחד מכפתורי ה-GateButton הרקע של הכפתור שנלחץ יהפוך לצהוב. יחד עם זה יופיע Label בתא השמאלי התחתון ב-TableLayoutPanel שיימנה כמה שניות נותרו כדי ללחוץ על שאר הכפתורים. על ה-Label בפינה השמאלית התחתונה לעדכן את עצמו בכל שנייה שעוברת (בעצם בכל שנייה על ה-Label להציג מספר קטן בשנייה)", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("לדוגמא, נניח שערך תכונת Time של ה-MegaButton הוא 13. ושהמשתמש לחץ ברגע זה על כפתור ה-GateButton העליון, אז ה-MegaButoon ייראה כך: ", 33);
            par1.Range.InsertParagraphAfter();

            megaButton.Time = 13;
            megaButton.clickFirstGateButton();

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, megaButton);

            par1.Range.Text = String.Format("ואם אחרי 5 שניות שעברו המשתמש לחץ על כפתור ה-GateButton התחתון, אז ה-MegaButton ייראה כך:", 33);
            par1.Range.InsertParagraphAfter();

            megaButton.Time = 8;
            megaButton.clickThirdGateButton();

            par1.Range.Text = "XXXX";
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.InsertParagraphAfter();
            Thread.Sleep(1000);
            add_form_picture(wordDoc, megaButton);

            par1.Range.Text = String.Format("אם המונה יגיע ל-0 לפני שהמשתמש סיים ללחוץ על כל כפתורי ה-GateButton - הLabel ייעלם, הרקע של כל בכפתורים יחזור לצבע Control (שהיה לפני שהפך אולי לצהוב).", 33);
            par1.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("אם המשתמש סיים ללחוץ על כל כפתורי ה-GateButton לפני שהמונה הגיע ל-0(מה שאומר שבעצם התבצעה לחיצה על ה-MegaButton), - הLabel ייעלם, הרקע של כל בכפתורים יחזור לצבע Control וגם יופעל ארוע MegaClick.", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("כלומר, בשני המקרים ה-MegaButton יחזור למצבו ההתחלתי כפי שצויר בציור הראשון של שאלה זו.", 33);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("זהו בעצם. בתרגיל זה אין חשיבות למה שתעשו עם הטופס של הפרויקט שלכם ואתם רשאים לעשות עם הטופס כרצונכם. אני ממליץ להשתמש בטופס כדי לבדוק את תקינות הפקדים שבניתם בשני הסעיפים. הבודק האוטמטי ייבדוק את נכונות הפקדים ללא תלות בטופס שתגישו. בהצלחה לכם", 33);
            par1.Range.InsertParagraphAfter();

            pictures_form.Close();

            object fileName = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
            object missing = Type.Missing;
            wordDoc.SaveAs(ref fileName,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Close();
            oWord.Quit();

        }



        static Color[] MyColors = { Color.Black, Color.Blue, Color.Red, Color.Pink, Color.Green};

        public String hebrew(Color c)
        {
            if (c == Color.Black) return "שחור";
            if (c == Color.Blue) return "כחול";
            if (c == Color.Yellow) return "צהוב";
            if (c == Color.Red) return "אדום";
            if (c == Color.Pink) return "ורוד";
            if (c == Color.Green) return "ירוק";
            return null;
        }

        public String hebrew(SIDE s)
        {
            switch (s)
            {
                case SIDE.DOWN:
                    return "למטה";
                case SIDE.LEFT:
                    return "צד שמאל";
                case SIDE.RIGHT:
                    return "צד ימין";
                case SIDE.UP:
                    return "למעלה";
            }
            return null;
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[Enum.GetNames(typeof(GUI3_ARGS)).Length];
            args[(int)GUI3_ARGS.ID] = id;
            args[(int)GUI3_ARGS.GATE_BUTTON_SIDE] = (SIDE)(r.Next(0, 4));
            args[(int)GUI3_ARGS.GATE_DIS_COLOR] = MyColors[r.Next(0, MyColors.Length)];
            args[(int)GUI3_ARGS.GATE_PIX_WIDTH] = r.Next(2,5)*2;
            args[(int)GUI3_ARGS.MEGA_PATTERN] = r.Next(0, 4);
            return args;

        }

    }
    public enum GUI3_ARGS
    {
        ID,
        GATE_BUTTON_SIDE,
        GATE_DIS_COLOR,
        GATE_PIX_WIDTH,
        MEGA_PATTERN
    }

}
