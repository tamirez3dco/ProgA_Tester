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

namespace HWs_Generator
{
    class GUI2 : GUI1
    {
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
