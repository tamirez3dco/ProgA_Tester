using Microsoft.Office.Interop.Word;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HWs_Generator
{
    public class HW2 : HW1
    {
        public HW2()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW2";
            exampleRectangleSize = new Size(500, 900);
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[4];
            args[0] = id;
            args[1] = r.Next(0, 3); // Q1
            args[2] = r.Next(0, 2); // Q2
            args[3] = r.Next(3, 7); // Q3
            return args;

        }

        public override void Create_HW(Object[] args, bool real_input)
        {
            if (!real_input) Console.Clear();
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
                creators.Add(new Creators(Create_Q4, Create_Q4_doc, Create_Q4_Input));

                while (creators.Count > 0)
                {
                    int rndIdx = r.Next(0, creators.Count);
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

            par1.Range.Text = "ש\"ב 2 נועדו לתרגל אתכם על לולאות ומשפטי בקרה כפי שנלמדו בהרצאה ובתרגול. כרגיל, עליכם לייצר בדיוק את הפלט המצופה כדי שהבודק האוטומטי לא יכשיל אתכם. ושוב כרגיל, דוגמא לפלט המצופה מופיעה בסוף המסמך הזה.";
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

            par1.Range.Text = String.Format("תאריך הגשה אחרון - 4/12/2016 בשעה 23:55");
            par1.Range.InsertParagraphAfter();
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

        public new void Create_Q1(Object[] args, bool real_input)
        {
            int shape = (int)args[1];
            Console.WriteLine("Type shape size");

            int shapeSize = int.Parse(get_input_integer(real_input, 7));

            switch (shape)
            {
                case 0:
                    print_meshulash(shapeSize);
                    break;
                case 1:
                    print_meuyan(shapeSize);
                    break;
                case 2:
                    print_square(shapeSize);
                    break;
            }
        }

        public new String[] Create_Q1_Input(Object[] args)
        {
            String[] res = new String[1];
            res[0] = r.Next(15, 25).ToString();
            return res;
        }

        public new void Create_Q1_doc(Object[] args, Document wordDoc, int seif)
        {
            int shape = (int)args[1];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("{0}) מדפיסה את השורה:", seif);
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = String.Format("Type shape size");
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("     קולטת מספר מהמשתמש, מספר זה יהיה גודל הצורה אותה יש להדפיס.");
            par1.Range.InsertParagraphAfter();

            String[] shapes = { "משולש" , "מעוין" , "מרובע"};

            par1.Range.Text = String.Format("מדפיסה {0} בגודל שהוזן על ידי המשתמש. בסוף המסמך ישנה דוגמא ל{0} שכזה.", shapes[shape]);
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.InsertParagraphAfter();
        }

        public new void Create_Q2(Object[] args, bool real_input)
        {
            int operation = (int)args[2];
            Console.WriteLine("Type number");

            int number = int.Parse(get_input_integer(real_input, 13456885));
            int number_copy = number;
            int sum = 0;
            if (operation == 1) sum = 1;
            while (number != 0)
            {
                switch (operation)
                {
                    case 0:
                        sum += number % 10;
                        break;
                    case 1:
                        sum *= number % 10;
                        break;
                }
                number /= 10;
            }
            if (operation == 0) Console.WriteLine("The sum of all digits in the number {0} is {1}", number_copy, sum);
            else Console.WriteLine("The product of all digits in the number {0} is {1}", number_copy, sum);

        }

        public virtual String[] Create_Q2_Input(Object[] args)
        {
            String[] res = new string[1];
            int num = 0;
            for (int i = 0; i < 8; i++)
            {
                num = num * 10 + r.Next(1, 10);
            }
            res[0] = num.ToString();
            return res;
        }

        public virtual void Create_Q2_doc(Object[] args, Document wordDoc, int seif)
        {
            int operation = (int)args[2];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("{0}) מדפיסה את השורה:", seif);
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = String.Format("Type number");
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            if (operation == 0) par1.Range.Text = String.Format("     קולטת מספר מהמשתמש, מסכמת את סך כל הספרות במספר ומדפיסה את סכום זה.");
            else par1.Range.Text = String.Format("     קולטת מספר מהמשתמש, מכפילה את כל הספרות במספר ומדפיסה את תוצאת ההכפלה הזו.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("לדוגמא, אם המספר שהוקלט הוא 1845667 אז עליכם להדפיס:");
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            if (operation == 0) par1.Range.Text = String.Format("The sum of all digits in the number {0} is {1}", "1845667",37);
            else par1.Range.Text = String.Format("The product of all digits in the number {0} is {1}", "1845667", 8*4*5*6*6*7);
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }

        public new void Create_Q3(Object[] args, bool real_input)
        {
            int digiter = (int)args[3];
            Console.WriteLine("Type small number");
            int small_number = int.Parse(get_input_integer(real_input, 30));
            Console.WriteLine("Type big number");
            int big_number = int.Parse(get_input_integer(real_input, 70));

            for (int i = small_number; i <= big_number; i++)
            {
                int d1 = i % 10;
                int d2 = (i / 10) % 10;
                if ((i % digiter == 0) || (d1 == digiter) || (d2 == digiter)) Console.WriteLine(i);
            }
        }

        public new String[] Create_Q3_Input(Object[] args)
        {
            String[] res = new string[2];
            int small_num = r.Next(10234, 998654);
            res[0] = small_num.ToString();
            int big_num = small_num + r.Next(30, 50);
            res[1] = big_num.ToString();
            return res;
        }

        public new void Create_Q3_doc(Object[] args, Document wordDoc, int seif)
        {
            int digiter = (int)args[3];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("{0}) מדפיסה את השורה:", seif);
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = String.Format("Type small number");
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("     קולטת מספר מהמשתמש, נקרא למספר זה המספר הקטן.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("    מדפיסה את השורה:", seif);
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = String.Format("Type big number");
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("     קולטת מספר מהמשתמש, נקרא למספר זה המספר הגדול.");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("     לאחר מכן התוכנית מדפיסה את כל המספרים (מהמספר הקטן עד למספר הגדול) שמתחלקים ב{0} או שספרת האחדות שלהם או שספרת העשרות שלהם היא {0}",digiter);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("     לדוגמא אם המספר הקטן הוא 30 והמספר הגדול הוא 70 אז התוכנית תדפיס את", digiter);
            par1.Range.InsertParagraphAfter();

            float spaceAfter = par1.Format.SpaceAfter;
            WdLineSpacing lineSpacing_rule = par1.Format.LineSpacingRule;
            float lineSpacing = par1.Format.LineSpacing;

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            for (int i = 30; i <= 70; i++)
            {
                int d1 = i % 10;
                int d2 = (i / 10) % 10;
                if ((d1 == digiter) ||(d2 == digiter) || (i % digiter == 0))
                {
                    par1.Range.Text = i.ToString();
                    par1.Format.SpaceAfter = 0f;
                    par1.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                    par1.Format.LineSpacing = 10f;
                    par1.Range.InsertParagraphAfter();
                }
            }

            par1.Format.SpaceAfter = spaceAfter;
            par1.Format.LineSpacingRule = lineSpacing_rule;
            par1.Format.LineSpacing = lineSpacing;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
        }
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
            String[] inputs = { "9", "12.0", "-34", "45.00","101","-137.0","-66.0","32.0","126.5" };
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

    }
}
