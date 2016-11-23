using Microsoft.Office.Interop.Word;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace HWs_Generator
{
    public delegate void CreateQuestionFunc(Object[] args, bool real_input);
    public delegate String[] CreateInputFunc(Object[] args);
    public delegate void CreateDocFunc(Object[] args, Document wordDoc, int seif);

    public class Creators
    {
        public CreateQuestionFunc questFunc;
        public CreateDocFunc   docFunc;
        public CreateInputFunc inputFunc;

        public Creators (CreateQuestionFunc _qf, CreateDocFunc _df, CreateInputFunc _if)
        {
            questFunc = _qf;
            docFunc = _df;
            inputFunc = _if;
        }
    }
    public class HW1 : HW0
    {
        public static String get_input_integer(bool realinput, int num)
        {
            if (!realinput)
            {
                String res = num.ToString();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(res);
                Console.ForegroundColor = ConsoleColor.White;
                return res;
            }
            else return Console.ReadLine();
        }

        public HW1()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW1";
            exampleRectangleSize = new Size(450, 900);
        }

        public void Create_Q1(Object[] args, bool real_input)
        {
            Console.WriteLine("Please enter some integer number");

            String kelet1;
            if (real_input) kelet1 = Console.ReadLine();
            else
            {
                int randomNum = r.Next(10, 30);
                kelet1 = get_input_integer(false, randomNum);
            }
            int in1 = int.Parse(kelet1);
            Console.WriteLine("The number you entered is : {0}", in1);
            Console.WriteLine("Its predecessor is : {0}", in1 - 1);
            Console.WriteLine("Its successor is : {0}", in1 + 1);
        }

        public String[] Create_Q1_Input(Object[] args)
        {
            String[] res = new String[1];
            res[0] = r.Next(1, 5000).ToString();
            return res;
        }

        public void Create_Q1_doc(Object[] args, Document wordDoc, int seif)
        {
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("{0}) קולטת מספר שלם מה-Console ומדפיסה אותו, את הקודם לו ואת העוקב אחריו. לדוגמא, אם המספר שהקלדתם ב-Console הוא 7 - אז על התוכנית שלכם להדפיס את 3 השורות הבאות:", seif);
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            //par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            //par1.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Format.SpaceAfter = 0;
            //par1.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            //par1.Format.LineSpacing = 1f;
            par1.Range.Text = "The number you entered is : 7";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "Its predecessor is : 6";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "Its successor is : 8";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Empty;
            par1.Range.InsertParagraphAfter();
        }

        public void Create_Q2(Object[] args, bool real_input)
        {
            Console.WriteLine("Please enter your full name");
            String kelet_name = get_input_string(real_input, "Tamir Levy");
            Console.WriteLine("Please enter your age");
            int kelet_age = int.Parse(get_input_integer(real_input, 44));
            Console.WriteLine("Please enter your address (in one line)");
            String kelet_address = get_input_string(real_input, "Antar road, 6, Neve klumlum, Israel");
            Console.WriteLine(new String('*', kelet_age));
            Console.WriteLine("* Name:{0} *", kelet_name);
            Console.WriteLine("* Age:{0} *", kelet_age);
            Console.WriteLine("* Address:{0} *", kelet_address);
            Console.WriteLine(new String('*', kelet_age));
        }


        public void Create_Q3(Object[] args, bool real_input)
        {
            int num_of_numbers = (int)args[1];
            Console.WriteLine("Please enter {0} integer numbers", num_of_numbers);
            int sum = 0;
            for (int i = 0; i < num_of_numbers; i++)
            {
                sum += int.Parse(get_input_integer(real_input, r.Next(10, 40)));
            }
            // needs to add indoc eplenation about what if average comes out integer or with 1 decimal digit
            Console.WriteLine("Average of all {0} numbers is {1}", num_of_numbers, Math.Round((double)sum / num_of_numbers, 2));
        }

        public String[] Create_Q3_Input(Object[] args)
        {
            int num_of_numbers = (int)args[1];
            String[] res = new String[num_of_numbers];
            for (int i = 0; i < num_of_numbers; i++)
            {
                res[i] = r.Next(0, 100).ToString();
            }
            return res;
        }

        public void Create_Q3_doc(Object[] args, Document wordDoc, int seif)
        {

            int num_of_numbers = (int)args[1];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) קולטת {1} מספרים ומחשבת את הממוצע שלהם. עליכם להדפיס את הממוצע בדיוק של עד 2 מקומות אחרי הנקודה העשרונית. לצורך כך מומלץ להשתמש ב-.()Math.Round", seif, num_of_numbers);
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("לדוגמא, אם המספרים שקלטתם הם 36, 27 ו-20 אז עליכם להדפיס:");
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = "Average of all 3 numbers is 27.67";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = String.Format("(השימוש ב-Math.Round יגרום לכך שאם הממוצע יצא מספר שלם לא תודפס כלל הנקודה העשרונית וכן אם הממוצע ייצא עם ספרה אחת בלבד אחרי הנקודה, לא תודפס הספרה השנייה.)");
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Empty;
            par1.Range.InsertParagraphAfter();

/*
            // paint Math in green...
            Find findObject = wordDoc.Application.Selection.Find;
            findObject.Text = "Math";
            object replaceNone = WdReplace.wdReplaceNone;
            object missing = Type.Missing;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceNone, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Application.Selection.Font.Color = WdColor.wdColorBrightGreen;
*/
            wordDoc.Application.Selection.Collapse();

        }

        public void Create_Q4(Object[] args, bool real_input)
        {
            int up_down = (int)args[2];
            Console.WriteLine("Please enter 3 integer numbers");
            int num1 = int.Parse(get_input_integer(real_input, 16));
            int num2 = int.Parse(get_input_integer(real_input, 20));
            int num3 = int.Parse(get_input_integer(real_input, 12));

            // needs to add in doc hint how to get middle number
            int max = Math.Max(Math.Max(num1, num2), num3);
            int min = Math.Min(Math.Min(num1, num2), num3);
            int mid = (num1 + num2 + num3) - max - min;
            if (up_down == 0)
            {
                Console.WriteLine("minimal={0}", min);
                Console.WriteLine("middle={0}", mid);
                Console.WriteLine("maximal={0}", max);
            }
            else
            {
                Console.WriteLine("maximal={0}", max);
                Console.WriteLine("middle={0}", mid);
                Console.WriteLine("minimal={0}", min);
            }
        }

        public String[] Create_Q4_Input(Object[] args)
        {
            int num_of_numbers = 3;
            String[] res = new String[num_of_numbers];
            for (int i = 0; i < num_of_numbers; i++)
            {
                res[i] = r.Next(-100, 100).ToString();
            }
            return res;
        }

        public void Create_Q4_doc(Object[] args, Document wordDoc, int seif)
        {

            int order = (int)args[2];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            if (order == 0)
            {
                par1.Range.Text = String.Format("{0}) קולטת 3 מספרים מה-Console ומדפיסה אותם בסדר עולה. כלומר תחילה את המספר הנמוך ביותר, אחר כך את המספר האמצעי ובסוף את המספר הגבוהה ביותר. לדוגמא, אם המספרים שנקלטו הם 19, 25 ו-14 אז הפלט צריך להיות:.", seif);
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "minimal=14";
                par1.Range.InsertParagraphAfter();
                par1.Range.Text = "middle=19";
                par1.Range.InsertParagraphAfter();
                par1.Range.Text = "maximal=25";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }
            else
            {
                par1.Range.Text = String.Format("{0}) קולטת 3 מספרים מה-Console ומדפיסה אותם בסדר יורד. כלומר תחילה את המספר הגבוהה ביותר, אחר כך את המספר האמצעי ובסוף את המספר הנמוך ביותר. לדוגמא, אם המספרים שנקלטו הם 19, 25 ו-14 אז הפלט צריך להיות::", seif);
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "maximal=25";
                par1.Range.InsertParagraphAfter();
                par1.Range.Text = "middle=19";
                par1.Range.InsertParagraphAfter();
                par1.Range.Text = "minimal=14";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }


            Find findObject = wordDoc.Application.Selection.Find;
            findObject.Text = "רמז:";
            object replaceNone = WdReplace.wdReplaceNone;
            object missing = Type.Missing;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceNone, ref missing, ref missing, ref missing, ref missing);

            wordDoc.Application.Selection.BoldRun();
            wordDoc.Application.Selection.Collapse();

            par1.Range.Text = "רמז: בעזרת Math.Max קל למצוא את המספר הגבוהה ביותר - שימו לב שתצטרכו יותר מקריאה אחת ל-Math.Max כדי למצוא את הגבוהה מבין שלושה מספרים.";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "בעזרת Math.Min קל למצוא את המספר הנמוך ביותר - שימו לב שתצטרכו יותר מקריאה אחת ל-Math.Max כדי למצוא את הנמוך מבין שלושה מספרים.";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("ובכן, אם אתם יודעים כבר מהו המקסימום של השלושה ומהו המינימום של השלושה - אז סכום כל שלושת המספרים פחות סכום המקסימום והמינימום ייתן לכם את המספר האמצעי.");
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("");
            par1.Range.InsertParagraphAfter();
        }

        public void Create_Q5(Object[] args, bool real_input)
        {
            Console.WriteLine("Please enter 4 integer numbers as x1,y1,x2,y2:");
            int x1 = int.Parse(get_input_integer(real_input, r.Next(0, 10)));
            int y1 = int.Parse(get_input_integer(real_input, r.Next(0, 10)));
            int x2 = int.Parse(get_input_integer(real_input, r.Next(0, 10)));
            int y2 = int.Parse(get_input_integer(real_input, r.Next(0, 10)));
            double distance = Math.Sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2));
            Console.WriteLine("Distance from ({0},{1}) to ({2},{3}) is {4}", x1, y1, x2, y2,
                Math.Round(distance, 2));
        }

        public String[] Create_Q5_Input(Object[] args)
        {
            int num_of_numbers = 4;
            String[] res = new String[num_of_numbers];
            for (int i = 0; i < num_of_numbers; i++)
            {
                res[i] = r.Next(0, 100).ToString();
            }
            return res;
        }

        public void Create_Q5_doc(Object[] args, Document wordDoc, int seif)
        {

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) קולטת 4 מספרים (נקרא להם x1,y1,x2,y2) מה-Console. שני המספרים הראשונים ייצגו נקודה אחת במערכת הצירים ושני המספרים האחרונים מתוך ה-4 שקלטתם ייצגו נקודה שנייה במערכת הצירים. עליכם להדפיס את המרחק בין 2 הנקודות על ידי שימוש ב: ", seif);
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            wordDoc.Application.Selection.Collapse();
            InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW1-nusha.png");

            par1.Range.Text = "מומלץ כמובן להשתמש בשאלה ב-Math.Sqrt. יש להדפיס את המרחק בדיוק של 2 מקומות אחרי הנקודה העשרונית על ידי שימוש ב-Math.Round";
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "לדוגמא אם הקלטים היו 9,12,16,25 אז הפלט צריך להיות:";
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            double res = Math.Round(Math.Sqrt(49 + 13 * 13), 2);
            par1.Range.Text = String.Format("Distance from ({0},{1}) to ({2},{3}) is {4}", 9, 12, 16, 25, res);
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
        }

        public void Create_Q6(Object[] args, bool real_input)
        {
            int num_of_digits = (int)args[3];
            Console.WriteLine("Please enter one integer number with {0} digits:", num_of_digits);
            int num;
            if (num_of_digits == 3)
            {
                int[] options = { 258, 255, 913, 390, 761 };
                num = int.Parse(get_input_integer(real_input, options[r.Next(0, options.Length)]));
            }
            else
            {
                int[] options = { 2589, 2559, 9132, 3904, 7610 };
                num = int.Parse(get_input_integer(real_input, options[r.Next(0, options.Length)]));
            }
            int d0 = num % 10;
            int d1 = (num / 10) % 10;
            int d2 = (num / 100) % 10;
            int d3 = (num / 1000) % 10;
            bool res;
            if (num_of_digits == 3) res = (d0 < d1 && d1 < d2) || (d0 > d1 && d1 > d2);
            else res = (d0 < d1 && d1 < d2 && d2 < d3) || (d0 > d1 && d1 > d2 && d2 > d3);
            Console.WriteLine(res);
            // Must give in doc file a whole set of examples...
        }

        public String[] Create_Q6_Input(Object[] args)
        {
            int num_of_digits = (int)args[3];
            int minimalNum = (int)(Math.Pow(10, num_of_digits - 1));
            int maximalNum = (int)(Math.Pow(10, num_of_digits));
            String[] res = new String[1];
            res[0] = r.Next(minimalNum,maximalNum).ToString();
            return res;
        }

        public void Create_Q6_doc(Object[] args, Document wordDoc, int seif)
        {
            int num_of_digits = (int)(args[3]);

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) קולטת מספר בן {1} ספרות ותדפיס True אם הספרות מהוות רצף עולה או רצף יורד. אחרת (הספרות לא מהוות לא רצף עולה ולא רצף יורד) על התכנית להדפיס False.", seif, num_of_digits);
            par1.Range.InsertParagraphAfter();

            if (num_of_digits == 3)
            {
                par1.Range.Text = String.Format("לדוגמא, עבור הקלטים 123, 579, 750 ו-963 על התוכנית להדפיס True.");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "True";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                par1.Range.Text = String.Format("אבל, עבור הקלטים 122, 576, 755 ו-563 על התוכנית להדפיס False.");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "False";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }
            else // (num_of_digits == 3)
            {
                par1.Range.Text = String.Format("לדוגמא, עבור הקלטים 1234, 5789, 7530 ו-9763 על התוכנית להדפיס:");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "True";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                par1.Range.Text = String.Format("   אבל, עבור הקלטים 1224, 5786, 7550 ו-9563 על התוכנית להדפיס:");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                par1.Range.Text = "False";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
        }

        public void Create_Q7_doc(Object[] args, Document wordDoc, int seif)
        {
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.Text = String.Format("{0}) קולטת מספר שלם X באורך כלשהוא ומספר שלם נוסף k שמייצג מיקום של ספרה בתוך המספר X.", seif);
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format("התוכנית תדפיס את הספרה שנמצאת במקום ה-k במספר X כאשר הספירה מתחילה מצד ימין. כלומר, ספרת האחדות היא במקום 1, ספרת העשרות במקום 2 וכו'");
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = "לדוגמא אם הקלט X היה 15243 והקלט השני k הוא 2, אז הפלט צריך להיות:";
            par1.Range.InsertParagraphAfter();

            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = "Digit at place 2 is 4";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
        }

        public String[] Create_Q7_Input(Object[] args)
        {
            int num_of_digits = 6;
            int minimalNum = (int)(Math.Pow(10, num_of_digits - 1));
            int maximalNum = (int)(Math.Pow(10, num_of_digits + 2));
            String[] res = new String[2];
            int chosenNum = r.Next(minimalNum, maximalNum);
            res[0] = chosenNum.ToString();
            int chosenNumLength = res[0].Trim().Length;
            res[1] = r.Next(1, chosenNumLength + 1).ToString();
            return res;
        }

        public void Create_Q7(Object[] args, bool real_input)
        {
            Console.WriteLine("Please enter one integer number:");
            int[] options = { 91134, 26145, 987654, 12, 3210 };
            int num = int.Parse(get_input_integer(real_input, options[r.Next(0, options.Length)]));
            Console.WriteLine("Please enter digit location:");
            int location = int.Parse(get_input_integer(real_input, r.Next(1, num.ToString().Length + 1)));
            Console.WriteLine("Digit at place {0} is {1}", location, (num / (int)(Math.Pow(10, location - 1))) % 10);
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
                afterRandom.Add(new Creators(Create_Q1, Create_Q1_doc, Create_Q1_Input));

                List<Creators> creators = new List<Creators>();
                creators.Add(new Creators(Create_Q3, Create_Q3_doc, Create_Q3_Input));
                creators.Add(new Creators(Create_Q4, Create_Q4_doc, Create_Q4_Input));

                while (creators.Count > 0)
                {
                    int rndIdx = r.Next(0, creators.Count);
                    Debug.WriteLine("rndIdx=" + rndIdx);
                    afterRandom.Add(creators[rndIdx]);
                    creators.RemoveAt(rndIdx);
                }

                creators.Add(new Creators(Create_Q5, Create_Q5_doc, Create_Q5_Input));
                creators.Add(new Creators(Create_Q6, Create_Q6_doc, Create_Q6_Input));
                creators.Add(new Creators(Create_Q7, Create_Q7_doc, Create_Q7_Input));

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

        public override void createRandomInputFile(int id, String filePath)
        {
            String[] input_function_names = loadInputFunctions(id);
            Object[] args = LoadArgs(id);
            Type magicType = this.GetType();
            ConstructorInfo magicConstructor = magicType.GetConstructor(Type.EmptyTypes);
            object magicClassObject = magicConstructor.Invoke(new object[] { });

            using (StreamWriter sw = new StreamWriter(filePath, false))
            {
                for (int i = 0; i < input_function_names.Length; i++)
                {
                    // Get the ItsMagic method and invoke with a parameter value of 100
                    MethodInfo magicMethod = magicType.GetMethod(input_function_names[i]);
                    String[] resultingLines = (String[])(magicMethod.Invoke(magicClassObject, new object[] { args }));
                    for (int x = 0; x < resultingLines.Length; x++)
                    {
                        sw.WriteLine(resultingLines[x]);
                    }
                }
                sw.WriteLine(); // for all the Console.ReadLine() students may add at the end
            }

        }

        public String[] loadOrderFunctions(int id)
        {
            String questionsOrderFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_questions_order.txt";
            return File.ReadAllLines(questionsOrderFilePath);
        }

        public String[] loadInputFunctions(int id)
        {
            String inputsOrderFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_inputs_order.txt";
            return File.ReadAllLines(inputsOrderFilePath);
        }

        public String[] saveOrderFunctions(int id, List<Creators> list)
        {
            String[] res = new String[list.Count];
            String questionsOrderFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_questions_order.txt";
            using (StreamWriter sw = new StreamWriter(questionsOrderFilePath, false))
            {
                for (int i = 0; i < list.Count; i++)
                {
                    sw.WriteLine(list[i].questFunc.Method.Name);
                    res[i] = list[i].questFunc.Method.Name;
                }
            }
            String inputsOrderFilePath = Students_Hws_dirs + "\\" + id.ToString() + "_inputs_order.txt";
            using (StreamWriter sw = new StreamWriter(inputsOrderFilePath, false))
            {
                for (int i = 0; i < list.Count; i++)
                {
                    sw.WriteLine(list[i].inputFunc.Method.Name);
                }
            }
            return res;
        }
        public virtual void Create_DocFile_By_Creators(Object[] args, List<Creators> creators)
        {
            int id = (int)(args[0]);

            String student_full_name = Students.students_dic[id].first_name + " " + Students.students_dic[id].last_name;


            //            String orig_file_path = pattern_dir + "//" + pattern_file_orig;
            //ADDING A NEW DOCUMENT TO THE APPLICATION
            Application oWord = new Application();
            oWord.Visible = true;
            Document wordDoc = oWord.Documents.Add();

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "שלום " + student_full_name;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "ש\"ב 1 נועדו לתרגל אתכם על שימוש במשתנים ובאופרטורים כפי שנלמדו בהרצאה ובתרגול. כרגיל, עליכם לייצר בדיוק את הפלט המצופה כדי שהבודק האוטומטי לא יכשיל אתכם. ושוב כרגיל, דוגמא לפלט המצופה מופיעה בסוף המסמך הזה.";
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

            par1.Range.Text = String.Format("תאריך הגשה אחרון - 30/11/2016 בשעה 23:55");
            par1.Range.InsertParagraphAfter();
            par1.Range.Underline = WdUnderline.wdUnderlineNone;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עליך ליכתוב תוכנית אשר:";
            par1.Range.InsertParagraphAfter();

            for (int i = 0; i < 3; i++)
            {
                creators[i].docFunc(args, wordDoc, (i + 1));
            }

            par1.Range.InsertBreak(WdBreakType.wdPageBreak);

            for (int i = 3; i < creators.Count; i++)
            {
                creators[i].docFunc(args, wordDoc, (i + 1));
            }
/*
            Create_Q5_doc(args, wordDoc, 4);
            Create_Q6_doc(args, wordDoc, 5);
            Create_Q7_doc(args, wordDoc, 6);
*/
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

        public override void Create_DocFile(Object[] args)
        {
            int id = (int)args[0];
            int shape = (int)args[1];
            int shape_size = (int)args[2];
            int kelet_repetitions = (int)args[3];
            int shave_reps = (int)args[4];

            String student_full_name = Students.students_dic[id].first_name + " " + Students.students_dic[id].last_name;


            //            String orig_file_path = pattern_dir + "//" + pattern_file_orig;
            //ADDING A NEW DOCUMENT TO THE APPLICATION
            Application oWord = new Application();
            oWord.Visible = true;
            Document wordDoc = oWord.Documents.Add();

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "שלום " + student_full_name;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "ש\"ב 1 נועדו לתרגל אתכם על שימוש במשתנים ובאופרטורים כפי שנלמדו בהרצאה ובתרגול. כרגיל, עליכם לייצר בדיוק את הפלט המצופה כדי שהבודק האוטומטי לא יכשיל אתכם. ושוב כרגיל, דוגמא לפלט המצופה מופיעה בסוף המסמך הזה.";
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
            par1.Range.Text = "תאריך הגשה אחרון - 25/11/2016 בשעה 23:55";
            par1.Range.InsertParagraphAfter();
            par1.Range.Underline = WdUnderline.wdUnderlineNone;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עליך ליכתוב תוכנית אשר:";
            par1.Range.InsertParagraphAfter();

            Create_Q1_doc(args, wordDoc, 1);
            Create_Q3_doc(args, wordDoc, 2);
            Create_Q4_doc(args, wordDoc, 3);
            par1.Range.InsertBreak(WdBreakType.wdPageBreak);

            Create_Q5_doc(args, wordDoc, 4);
            Create_Q6_doc(args, wordDoc, 5);
            Create_Q7_doc(args, wordDoc, 6);

            par1.Range.Text = "בעמוד הבא מופיעה דוגמא לפלט המצופה מהתוכנית שלך. זיכרו כי בדוגמא זו, השורות הכחולות מציינות שורות קלט מה-Console שהוכנסו על ידי מריץ התוכנית. השורות הלבנות מסמנות שורות פלט שנכתבו על ידי התוכנית אל ה-Console.";
            par1.Range.InsertParagraphAfter();

            par1.Range.InsertBreak(WdBreakType.wdPageBreak);
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            wordDoc.Application.Selection.Collapse();
            InlineShape bigExamplePicture = Worder.Replace_to_picture(wordDoc, "XXXX", Students_Hws_dirs + "\\" + id.ToString() + ".png");

            /*
                        wordDoc.Save();
                        wordDoc.Close();
                        oWord.Quit();
                        String studentDocFilePath = Students_Hws_dirs + "\\" + id.ToString() + ".docx";
                        System.Threading.Thread.Sleep(500);
                        if (File.Exists(studentDocFilePath)) File.Delete(studentDocFilePath);
                        File.Move(orig_file_path, studentDocFilePath);
                        System.Threading.Thread.Sleep(300);
                        String copy_file_path = pattern_dir + "//" + pattern_file_copy;
                        File.Copy(copy_file_path, orig_file_path);
            */
            return;
        }


        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[4];
            args[0] = id;
            args[1] = r.Next(3, 5); // Q3
            args[2] = r.Next(0, 2); // Q4
            args[3] = r.Next(3, 5); // Q6
            return args;

        }
    }
}
