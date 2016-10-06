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
    public class HW4 : HW2
    {
        public HW4()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW3";
            exampleRectangleSize = new Size(500, 900);
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[5];
            args[0] = id;
            args[1] = r.Next(0, 2); // Q1
            args[2] = r.Next(2, 6); // Q1
            args[3] = r.Next(1, 2); // Q3
            args[4] = r.Next(3, 7); // Q3
            return args;

        }


        public override void Create_HW(Object[] args, bool real_input)
        {
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
            oWord.Visible = true;
            Document wordDoc = oWord.Documents.Add();

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "שלום " + student_full_name;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "ש\"ב 3 נועדו לתרגל אתכם על כתיבת פונקציות וקריאה לפונקציות כפי שנלמדו בהרצאה ובתרגול. כרגיל, על הפונקציות שלכם לעשות בדיוק את המצופה מהן כדי שהבודק האוטומטי לא יכשיל אתכם.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "בכל המקומות בתרגיל שמצוין מספר שלם - הכוונה היא לטיפוס int.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "בכל המקומות בתרגיל שמצוין מספר לא שלם - הכוונה היא לטיפוס Double.";
            par1.Range.InsertParagraphAfter();

/*
            par1.Range.Text = "לפני כל סעיף אבקש להדפיס שורה של 10 כוכביות ומספר הסעיף. לדוגמא, לפני הביצוע של סעיף 3 יש להדפיס את השורה:";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            par1.Range.Text = "3**********";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.Range.Text = "אם לא ברור, ניתן להסתכל בדוגמת הפלט הנדרש בסוף המסמך.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Underline = WdUnderline.wdUnderlineSingle;
*/
            par1.Range.Text = String.Format("תאריך הגשה אחרון - 11/12/2016 בשעה 23:55");
            par1.Range.InsertParagraphAfter();
            par1.Range.Underline = WdUnderline.wdUnderlineNone;

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עליך ליכתוב את הפונקציות (השיטות) הבאות:";
            par1.Range.InsertParagraphAfter();

            Create_Q1_doc(args, wordDoc, 1);

            Create_Q2_doc(args, wordDoc, 2);
            /*
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
            */
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



        public RunResults test_Hw(Object[] args, FileInfo executable)
        {
            RunResults rr = new RunResults();
            Assembly studentApp = Assembly.LoadFile(executable.FullName);
            Type[] appTypes = studentApp.GetTypes();
            if (appTypes.Length < 1)
            {
                rr.grade = 30;
                rr.error_lines.Add("No classes in code");
                return rr;
            }
            if (appTypes.Length > 1)
            {
                rr.error_lines.Add("Too many classes in code");
            }
            MethodInfo main_method = studentApp.EntryPoint;
            if (main_method == null)
            {
                rr.grade = 30;
                rr.error_lines.Add("No entry point (Main method) available in code");
                return rr;
            }

            Type t = main_method.DeclaringType;

            RunResults rr1 =  test_Q1(args, t);
            return rr;
        }

        public MethodInfo GetClosestMethod(Type t, String expectedName)
        {

            List<MethodInfo> all_methods = new List<MethodInfo>(t.GetMethods(BindingFlags.Static | BindingFlags.NonPublic));
            all_methods.AddRange(t.GetMethods(BindingFlags.Static | BindingFlags.Public));

            int lev_dist_min = 999;
            MethodInfo min = null;

            foreach (MethodInfo m in all_methods)
            {
                int lev_dist = LevenshteinDistance.Compute(expectedName.ToLower(), m.Name.ToLower());
                if (lev_dist < lev_dist_min)
                {
                    lev_dist_min = lev_dist;
                    min = m;
                }
            }

            return min;
        }

        public RunResults test_Q1(Object[] args, Type t)
        {
            RunResults rr = new RunResults();

            int id = (int)args[0];
            int q1_func = (int)args[1];

            
            String method_name = "CalcCirclePerimiter";
            if (q1_func == 1) method_name = "CalcCircleArea";
            MethodInfo m1 = GetClosestMethod(t, method_name);
            if (m1 == null)
            {
                int grade_lost = 20;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("No method \"{0}\" found. Minus {1} points", method_name, grade_lost));
                return rr;
            }

            if (m1.Name != method_name)
            {
                int grade_lost = 2;
                if (m1.Name.ToLower() != method_name.ToLower()) grade_lost = 5;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" was expected to be called \"{1}\". Minus {2} points", m1.Name, method_name, grade_lost));
            }

            ParameterInfo[] method_params = m1.GetParameters();
            if (method_params.Length != 1)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong params list length. Minus {1} points", m1.Name, grade_lost));
                return rr;
            }

            ParameterInfo pi = method_params[0];
            if (pi.ParameterType != typeof(Double))
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong param type. Should be Double but instead its param is of type {1}. Minus {2} points", m1.Name, pi.ParameterType.ToString(), grade_lost));
                return rr;
            }

            ParameterInfo rp = m1.ReturnParameter;
            if (pi.ParameterType != typeof(Double))
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong return type. Should be Double but instead its return type is of type {1}. Minus {2} points", m1.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            if (!m1.IsStatic)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" should be static. Minus {1} points", m1.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            
            for (int test = 0; test < 50; test++)
            {
                Object[] test_params = new Object[1];
                Double test_radious = r.NextDouble() * 100;
                test_params[0] = test_radious;
                Double retValue = (Double)m1.Invoke(null, test_params);

                Double expected_ret_value = Q1(args,test_radious);
                if (expected_ret_value != retValue)
                {
                    int grade_lost = 10;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" returned wrong value for radius={1}. Expected value of {2} and actually returned value {3}. Minus {4} points", m1.Name, test_radious, expected_ret_value, retValue, grade_lost));
                    return rr;
                }
            }
            return rr;
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

        public static double Q1(Object[] args, double radius)
        {
            int rounding = (int)args[2];
            if ((int)args[1] == 0) return Math.Round(Math.PI * 2 * radius, rounding);
            else return Math.Round(Math.PI * radius * radius, rounding);
        }
        public new void Create_Q1_doc(Object[] args, Document wordDoc, int seif)
        {
            int func_name = (int)args[1];
            int rounding = (int)args[2]; 
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            if (func_name == 0)
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcCirclePerimeter .שמקבלת כפרמטר רדיוס של מעגל (מספר לא שלם) ומחזירה את היקפו של המעגל מעוגל עד ל{1} נקודות אחרי הספרה העשרונית. יש להשתמש בקבוע Math.PI כדי לקבל תוצאות מדויקות.:", seif,rounding);
                par1.Range.InsertParagraphAfter();
            }
            else
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcCircleArea .שמקבלת כפרמטר רדיוס של מעגל (מספר לא שלם) ומחזירה את שיטחו של המעגל מעוגל עד ל{1} נקודות אחרי הספרה העשרונית. יש להשתמש בקבוע Math.PI כדי לקבל תוצאות מדויקות.:", seif,rounding);
                par1.Range.InsertParagraphAfter();
            }
            Double example_radius = Math.Round(r.NextDouble() * 10, 5);
            par1.Range.Text = String.Format("לדוגמא, אם הרדיוס שהועבר כפראמטר לפונקציה הוא {0} אז על הפונקציה להחזיר את הערך {1}. ", example_radius, Q1(args, example_radius));
            par1.Range.InsertParagraphAfter();
        }


        public static double Q2(Object[] args, double x1, double y1, double x2, double y2)
        {
            int rounding = (int)args[4];
            if ((int)args[3] == 0) return Math.Round(Math.Sqrt((x1- x2)*(x1- x2)+(y1- y2)*(y1- y2)), rounding);
            else return Math.Round(Math.Abs(x1- x2) + Math.Abs(y1- y2), rounding);
        }

        public new void Create_Q2_doc(Object[] args, Document wordDoc, int seif)
        {
            int func_name = (int)args[1];
            int rounding = (int)args[2];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            if (func_name == 0)
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcDistance .שמקבלת כפרמטר 4 מספרים לא שלמים שמייצגים 2 נקודות במישור דו-מימדי. נקרא לפראמטרים x1,y1 ו- x2,y2.הפונקציה תחשב את המרחק בין הנקודות (x1,y1) ל- (x2,y2).מעוגל עד ל{1} נקודות אחרי הספרה העשרונית. יש להשתמש בנוסחה הבאה לחישוב המרחק (הידועה גם בשמה \"חוק פיתגורס\":", seif, rounding);
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                par1.Range.Text = "XXXX";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                wordDoc.Application.Selection.Collapse();
                InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW1-nusha.png");

            }
            else
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcStairsDistance .שמקבלת כפרמטר 4 מספרים לא שלמים שמייצגים 2 נקודות במישור דו-מימדי. נקרא לפראמטרים x1,y1 ו- x2,y2.הפונקציה תחשב את מרחק המדרגות בין הנקודות (x1,y1) ל- (x2,y2).מעוגל עד ל{1} נקודות אחרי הספרה העשרונית. מרחק המדרגות הוא פשוט סכום המרחק בין הנקודות על ציר ה-x עם המרחק בין הנקודות על ציר y. כלומר, יש להשתמש בנוסחה הבאה לחישוב המרחק:", seif, rounding);
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                par1.Range.Text = "XXXX";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                wordDoc.Application.Selection.Collapse();
                InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW3-stairs_distance.png");
            }
            Double example_x1 = Math.Round(r.NextDouble() * 10, 5);
            Double example_y1 = Math.Round(r.NextDouble() * 10, 5);
            Double example_x2 = Math.Round(r.NextDouble() * 10, 5);
            Double example_y2 = Math.Round(r.NextDouble() * 10, 5);
            par1.Range.Text = String.Format("לדוגמא, אם הנקודות שיועברו לפונקציה הן (x1={0}, y1={1}) ו-(x2={2}, y2={3}) אז על הפונקציה להחזיר את הערך {4}  ", example_x1, example_y1, example_x2, example_y2, Q2(args, example_x1, example_y1,example_x2,example_y2));
//            par1.Range.Text = String.Format("לדוגמא, אם הנקודות שיועברו לפונקציה הן (x1={0},y1={1}) ו-(x2={2},y2={3}) אז על הפונקציה להחזיר את הערך}. ", example_x1, example_y1, example_x2, example_y2, Q2(args, example_x1, example_x1, example_x1, example_x1));
            par1.Range.InsertParagraphAfter();
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

            par1.Range.Text = String.Format("     לאחר מכן התוכנית מדפיסה את כל המספרים (מהמספר הקטן עד למספר הגדול) שמתחלקים ב{0} או שספרת האחדות שלהם או שספרת העשרות שלהם היא {0}", digiter);
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
                if ((d1 == digiter) || (d2 == digiter) || (i % digiter == 0))
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
            String[] inputs = { "9", "12.0", "-34", "45.00", "101", "-137.0", "-66.0", "32.0", "126.5" };
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
