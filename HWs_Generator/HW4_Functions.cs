using Microsoft.Office.Interop.Word;
using Mono.Reflection;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HWs_Generator
{
    public class HW4 : HW2
    {
        public HW4()
        {
            Students_Hws_dirs = Students_All_Hws_dirs + @"\HW4";
            exampleRectangleSize = new Size(500, 900);
        }

        public enum HW4_ARGS
        {
            ID,
            Q1_FUNC_CIRCLE_AREA_OR_CIRCLE_PERIMITER,
            Q1_ROUNDING,
            Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE,
            Q3_ROUNDING,
            Q4_START,
            Q5_TYPE
        }
        public override Object[] get_random_args(int id)
        {
            Object[] args = new Object[Enum.GetNames(typeof(HW4_ARGS)).Length];
            args[(int)HW4_ARGS.ID] = id;
            args[(int)HW4_ARGS.Q1_FUNC_CIRCLE_AREA_OR_CIRCLE_PERIMITER] = r.Next(0, 2); // Q1 - funcName (Circle Area, Circle Perimiter)
            args[(int)HW4_ARGS.Q1_ROUNDING] = r.Next(3, 6); // Q1 - rounding Q1
            args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE] = r.Next(1, 2); // Q2 - func Name (Dist, Stairs Dist)
            args[(int)HW4_ARGS.Q3_ROUNDING] = r.Next(3, 6); // Q3 - Rounding Q3
            args[(int)HW4_ARGS.Q4_START] = r.Next(2, 5); // Q4 - starting point
            args[(int)HW4_ARGS.Q5_TYPE] = r.Next(0, 3); // Q5 - type (int, double, string)
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

            /*
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
            */
            Create_DocFile_By_Creators(args, afterRandom);
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

            par1.Range.Text = "ש\"ב 4 נועדו לתרגל אתכם על כתיבת פונקציות וקריאה לפונקציות כפי שנלמדו בהרצאה ובתרגול. כרגיל, על הפונקציות שלכם לעשות בדיוק את המצופה מהן כדי שהבודק האוטומטי לא יכשיל אתכם.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "בכל המקומות בתרגיל שמצוין מספר שלם - הכוונה היא לטיפוס int.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "בכל המקומות בתרגיל שמצוין מספר לא שלם - הכוונה היא לטיפוס double.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "זיכרו כי, בשלב הזה של הקורס, כל הפונקציות שתכתבו יהיו static.";
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "";
            par1.Range.InsertParagraphAfter();
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = "עליך ליכתוב את הפונקציות (השיטות) הבאות:";
            par1.Range.InsertParagraphAfter();

            Create_Q5_doc(args, wordDoc, 1);

            Create_Q1_doc(args, wordDoc, 2);

            Create_Q2_doc(args, wordDoc, 3);

            Create_Q3_doc(args, wordDoc, 4);

            Create_Q4_doc(args, wordDoc, 5);

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



        public virtual RunResults test_Hw_by_assembly(Object[] args, FileInfo executable)
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

            RunResults rr2 = test_Q2(args, t);
            RunResults rr3 = test_Q3(args, t);
            RunResults rr4 = test_Q4(args, t);
            RunResults rr5 = test_Q5(args, t);
            RunResults rr1 =  test_Q1(args, t);
            rr = rr1 + rr2 + rr3 + rr4 + rr5;
            return rr;
        }

        public override RunResults Test_HW(object[] args, string resulting_exe_path)
        {
            return test_Hw_by_assembly(args, new FileInfo(resulting_exe_path));
        }

        static MethodInfo method_to_run;
        static Object[] method_params_to_run;
        static Object return_value_method_to_run;

        static void ActualMethodWrapper()
        {
            Object res = null;
            try
            {
                return_value_method_to_run = method_to_run.Invoke(null, method_params_to_run);
            }
            catch (ThreadAbortException)
            {
                Console.WriteLine("Method aborted early");
            }
            catch (Exception e)
            {
                Console.WriteLine("Other exception:"+e.Message);
            }
        }



        static bool CallTimedOutMethod(int milliseconds)
        {
            ThreadStart ts = new ThreadStart(ActualMethodWrapper);
            Thread t = new Thread(ts);
            t.Start();
            int millisElapsed = 0;
            int millisStep = 300;
            do
            {
                Thread.Sleep(millisStep);
                if (!t.IsAlive) break;
                millisElapsed += millisStep;
            } while (millisElapsed <= milliseconds);
            //Thread.Sleep(milliseconds);
            if (t.IsAlive)
            {
                t.Abort();
                return false;
            }
            else
            {
                return true;
            }
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

            
            String method_name = "CalcCirclePerimeter";
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
            if (rp.ParameterType != typeof(Double))
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

            
            for (int test = 0; test < 20; test++)
            {
                Object[] test_params = new Object[1];
                Double test_radious = r.NextDouble() * 100;
                test_params[0] = test_radious;

                if (!VerifyMethodEndsOnTime(m1, test_params, 5000, rr))
                {
                    return rr;
                }


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

        private bool VerifyMethodEndsOnTime(MethodInfo m, object[] test_params, int millisTimeout, RunResults rr)
        {
            HW4.method_to_run = m;
            HW4.method_params_to_run = test_params;
            bool completed = CallTimedOutMethod(millisTimeout);
            if (!completed)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add("Method \"" + m.Name + "\" did not complete in 10 seconds !! for params:" + ". Minus " + grade_lost + " points");
                for (int i = 0; i < HW4.method_to_run.GetParameters().Count(); i++)
                {
                    rr.error_lines.Add(String.Format("ParamName=\"{0}\" , ParamValue=\"{1}\"", HW4.method_to_run.GetParameters().ElementAt(i).Name, test_params[i]));
                }
            }
            return completed;
        }

        MethodInfo q2_method = null;
        public RunResults test_Q3(Object[] args, Type t)
        {
            RunResults rr = new RunResults();

            int id = (int)args[0];
            int q2_func = (int)args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE];
            int qr_rounding = (int)args[(int)HW4_ARGS.Q3_ROUNDING];


            String method_name = "CalcPathDistance";
            if (q2_func == 1) method_name = "CalcPathStairsDistance";
            MethodInfo m2 = GetClosestMethod(t, method_name);
            if (m2 == null)
            {
                int grade_lost = 20;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("No method \"{0}\" found. Minus {1} points", method_name, grade_lost));
                return rr;
            }

            if (m2.Name != method_name)
            {
                int grade_lost = 2;
                if (m2.Name.ToLower() != method_name.ToLower()) grade_lost = 5;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" was expected to be called \"{1}\". Minus {2} points", m2.Name, method_name, grade_lost));
            }

            ParameterInfo[] method_params = m2.GetParameters();
            if (method_params.Length != 2)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong params list length. Minus {1} points", m2.Name, grade_lost));
                return rr;
            }

            for (int i = 0; i < method_params.Length; i++)
            {
                if (method_params[i].ParameterType != typeof(Double[]))
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong param type in parameter # {0}. Should be Double[] but instead its param is of type {1}. Minus {2} points", i, m2.Name, method_params[i].ParameterType.ToString(), grade_lost));
                    return rr;
                }
            }

            ParameterInfo rp = m2.ReturnParameter;
            if (rp.ParameterType != typeof(Double))
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong return type. Should be Double but instead its return type is of type {1}. Minus {2} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            if (!m2.IsStatic)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" should be static. Minus {1} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }


            for (int test = 0; test < 20; test++)
            {
                Object[] test_params = new Object[2];
                int path_length = r.Next(1, 50);
                if (test == 0) path_length = 1;
                Double[] Xs = new double[path_length];
                Double[] Ys = new double[path_length];
                for (int i = 0; i < path_length; i++)
                {
                    Xs[i] = r.NextDouble() * 100;
                    Ys[i] = r.NextDouble() * 100;
                }
                test_params[0] = Xs;
                test_params[1] = Ys;

                /*
                                HW4.method_to_run = m2;
                                HW4.method_params_to_run = test_params;
                                bool completed = CallTimedOutMethod(10000);
                                if (!completed)
                                {
                                    int grade_lost = 15;
                                    rr.grade -= grade_lost;
                                    rr.error_lines.Add("Method \"" + method_name + "\" did not complete in 10 seconds !! for params: Xs={" + StringfromDoubleArray(Xs) + "}, Ys={" + StringfromDoubleArray(Ys) + "}. Minus " + grade_lost + " points");
                                    return rr;
                                }
                */

                if (!VerifyMethodEndsOnTime(m2, test_params, 5000, rr))
                {
                    return rr;
                }

                Double retValue = 0;
                try
                {
                    retValue = (Double)m2.Invoke(null, test_params);
                }
                catch (TargetInvocationException e)
                {

                    int grade_lost = 10;
                    rr.grade -= grade_lost;

                    rr.error_lines.Add("Method \"" + method_name +"\" threw an unexpected exception !!! for params: Xs={" + StringfromDoubleArray(Xs) + "}, Ys={" + StringfromDoubleArray(Ys)  + "}. Minus " + grade_lost +" points");
                    rr.error_lines.Add(String.Format("Exception was : {0}",e.InnerException.Message));

                    return rr;
                }


                Double expected_ret_value = Q3(args, Xs,Ys);
                if (expected_ret_value != retValue)
                {
                    int grade_lost = 10;
                    rr.grade -= grade_lost;
                    
                    rr.error_lines.Add(String.Format("Method \"{0}\" returned wrong value for params: Xs={{1}}, Ys={{2}}. Expected value of {3} and actually returned value {4}. Minus {5} points",
                        m2.Name, StringfromDoubleArray(Xs), StringfromDoubleArray(Ys), expected_ret_value, retValue, grade_lost));
                    //return rr;
                }
            }

            if (q2_method == null)
            {
                int grade_lost = 5;
                rr.grade -= grade_lost;

                rr.error_lines.Add(String.Format("Method of Q2 is null - can not verify its usage in \"{0}\". Minus {1} points",
                    m2.Name, grade_lost));
                return rr;
            }

            // Get readable instructions.
            IList<Instruction> instructions = Disassembler.GetInstructions(m2);
            bool q2_method_found = false;
            for (int i = 0; i < instructions.Count; i++)
            {
                Instruction inst = instructions[i];
                if (inst.ToString().Contains(q2_method.Name))
                {
                    q2_method_found = true;
                    break;
                }
            }

            if (!q2_method_found)
            {
                int grade_lost = 10;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Could not locate usage of Q2 method=\"{0}\" inside Q3 method=\"{1}\". Minus {2} points",
                    q2_method.Name, m2.Name, grade_lost));
            }

            return rr;
        }

        /*
        stam1:i=0
min=-12,max=100
stam1:i=1
min=-12,max=100
stam1:i=2
min=21,max=100
stam1:i=3
min=21,max=99
stam1:i=4
min=21,max=77
stam1:i=5
min=43,max=43
stam1:i=6
stam1:i=7
stam1:i=8
stam1:i=9
stam1:i=10
        */
        public RunResults test_Q4(Object[] args, Type t)
        {
            RunResults rr = new RunResults();

            int id = (int)args[0];
            int step = (int)args[(int)HW4_ARGS.Q4_START];

            String method_name = "Main";

            MethodInfo m2 = GetClosestMethod(t, method_name);
            if (m2 == null)
            {
                int grade_lost = 20;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("No method \"{0}\" found. Minus {1} points", method_name, grade_lost));
                return rr;
            }

            if (m2.Name != method_name)
            {
                int grade_lost = 2;
                if (m2.Name.ToLower() != method_name.ToLower()) grade_lost = 5;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" was expected to be called \"{1}\". Minus {2} points", m2.Name, method_name, grade_lost));
            }

            ParameterInfo[] method_params = m2.GetParameters();
            if (method_params.Length != 1)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong params list length. Minus {1} points", m2.Name, grade_lost));
                return rr;
            }

            for (int i = 0; i < method_params.Length; i++)
            {
                if (method_params[i].ParameterType != typeof(string[]))
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong param type in parameter # {0}. Should be int but instead its param is of type {1}. Minus {2} points", i, m2.Name, method_params[i].ParameterType.ToString(), grade_lost));
                    return rr;
                }
            }

            ParameterInfo rp = m2.ReturnParameter;
            if (rp.ParameterType != typeof(void))
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong return type. Should be {1} but instead its return type is of type {2}. Minus {3} points", m2.Name, "void", rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            if (!m2.IsStatic)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" should be static. Minus {1} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }


            for (int test = 0; test < 1; test++)
            {
                Object[] test_params = new Object[1];
                test_params[0] = new string[0];

                HW4.method_to_run = m2;
                HW4.method_params_to_run = test_params;

                bool completed = CallTimedOutMethod(10000);
                if (!completed)
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add("Method \"" + method_name + "\" did not complete in 10 seconds !! Minus " + grade_lost + " points");
                    return rr;
                }


                TextWriter tw = Console.Out;
                String q5_fileResult = "q5_output.txt";
                using (StreamWriter sw = new StreamWriter(q5_fileResult))
                {
                    Console.SetOut(sw);

                    Array retValue = null;
                    try
                    {

                        retValue = (Array)m2.Invoke(null, test_params);
                    }
                    catch (TargetInvocationException e)
                    {

                        int grade_lost = 10;
                        rr.grade -= grade_lost;
                        rr.error_lines.Add("Method \"" + method_name + "\" threw an unexpected exception !!!. Minus " + grade_lost + " points");
                        rr.error_lines.Add(String.Format("Exception was : {0}", e.InnerException.Message));
                        Console.SetOut(tw);
                        return rr;
                    }
                    Console.SetOut(tw);
                }

                String[] expectedLines = File.ReadAllLines(@"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW4-expected.txt");
                String expectedAnswer = String.Empty;
                for (int i = step * 2 + 1; i < expectedLines.Length; i++)
                {
                    expectedAnswer += (expectedLines[i] + "\n");
                }

                String studentAnswer = File.ReadAllText(q5_fileResult);

                RunResults temp_rr = Test_Text(args, expectedAnswer, studentAnswer, "Q5", true);

                if (temp_rr.grade < 100)
                {
                    int actual_points_lost = Math.Min(15, temp_rr.Grade_Lost);
                    rr.grade -= actual_points_lost;
                    rr.error_lines.Add(String.Format("Your Q5 was not correct. System will not tell you exactly where you got it wrong. Should have lost {0} points. Deducting only {1} points",temp_rr.Grade_Lost, actual_points_lost));
                    rr.error_lines.Add(String.Format("However, system can advise that you had {0} missing lines, {1} minor lines diffs, {2} major line diffs, {3} extra non-blank lines and {4} extra blank//empty lines",
                        temp_rr.changes_counter[(int)TextDiffs.Missing], temp_rr.changes_counter[(int)TextDiffs.Modified_Minor], temp_rr.changes_counter[(int)TextDiffs.Modified_Major], temp_rr.changes_counter[(int)TextDiffs.Extra_line], temp_rr.changes_counter[(int)TextDiffs.Extra_blanks]+ temp_rr.changes_counter[(int)TextDiffs.Extra_Empty]));
                }
            }
            return rr;
        }


        public RunResults test_Q5(Object[] args, Type t)
        {
            RunResults rr = new RunResults();

            int id = (int)args[0];
            int type_num = (int)args[(int)HW4_ARGS.Q5_TYPE];

            Type[] types = { typeof(int), typeof(double), typeof(string) };

            String method_name = "Fill_Array";

            MethodInfo m2 = GetClosestMethod(t, method_name);
            if (m2 == null)
            {
                int grade_lost = 20;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("No method \"{0}\" found. Minus {1} points", method_name, grade_lost));
                return rr;
            }

            if (m2.Name != method_name)
            {
                int grade_lost = 2;
                if (m2.Name.ToLower() != method_name.ToLower()) grade_lost = 5;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" was expected to be called \"{1}\". Minus {2} points", m2.Name, method_name, grade_lost));
            }

            ParameterInfo[] method_params = m2.GetParameters();
            if (method_params.Length != 1)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong params list length. Minus {1} points", m2.Name, grade_lost));
                return rr;
            }

            for (int i = 0; i < method_params.Length; i++)
            {
                if (method_params[i].ParameterType != typeof(int))
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong param type in parameter # {0}. Should be int but instead its param is of type {1}. Minus {2} points", i, m2.Name, method_params[i].ParameterType.ToString(), grade_lost));
                    return rr;
                }
            }

            Type[] retType = { typeof(int[]), typeof(double[]), typeof(string[]) };
            string[] str_ret_types = { "int[]" , "double[]" , "string[]" };
            ParameterInfo rp = m2.ReturnParameter;
            if (rp.ParameterType != retType[type_num])
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong return type. Should be {1} but instead its return type is of type {2}. Minus {3} points", m2.Name, str_ret_types[type_num], rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            if (!m2.IsStatic)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" should be static. Minus {1} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }


            for (int test = 0; test < 1; test++)
            {
                Object[] test_params = new Object[1];
                int arr_length = r.Next(1, 50);
                test_params[0] = arr_length;

                Array arr = Array.CreateInstance(types[type_num], arr_length);
                for (int i = 0; i < arr_length; i++)
                {
                    switch (type_num)
                    {
                        case 0:
                            arr.SetValue(r.Next(0, 10000), i);
                            break;
                        case 1:
                            arr.SetValue(Math.Round(r.NextDouble()*10,5), i);
                            break;
                        case 2:
                            arr.SetValue(getRandomString(), i);
                            break;
                    }
                }

                String tempFileName = "temp.txt";
                using (StreamWriter sw = new StreamWriter(tempFileName))
                {
                    for (int i = 0; i < arr_length; i++) sw.WriteLine(arr.GetValue(i).ToString());
                }
                TextReader tr = Console.In;
                using (StreamReader sr = new StreamReader(tempFileName))
                {
                    Console.SetIn(sr);
                    HW4.method_to_run = m2;
                    HW4.method_params_to_run = test_params;
                    bool completed = CallTimedOutMethod(10000);
                    if (!completed)
                    {
                        int grade_lost = 15;
                        rr.grade -= grade_lost;
                        rr.filesToAttach.Add(tempFileName);
                        rr.error_lines.Add("Method \"" + method_name + "\" did not complete in 10 seconds !! The input lines used for this method are attached. Minus " + grade_lost + " points");
                        Console.SetIn(tr);
                        return rr;
                    }
                    Console.SetIn(tr);
                }

                Array retValue = null;
                using (StreamReader sr = new StreamReader(tempFileName))
                {
                    Console.SetIn(sr);

                    
                    try
                    {
                        retValue = (Array)m2.Invoke(null, test_params);
                    }
                    catch (TargetInvocationException e)
                    {

                        int grade_lost = 10;
                        rr.grade -= grade_lost;
                        rr.filesToAttach.Add(tempFileName);
                        rr.error_lines.Add("Method \"" + method_name + "\" threw an unexpected exception !!! Input lines attached at file " + tempFileName +". Minus " + grade_lost + " points");
                        rr.error_lines.Add(String.Format("Exception was : {0}", e.InnerException.Message));
                        Console.SetIn(tr);
                        return rr;
                    }
                    Console.SetIn(tr);
                }

                if (retValue.GetLength(0) != arr_length)
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.filesToAttach.Add(tempFileName);
                    rr.error_lines.Add(String.Format("Method \"{0}\" returned array of wrong size for input lines attached. Expected array size is {1} and actually array size received is {2}. Minus {3} points",
                        m2.Name, arr_length, arr.GetLength(0), grade_lost));
                    return rr;

                }
                for (int i = 0; i < arr_length; i++)
                {
                    bool comp_failed = false;
                    switch (type_num)
                    {
                        case 0:
                            int i1 = (int)arr.GetValue(i);
                            int i2 = (int)retValue.GetValue(i);
                            comp_failed = (i1 != i2);
                            break;
                        case 1:
                            double d1 = (double)arr.GetValue(i);
                            double d2 = (double)retValue.GetValue(i);
                            comp_failed = (d1 != d2);
                            break;
                        case 2:
                            string s1 = (string)arr.GetValue(i);
                            string s2 = (string)retValue.GetValue(i);
                            comp_failed = (s1 != s2);
                            break;
                    }
                    if (comp_failed)
                    {
                        int grade_lost = 10;
                        rr.grade -= grade_lost;
                        rr.filesToAttach.Add(tempFileName);
                        rr.error_lines.Add(String.Format("Method \"{0}\" returned wrong value for input lines attached. Expected value at index {1} is {2} and actually value at this index is {3}. Minus {4} points",
                            m2.Name, i, arr.GetValue(i).ToString(), retValue.GetValue(i).ToString(), grade_lost));
                        return rr;
                    }
                }
            }
            return rr;
        }


        public RunResults test_Q2(Object[] args, Type t)
        {
            RunResults rr = new RunResults();

            int id = (int)args[0];
            int q2_func = (int)args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE];


            String method_name = "CalcDistance";
            if (q2_func == 1) method_name = "CalcStairsDistance";
            MethodInfo m2 = GetClosestMethod(t, method_name);
            if (m2 == null)
            {
                int grade_lost = 20;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("No method \"{0}\" found. Minus {1} points", method_name, grade_lost));
                return rr;
            }

            if (m2.Name != method_name)
            {
                int grade_lost = 2;
                if (m2.Name.ToLower() != method_name.ToLower()) grade_lost = 5;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" was expected to be called \"{1}\". Minus {2} points", m2.Name, method_name, grade_lost));
            }

            q2_method = m2;

            ParameterInfo[] method_params = m2.GetParameters();
            if (method_params.Length != 4)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong params list length. Minus {1} points", m2.Name, grade_lost));
                return rr;
            }

            for (int i = 0; i < method_params.Length; i++)
            {
                if (method_params[i].ParameterType != typeof(Double))
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong param type in parameter # {0}. Should be Double but instead its param is of type {1}. Minus {2} points", i, m2.Name, method_params[i].ParameterType.ToString(), grade_lost));
                    return rr;
                }
            }

            String[] expected_param_names = { "x1", "y1", "x2", "y2" };
            List<String> vuzvuz = new List<string>();
            
            Dictionary<String, int> params_placer = new Dictionary<string, int>();
            for (int i = 0; i < method_params.Length; i++)
            {
                String param_name = method_params[i].Name;
                int idx = Array.IndexOf(expected_param_names, param_name.ToLower());

                if (idx < 0)
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Param # {0} in Method \"{1}\" has wrong name. Its name is {2} but should be one of x1,y1,x2,y2. Minus {2} points", 
                        i, m2.Name, param_name, grade_lost));
                    return rr;
                }
                else
                {
                    params_placer[expected_param_names[idx]] = i;
                }
            }
            if (params_placer.Count != 4)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Params in Method \"{1}\" has wrong names. Could not make which one is x1,y1,x2,y2. Minus {2} points",
                    m2.Name, grade_lost));
                return rr;
            }


            ParameterInfo rp = m2.ReturnParameter;
            if (rp.ParameterType != typeof(Double))
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" found has wrong return type. Should be Double but instead its return type is of type {1}. Minus {2} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }

            if (!m2.IsStatic)
            {
                int grade_lost = 15;
                rr.grade -= grade_lost;
                rr.error_lines.Add(String.Format("Method \"{0}\" should be static. Minus {1} points", m2.Name, rp.ParameterType.ToString(), grade_lost));
                return rr;
            }


            for (int test = 0; test < 20; test++)
            {
                Object[] test_params = new Object[4];
                Double test_x1 = r.NextDouble() * 100;
                Double test_y1 = r.NextDouble() * 100;
                Double test_x2 = r.NextDouble() * 100;
                Double test_y2 = r.NextDouble() * 100;
                test_params[params_placer["x1"]] = test_x1;
                test_params[params_placer["y1"]] = test_y1;
                test_params[params_placer["x2"]] = test_x2;
                test_params[params_placer["y2"]] = test_y2;

                if (!VerifyMethodEndsOnTime(m2, test_params, 5000, rr))
                {
                    return rr;
                }
/*
                HW4.method_to_run = m2;
                HW4.method_params_to_run = test_params;
                bool completed = CallTimedOutMethod(10000);
                if (!completed)
                {
                    int grade_lost = 15;
                    rr.grade -= grade_lost;
                    String tempErrorLine = "Method \"" + method_name + "\" did not complete in 10 seconds !! for params:";
                    for (int i = 0; i < HW4.method_to_run.GetParameters().Count(); i++)
                    {
                        tempErrorLine += String.Format("ParamName=\"{0}\"", HW4.method_to_run.GetParameters().ElementAt(i).Name);
                        tempErrorLine += String.Format("ParamValue=\"{0}\"", test_params[i]);
                    }
                    tempErrorLine += ". Minus " + grade_lost + " points";
                    rr.error_lines.Add(tempErrorLine);
                    return rr;
                }
*/

                Double retValue = (Double)m2.Invoke(null, test_params);

                Double expected_ret_value = Q2(args, test_x1, test_y1, test_x2, test_y2);
                if (expected_ret_value != retValue)
                {
                    int grade_lost = 10;
                    rr.grade -= grade_lost;
                    rr.error_lines.Add(String.Format("Method \"{0}\" returned wrong value for params: x1={1}, y1={2}, x2={3}, y2={4}. Expected value of {5} and actually returned value {6}. Minus {7} points",
                        m2.Name, test_x1,test_y1,test_x2,test_y2, expected_ret_value, retValue, grade_lost));
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
            int rounding = (int)args[(int)HW4_ARGS.Q1_ROUNDING];
            if ((int)args[(int)HW4_ARGS.Q1_FUNC_CIRCLE_AREA_OR_CIRCLE_PERIMITER] == 0) return Math.Round(Math.PI * 2 * radius, rounding);
            else return Math.Round(Math.PI * radius * radius, rounding);
        }
        public new void Create_Q1_doc(Object[] args, Document wordDoc, int seif)
        {
            int func_name = (int)args[(int)HW4_ARGS.Q1_FUNC_CIRCLE_AREA_OR_CIRCLE_PERIMITER];
            int rounding = (int)args[(int)HW4_ARGS.Q1_ROUNDING]; 
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

            if (func_name == 0)
            {
                par1.Range.Text = String.Format("להזכירכם, חישוב היקף של מעגל מתבצע ע\"י הנוסחה:");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                par1.Range.Text = "XXXX";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                wordDoc.Application.Selection.Collapse();
                InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\Circle_Perimiter.png");
                wordDoc.Application.Selection.Collapse();
            }
            else
            {
                par1.Range.Text = String.Format("להזכירכם, חישוב שטח של מעגל מתבצע ע\"י הנוסחה:");
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                par1.Range.Text = "XXXX";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                wordDoc.Application.Selection.Collapse();
                InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\Circle_Area.png");
                wordDoc.Application.Selection.Collapse();
            }
        }

        public new void Create_Q5_doc(Object[] args, Document wordDoc, int seif)
        {
            int type_num = (int)args[(int)HW4_ARGS.Q5_TYPE];

            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            String[] types = { "int", "double", "string" };

            par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם Fill_Array .שמקבלת כפרמטר מספר שלם שמציין אורך של מערך רצוי ומחזירה מערך של {1} בגודל המתאים אשר מולא בערכים מה-Console ע\"י מריץ התוכנית. הפונקציה לא אמורה להדפיס שום כלום למסך - רק לקרוא ל-()Console.ReadLine במספר הנכון של הפעמים ולהחזיר מערך מלא בערכים שהוקלדו. סדר האיברים במערך יהיה זהה לסדר האיברים שהוקלדו ב-Console.", seif, types[type_num]);
            par1.Range.InsertParagraphAfter();
        }

        public static double Q2(Object[] args, double x1, double y1, double x2, double y2)
        {

            double x_diff = x1 - x2;
            double y_diff = y1 - y2;
            if ((int)args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE] == 0)
            {
                return Math.Sqrt((x_diff) * (x_diff) + (y_diff) * (y_diff));
            }
            else return Math.Abs(x_diff) + Math.Abs(y_diff);
        }

        public new void Create_Q2_doc(Object[] args, Document wordDoc, int seif)
        {
            int func_name = (int)args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            if (func_name == 0)
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcDistance .שמקבלת כפרמטר 4 מספרים לא שלמים שמייצגים 2 נקודות במישור דו-מימדי. יש לקרוא לפראמטרים בשמות x1,y1 ו- x2,y2.הפונקציה תחשב את המרחק בין הנקודות (x1,y1) ל- (x2,y2).יש להשתמש בנוסחה הבאה לחישוב המרחק (הידועה גם בשמה \"חוק פיתגורס\"):", seif);
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
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcStairsDistance .שמקבלת כפרמטר 4 מספרים לא שלמים שמייצגים 2 נקודות במישור דו-מימדי. נקרא לפראמטרים x1,y1 ו- x2,y2.הפונקציה תחשב את \"מרחק המדרגות\" בין הנקודות (x1,y1) ל- (x2,y2).מרחק המדרגות הוא פשוט סכום המרחק בין הנקודות על ציר ה-x עם המרחק בין הנקודות על ציר y. כלומר, יש להשתמש בנוסחה הבאה לחישוב המרחק:", seif);
                par1.Range.InsertParagraphAfter();

                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                par1.Range.Text = "XXXX";
                par1.Range.InsertParagraphAfter();
                par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

                wordDoc.Application.Selection.Collapse();
                InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW3-stairs_distance.png");
            }
            Double example_x1 = Math.Round(r.NextDouble() * 10, 1);
            Double example_y1 = Math.Round(r.NextDouble() * 10, 1);
            Double example_x2 = Math.Round(r.NextDouble() * 10, 1);
            Double example_y2 = Math.Round(r.NextDouble() * 10, 1);
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
            int func_name = (int)args[(int)HW4_ARGS.Q2_FUNC_DISTANCE_OR_STAIRS_DISTANCE];
            int rounding = (int)args[(int)HW4_ARGS.Q3_ROUNDING];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.InsertBreak(WdBreakType.wdPageBreak);

            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;
            double[] Xs = { 1.5, 1.5, 3, 3, 4.5 };
            double[] Ys = { 1.5, 2, 2, 4, 5 };

            if (func_name == 0)
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcPathDistance .שמקבלת כפרמטר מסלול המורכב ממספר נקודות במישור הדו-מימדי. באופן מדויק, הפונקציה תקבל כפראמטרים 2 מערכים של מספרים לא שלמים. המערך הראשון ייצג את ערכי הנקודות על ציר X, והמערך השני ייצג את ערכי הנקודות על ציר Y. הפונקציה תחשב את \"אורך המסלול\". אורך זה מחושב ע\"י סכום המרחקים בין נקודות עוקבות. על הפונקציה להשתמש בפונקציה CalcDistance שכתבנו בסעיף הקודם. את התשובה יש להחזיר בדיוק של עד {1} מקומות אחרי הנקודה העשרונית.", seif, rounding);
                par1.Range.InsertParagraphAfter();
            }
            else
            {
                par1.Range.Text = String.Format("{0}) כיתבו שיטה בשם CalcPathStairsDistance .שמקבלת כפרמטר מסלול המורכב ממספר נקודות במישור הדו-מימדי. באופן מדויק, הפונקציה תקבל כפראמטרים 2 מערכים של מספרים לא שלמים. המערך הראשון ייצג את ערכי הנקודות על ציר X, והמערך השני ייצג את ערכי הנקודות על ציר Y. הפונקציה תחשב את \"אורך מסלול מרחק המדרגות\". אורך זה מחושב ע\"י סכום מרחקי המדרגות בין נקודות עוקבות. על הפונקציה להשתמש בפונקציה CalcStairsDistance שכתבנו בסעיף הקודם. את התשובה יש להחזיר בדיוק של עד {1} מקומות אחרי הנקודה העשרונית.", seif, rounding);
                par1.Range.InsertParagraphAfter();
            }
            par1.Range.Text = String.Format("לדוגמא אם מערך ערכי ה-X שהפונקציה קיבלה כפראמטרים הוא:");
            par1.Range.InsertParagraphAfter();


            Range endDoc = wordDoc.Content;
            endDoc.Collapse(WdCollapseDirection.wdCollapseEnd);
            Table table = wordDoc.Tables.Add(endDoc, 2, 5);
            table.TableDirection = WdTableDirection.wdTableDirectionLtr;
            table.Spacing = 0;
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            for (int c = 0; c < 5; c++)
            {
                table.Cell(1, c + 1).Range.Text = "X" + c.ToString();
                table.Cell(1, c + 1).Range.Characters[2].Font.Subscript = 1;
                table.Cell(2, c + 1).Range.Text = Xs[c].ToString();
            }
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 40;

            par1.Range.Text = String.Format("ואם מערך ערכי ה-Y שהפונקציה קיבלה כפראמטרים הוא:");
            par1.Range.InsertParagraphAfter();

            endDoc = wordDoc.Content;
            endDoc.Collapse(WdCollapseDirection.wdCollapseEnd);
            Table tableY = wordDoc.Tables.Add(endDoc, 2, 5);
            tableY.TableDirection = WdTableDirection.wdTableDirectionLtr;
            tableY.Spacing = 0;
            tableY.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            tableY.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            Thread.Sleep(1000);
            tableY.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            Thread.Sleep(1000);
            tableY.PreferredWidth = 40;
            Thread.Sleep(1000);
            if (func_name == 1)
            {
                Thread.Sleep(1000);
                Ys[4] = 5.1234567;
                Thread.Sleep(1000);
                tableY.Columns[5].SetWidth(60, WdRulerStyle.wdAdjustProportional);
                Thread.Sleep(1000);
            }
            Thread.Sleep(1000);

            for (int c = 0; c < 5; c++)
            {
                tableY.Cell(1, c + 1).Range.Text = "Y" + c.ToString();
                tableY.Cell(1, c + 1).Range.Characters[2].Font.Subscript = 1;
                tableY.Cell(2, c + 1).Range.Text = Ys[c].ToString();
            }

            par1.Range.Text = String.Format("אז על הפונקציה להחזיר את הערך {0}. כיון ש:", Q3(args, Xs, Ys));
            par1.Range.InsertParagraphAfter();

            float spaceAfter = par1.Format.SpaceAfter;
            WdLineSpacing lineSpacing_rule = par1.Format.LineSpacingRule;
            float lineSpacing = par1.Format.LineSpacing;

            par1.Format.SpaceAfter = 0f;
            par1.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            par1.Format.LineSpacing = 11f;

            for (int i = 1; i < Xs.Length; i++)
            {
                if (func_name == 0)
                {
                    par1.Range.Text = String.Format("המרחק בין הנקודה (X{0},Y{0}) לנקודה (X{1},Y{1}) הוא {2}", i - 1, i, Q2(args, Xs[i - 1], Ys[i - 1], Xs[i], Ys[i]));
                    par1.Range.Characters[20].Font.Subscript = 1;
                    par1.Range.Characters[23].Font.Subscript = 1;
                    par1.Range.Characters[35].Font.Subscript = 1;
                    par1.Range.Characters[38].Font.Subscript = 1;
                }
                else
                {
                    par1.Range.Text = String.Format("\"מרחק המדרגות\" בין הנקודה (X{0},Y{0}) לנקודה (X{1},Y{1}) הוא {2}", i - 1, i, Q2(args, Xs[i - 1], Ys[i - 1], Xs[i], Ys[i]));
                    par1.Range.Characters[29].Font.Subscript = 1;
                    par1.Range.Characters[32].Font.Subscript = 1;
                    par1.Range.Characters[44].Font.Subscript = 1;
                    par1.Range.Characters[47].Font.Subscript = 1;
                }

                par1.Range.InsertParagraphAfter();
            }

            par1.Format.SpaceAfter = spaceAfter;
            par1.Format.LineSpacingRule = lineSpacing_rule;
            par1.Format.LineSpacing = lineSpacing;

            par1.Range.Text = String.Format("וסה\"כ המרחקים הללו יוצא    {0} (מעוגל ל-{1} מקומות אחרי הנקודה העשרונית)", Q3(args, Xs, Ys), rounding);
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("שימו לב : ניתן להניח כי המערכים שמתקבלים כפראמטרים לפונקציה הם באורך זהה.");
            par1.Range.InsertParagraphAfter();
            par1.Range.Text = String.Format(" שימו לב : עלולים להתקבל מערכים בגודל 1 - במקרה כזה על התוכנית להחזיר אורך מסלול 0.");
            par1.Range.InsertParagraphAfter();
        }

        public double Q3(Object[] args, double[] Xs, double[] Ys)
        {
            int rounding = (int)args[(int)HW4_ARGS.Q3_ROUNDING];
            double res = 0;
            for (int i = 1; i < Xs.Length; i++)
            {
                res += Q2(args, Xs[i - 1], Ys[i - 1], Xs[i], Ys[i]);
            }
            return Math.Round(res, rounding);
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
            int startPoint = (int)args[(int)HW4_ARGS.Q4_START];
            Paragraph par1 = wordDoc.Paragraphs.Add();
            par1.Range.InsertBreak(WdBreakType.wdPageBreak);
            par1.Range.Font.Name = "Ariel";
            par1.Range.Font.Size = 12;
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            par1.ReadingOrder = WdReadingOrder.wdReadingOrderRtl;

            par1.Range.Text = String.Format("{0}) בשאלה זו תצטרכו לנתח ריצה של תוכנית על סמך הקוד שלה ועל סמך מצב מחסנית הקריאות שמוצגים בהמשך. להלן קוד התוכנית לניתוח:", seif);
            par1.Range.InsertParagraphAfter();


            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            par1.Range.Text = "XXXX";
            par1.Range.InsertParagraphAfter();
            par1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            wordDoc.Application.Selection.Collapse();
            InlineShape shape = Worder.Replace_to_picture(wordDoc, "XXXX", @"D:\Tamir\Netanya_ProgrammingA\2017\Patterns_docs\HW4-program.png");
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoCTrue;
            shape.Height = 300;

            par1.Range.Text = String.Format("(החץ הצהוב מסמן את הפקודה הבאה שתבוצע -לאלה מכם שעוד לא עבדו עם ה-debugger של Visual Studio מומלץ מאד להתחיל). כלומר, אנחנו נמצאים כבר באמצע הריצה.");
            par1.Range.InsertParagraphAfter();
            
            par1.Range.Text = String.Format("בנוסף, להלן מצב מחסנית הקריאות לפני ביצוע השורה הצהובה:");
            par1.Range.InsertParagraphAfter();

            Range endDoc = wordDoc.Content;
            endDoc.Collapse(WdCollapseDirection.wdCollapseEnd);
            Table table = wordDoc.Tables.Add(endDoc, 4, 3);
            table.TableDirection = WdTableDirection.wdTableDirectionLtr;
            table.Spacing = 0;
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Columns[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[1].SetWidth(75, WdRulerStyle.wdAdjustProportional);
            table.Columns[2].SetWidth(150, WdRulerStyle.wdAdjustProportional);
            table.Cell(1, 1).Range.Text = "שם הפונקציה";
            table.Cell(1, 2).Range.Text = "פראמטרים שהועברו לפונקציה";
            table.Cell(1, 3).Range.Text = "משתנים שהוגדרו בתוך הפונקציה";
            table.Rows[1].Range.Font.BoldBi = 1;
            table.Rows[2].Range.Font.Size = 9;
            table.Rows[2].Range.Font.SizeBi = 9;
            table.Rows[2].Range.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;

            table.Rows[3].Range.Font.Size = 9;
            table.Rows[3].Range.Font.SizeBi = 9;
            table.Rows[3].Range.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;

            table.Rows[4].Range.Font.Size = 9;
            table.Rows[4].Range.Font.SizeBi = 9;
            table.Rows[4].Range.ParagraphFormat.ReadingOrder = WdReadingOrder.wdReadingOrderLtr;

            table.Cell(4, 1).Range.Text = "Main";
            table.Cell(3, 1).Range.Text = "stam1";
            table.Cell(2, 1).Range.Text = "stam2";

            table.Cell(4, 2).Range.Text = "String[] args={}";
            table.Cell(3, 2).Range.Text = "String[] array=*";
            table.Cell(2, 2).Range.Text = String.Format("String[] array=* ,x={0}",startPoint);

            table.Cell(4, 3).Range.Text = "int[] myArray={99, 12, 35, 67, 77, 43, 21, 99, 100, -12, 55}";
            table.Cell(3, 3).Range.Text = String.Format("int i={0}",startPoint);
            table.Cell(2, 3).Range.Text = "int min=? ,max=? , i=4";

            float spaceAfter = par1.Format.SpaceAfter;
            WdLineSpacing lineSpacing_rule = par1.Format.LineSpacingRule;
            float lineSpacing = par1.Format.LineSpacing;
            float fontSize = par1.Range.Font.Size;

            par1.Format.SpaceAfter = 0f;
            par1.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            par1.Format.LineSpacing = 10f;
            par1.Range.Font.Size = par1.Range.Font.SizeBi = 9;


            par1.Range.Text = String.Format("*=אותו ערך כמו למערך myArray שבפונקציה Main");
            par1.Range.InsertParagraphAfter();

            par1.Range.Text = String.Format("?=לא רציתי להגיד לכם את הערך של אלו");
            par1.Range.InsertParagraphAfter();


            par1.Format.SpaceAfter = spaceAfter;
            par1.Format.LineSpacingRule = lineSpacing_rule;
            par1.Format.LineSpacing = lineSpacing;
            par1.Range.Font.SizeBi = par1.Range.Font.Size = fontSize;

            par1.Range.Text = String.Format("עליכם לגלות אילו שורות תדפיס התוכנית מהמצב הזה ועד לסיומה. כדי להגיש את הפתרון פשוט הורו לפונקציה Main של הפרויקט שאותו אתם מגישים להדפיס שורות אלו. כלומר, הפונקציה Main בפרויקט שלכם תכלול בסה\"כ מספר שורות של Console.WriteLine עם הטקסט שאתם חושבים שהתוכנית של שאלה 4 תדפיס ממצבה הנוכחי.");
            par1.Range.InsertParagraphAfter();
        }


    }
}
