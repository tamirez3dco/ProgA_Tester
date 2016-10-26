using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentsLib;
using System.IO;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using System.Windows.Forms;

namespace HWs_Generator
{
    public class GUI1 : HW4
    {
        //[DllImport("user32")]

        [DllImport("gdi32.dll")]
        static extern uint GetBkColor(IntPtr hdc);
/*
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }

        /// <summary>
        /// Delegate for the EnumChildWindows method
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="parameter">Caller-defined variable; we use it for a pointer to our list</param>
        /// <returns>True to continue enumerating, false to bail.</returns>
        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);
*/

        public GUI1()
        {
            Students_All_Hws_dirs = @"D:\Tamir\Netanya_Desktop_App\2017\Students_HWs";
            Students_Hws_dirs = Students_All_Hws_dirs + @"\" + this.GetType().Name;
        }

    public override RunResults Test_HW(object[] args, string resulting_exe_path)
        {
            if (!File.Exists(resulting_exe_path))
            {
                throw new Exception(String.Format("resulting_exe_path={0} does not exist", resulting_exe_path));
            }
            FileInfo exe_fin = new FileInfo(resulting_exe_path);
            
            RunResults rr = new RunResults();
            p = new Process();
            p.StartInfo.FileName = resulting_exe_path;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.WorkingDirectory = exe_fin.Directory.FullName;
            p.ErrorDataReceived += P_ErrorDataReceived;
            p.OutputDataReceived += P_OutputDataReceived;
            p.EnableRaisingEvents = true;


            p.Start();
            String mainWindowTitle = p.MainWindowTitle;
            IntPtr mainWindowHandle = p.MainWindowHandle;
            while (true)
            {
                mainWindowTitle = p.MainWindowTitle;
                Thread.Sleep(1000);

                uint colorInt = GetBkColor(mainWindowHandle);
                Debug.WriteLine("colorInt={0}, title={1}", colorInt, mainWindowTitle);

            }
            //List<IntPtr> childrenMainWindows = GetChildWindows(mainWindowHandle);
            //List<IntPtr> controls = new List<IntPtr>();
            //EnumWindow(mainWindowHandle, controls);

            return rr;

        }


        public List<Control> ScreenControlsByType(Control.ControlCollection ctrls, Type x)
        {
            List<Control> res = new List<Control>();
            foreach (Control c in ctrls)
            {
                if (c.GetType() == x) res.Add(c);
            }
            return res;
        }

        public Control ScreenControlsByText(Control.ControlCollection ctrls, String x)
        {
            foreach (Control c in ctrls)
            {
                if (c.Text == null) continue;
                if (c.Text.ToLower() == x.ToLower()) return c;
            }
            return null;
        }

        public override RunResults test_Hw_by_assembly(object[] args, FileInfo executable)
        {
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


            ConstructorInfo form_empty_constructor = son_form.GetConstructor(new Type[0]);
            Form form_to_run = (Form)form_empty_constructor.Invoke(new Object[0]);

            form_to_run.Show();
            MySleep(10000);

            Debug.WriteLine("Form title=" + form_to_run.Text);
            Debug.WriteLine("First Back = " + form_to_run.BackColor.ToString());

            Label label_to_look_at = null;
            List<Control> labels_in_form = ScreenControlsByType(form_to_run.Controls, typeof(Label));
            for (int i = 0; i < labels_in_form.Count; i++)
            {
                Label l = (Label)labels_in_form[i];
                if (!l.Visible) continue;
                if (l.Text.Trim() == "") continue;
                label_to_look_at = l;
            }
            if (label_to_look_at == null)
            {
                rr.grade = 50;
                rr.error_lines.Add("No Visible label found!!!");
                return rr;
            }
            Button b = (Button)ScreenControlsByText(form_to_run.Controls,"counter");
            if (b == null)
            {
                rr.grade = 50;
                rr.error_lines.Add("No button \"counter\" found!!!");
                return rr;
            }

            int counter;
            if (!int.TryParse(label_to_look_at.Text, out counter))
            {
                rr.grade = 50;
                rr.error_lines.Add("Could not parse label content!!!");
                return rr;
            }
            for (int i = 0; i < counter; i++)
            {
                b.PerformClick();
                MySleep(500);

                Debug.WriteLine("BackColor=" + form_to_run.BackColor.ToString());
                MySleep(500);

            }

            return rr;



        }

        private void MySleep(int millis)
        {
            for (int i = 0; i < millis; i++)
            {
                Debug.Write(String.Empty);
            }

        }
    }
}
