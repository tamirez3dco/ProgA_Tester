using HWs_Generator;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace GuiTester1
{
    static class Program
    {
        public static Form form_to_run = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Assembly studentApp = Assembly.LoadFile(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI1_Mine\GUI1_Mine\bin\Debug\GUI1_Mine.exe");
            Type[] appTypes = studentApp.GetTypes();
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


            ConstructorInfo form_empty_constructor = son_form.GetConstructor(new Type[0]);
            form_to_run = (Form)form_empty_constructor.Invoke(new Object[0]);

            form_to_run.Show();
        }
    }
}
