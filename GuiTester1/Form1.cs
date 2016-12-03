using HWs_Generator;
using StudentsLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GuiTester1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //timer1.Start();
            //Cursor.Position = new Point();
            DoMouseClick(int.Parse(textBox1.Text), int.Parse(textBox2.Text));
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Cursor.Position = new Point(Cursor.Position.X +5,Cursor.Position.Y);
            //DoMouseClick();
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;


        public void DoMouseClick(int X, int Y)
        {
            Cursor.Position = new Point(X, Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)X, (uint)Y, 0, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new Students(@"D:\Tamir\Netanya_Desktop_App\2017\Shana_B_2017.xlsx");
            GUI2 hww = new GUI2();
            int tid = 029046117;
            //String resulting_exe_path;
            //Compiler.BuildZippedProject(@"D:\Tamir\Netanya_Desktop_App\2017\Students_Submissions\GUI1\312441710\13_11_2016_14_37.zip", out resulting_exe_path);
            //Object[] thw_args = hww.get_random_args(tid);
            Object[] thw_args = hww.LoadArgs(tid);
            RunResults rr = hww.test_Hw_by_assembly(thw_args, new FileInfo(@"D:\Tamir\Netanya_Desktop_App\2017\My_Solutions\GUI2_Mine\GUI2_Mine\bin\Debug\GUI2_Mine.exe"));
            MessageBox.Show(rr.ToString());

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Control c = GUI2.myPb;
            EventInfo evClick = c.GetType().GetEvent("MouseDown");
            FieldInfo eventClick = typeof(Control).GetField("EventMouseDown", BindingFlags.NonPublic | BindingFlags.Static);
            object secret = eventClick.GetValue(null);
            // Retrieve the click event
            PropertyInfo eventsProp = typeof(Component).GetProperty("Events", BindingFlags.NonPublic | BindingFlags.Instance);
            EventHandlerList events = (EventHandlerList)eventsProp.GetValue(c, null);
            //if (events.)
            Delegate click = events[secret];
            if (click == null) return;
            MethodInfo click_method = click.GetMethodInfo();
            //click.Method.Invoke(form_to_run,)
            ParameterInfo[] click_params = click_method.GetParameters();

            MouseEventArgs ea = new MouseEventArgs(MouseButtons.Left, 1, 1, 1, 0);
            Object[] click_objects = { c, ea };
            //MessageBox.Show("1");
            click_method.Invoke(GUI2.myForm, click_objects);

            for (int i = 0; i < 20000; i++)
            {
                Debug.Write(" ");
            }
            Debug.WriteLine("");

            //                MessageBox.Show("2");

        }
    }
}
