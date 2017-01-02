﻿using StudentsLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HWs_Generator
{
    public partial class GUI3_GateButton_Comparer : Form
    {
        //This is a replacement for Cursor.Position in WinForms
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

        public GUI3_GateButton_Comparer()
        {
            InitializeComponent();
        }

        RunResults rr;
        object[] args;
        Assembly studAssembly, benchAssembly;
        public GUI3_GateButton_Comparer(Assembly studAssembly, Assembly benchAssembly, object[] args, RunResults rr)
        {
            this.studAssembly = studAssembly;
            this.benchAssembly = benchAssembly;
            this.args = args;
            this.rr = rr;
            InitializeComponent();
        }

        private Type getClosestTypeByNameProximity(Assembly asm, String expectedName)
        {
            Type[] allTypes = asm.GetTypes();
            String[] allTypeNames = new string[allTypes.Length];
            for (int i = 0; i < allTypes.Length; i++)
            {
                allTypeNames[i] = allTypes[i].Name;
            }
            int idx = LevenshteinDistance.GetIndexOfClosest(allTypeNames, expectedName);
            return allTypes[idx];
        }

        //private void 
        public void sendCursor(Point from, Point to)
        {
            lastMovePoints = new List<PointF>();
            diff = new Size(-1 * this.PointToClient(this.Location).X, -1 * this.PointToClient(this.Location).Y);
            lastMovePoints.Add(from + diff);
            guiResults.Add(new GuiResults());
            guiResults.Last().destination = end;
            guiResults.Last().moveState = MoveState.INIT;
            current = start = from;
            end = to;
            dir = new SizeF( ((float)end.X - from.X) / 10, ((float)end.Y - from.Y) / 10);
            Console.WriteLine("1)dir=" + dir);

            Cursor.Position = PointToScreen(from);
            timer1.Start();
        }

        private static double getDistance(PointF p1, PointF p2)
        {
            return Math.Abs(p1.X - p2.X) + Math.Abs(p1.Y - p2.Y);
        }
        Control studControl, benchControl;

        enum MoveState
        {
            INIT,
            STARTED,
            FINISHED
        }
        class GuiResults
        {
            public Point destination;
            public Bitmap formAtDestination;
            public bool clicked;
            public String remarks;
            public MoveState moveState;
        }
        Point start;
        PointF current;
        Point end;
        SizeF dir;
        int pointsCounter = 0;
        List<GuiResults> guiResults;
        List<PointF> lastMovePoints;
        int stop = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (stop++ > 100)
            {
                timer1.Stop();
                return;
            }
            current += dir;
            Cursor.Position = PointToScreen(new Point((int)current.X,(int)current.Y));
            //lastMovePoints.Add(new PointF(current.X + ClientRectangle.X, current.Y + ClientRectangle.Y));
            lastMovePoints.Add(current + diff);
            double dist = getDistance(current, end);
            Console.WriteLine("currect=" + current +", end=" + end + ", dist="+dist);
            if (dist < 4)
            {
                timer1.Stop();
                GuiResults gr = guiResults.Last();
                gr.destination = Cursor.Position;
                gr.moveState = MoveState.FINISHED;
                gr.clicked = false;
                Bitmap bmp = new Bitmap(Width,Height);
                this.DrawToBitmap(bmp,new Rectangle(0,0,Width,Height));
                Graphics g = Graphics.FromImage(bmp);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                Pen p = new Pen(Color.Red);
                
                AdjustableArrowCap bigArrow = new AdjustableArrowCap(3, 3);
                p.CustomEndCap = bigArrow;
                p.Width = 2;
                g.DrawLines(p, lastMovePoints.ToArray());
                gr.formAtDestination = bmp;
                bmp.Save(pointsCounter + ".png", ImageFormat.Png);
                pointsCounter++;
                clickTimer.Start();
                LeftMouseClick(Cursor.Position.X,Cursor.Position.Y);
                return;
            }

            if (dist < 20 && (Math.Abs(dir.Width) + Math.Abs(dir.Height) > 7))
            {
                float dx = ((float)end.X - current.X)/ 5;
                float dy = ((float)end.Y - current.Y)/ 5;
                dir = new SizeF(dx, dy);
                Console.WriteLine("dir=" + dir);
            }
            else if (dist > 1000)
            {
                timer1.Stop();
                MessageBox.Show("BALAGAN");
            }
        }

        private List<Rectangle> CreateInsideRectangles()
        {
            List<Rectangle> res = new List<Rectangle>();
            int x = 20;
            int y = 10;
            Rectangle insideTop = new Rectangle(studControl.Left + x, studControl.Top + y, 
                                              studControl.Width - 2 * x, studControl.Height / 3);
            res.Add(insideTop);
            Rectangle insideBottom = new Rectangle(studControl.Left + x, studControl.Bottom - y - studControl.Height / 3
                , studControl.Width - 2 * x, studControl.Height / 3);
            res.Add(insideBottom);

            Rectangle insideLeft = new Rectangle(studControl.Left + x, studControl.Top + y
                , studControl.Width / 4, studControl.Height - 2 * y);
            res.Add(insideLeft);

            Rectangle insideRight = new Rectangle(studControl.Right - x - studControl.Width / 4, studControl.Top + y
                , studControl.Width / 4, studControl.Height - 2 * y);
            res.Add(insideRight);

            Rectangle insideMiddle = new Rectangle(studControl.Left + x + studControl.Width / 3, studControl.Bottom + studControl.Height / 3
                , studControl.Width / 3, studControl.Height / 3);
            res.Add(insideMiddle);

            return res;
        }

        private List<Rectangle> CreateOutRectangles()
        {
            List<Rectangle> res = new List<Rectangle>();
            int x = 20;
            int y = 10;
            Rectangle outsideTop = new Rectangle(studControl.Left + x, studControl.Top - 20
                , studControl.Width - 2 * x, 10);
            res.Add(outsideTop);
            Rectangle outsideBottom = new Rectangle(studControl.Left + x, studControl.Bottom + 10
                , studControl.Width - 2 * x, 10);
            res.Add(outsideBottom);

            Rectangle outsideLeft = new Rectangle(studControl.Left - 20, studControl.Top + y
                , 10, studControl.Height - 2 * y);
            res.Add(outsideLeft);

            Rectangle outsideRight = new Rectangle(studControl.Right + 10, studControl.Top + y
                , 10, studControl.Height - 2 * y);
            res.Add(outsideRight);

            return res;
        }

        Random r = new Random();
        Point chooseRandomPointIsideRect(Rectangle rect)
        {
            return new Point(r.Next(rect.Left, rect.Right), r.Next(rect.Top, rect.Bottom));
        }

        String expectedControlName = "GateButton";

        private void clickTimer_Tick(object sender, EventArgs e)
        {
            clickTimer.Stop();
            if (pointsCounter > 20) return;

            // set new destination
            Point from = end;
            start = end;

            Rectangle rect;
            Control childAt = GetChildAtPoint(end);
            Debug.WriteLine("childAt({0} is {1}", end, childAt);
            if (childAt != studControl) rect = insideRectangles[r.Next(0, insideRectangles.Count)];
            else rect = outsideRectangles[r.Next(0, outsideRectangles.Count)];
            double dist;
            Point to;
            do
            {
                dist = 99;
                to = chooseRandomPointIsideRect(rect);
                Segment s = new Segment(from, to);
                foreach (PointF kodkod in studPoints)
                    dist = Math.Min(dist, Segment.dist_Point_to_Segment(kodkod, s));
                Debug.WriteLine("to=" + to + ", dist=" + dist);
            } while (dist < 3);

            sendCursor(from, to);
        }

        List<Rectangle> insideRectangles;
        List<Rectangle> outsideRectangles;

        Size diff;
        private void button2_Click(object sender, EventArgs e)
        {
            PointF p1 = new PointF(0,1);
            PointF p2 = new PointF(1, 0);
            PointF p3 = new PointF(0, 0);
            double dist = Segment.dist_Point_to_Segment(p3, new Segment(p1, p2));
        }
        List<PointF> studPoints;
        private void button1_Click(object sender, EventArgs e)
        {
            guiResults = new List<GuiResults>();
            Type ctrlType = getClosestTypeByNameProximity(studAssembly, expectedControlName);
            ConstructorInfo emptyCons = ctrlType.GetConstructor(new Type[0]);
            studControl = (Control)emptyCons.Invoke(new Object[0]);
            studControl.Location = new Point(200, 150);
            studControl.Size = new Size(70, 70);
            studControl.Text = "AutoTest";
            studControl.Name = "StudButton";
            studControl.Click += StudControl_Click;
            studPoints = new List<PointF>();
            studPoints.Add(new PointF(studControl.Left, studControl.Top));
            studPoints.Add(new PointF(studControl.Left, studControl.Bottom));
            studPoints.Add(new PointF(studControl.Right, studControl.Top));
            studPoints.Add(new PointF(studControl.Right, studControl.Bottom));
            this.Controls.Add(studControl);

            Point from = new Point(1,1);
            insideRectangles = CreateInsideRectangles();
            outsideRectangles = CreateOutRectangles();
            Rectangle rect = insideRectangles[r.Next(0, insideRectangles.Count)];
            double dist;
            Point to;
            do
            {
                dist = 99;
                to = chooseRandomPointIsideRect(rect);
                Segment s = new Segment(from, to);
                foreach (PointF kodkod in studPoints)
                    dist = Math.Min(dist, Segment.dist_Point_to_Segment(kodkod, s));
                Debug.WriteLine("to=" + to + ", dist=" + dist);
            } while (dist < 3);
            
            sendCursor(from, to);

        }

        private void StudControl_Click(object sender, EventArgs e)
        {
            guiResults.Last().clicked = true;
        }
    }


}