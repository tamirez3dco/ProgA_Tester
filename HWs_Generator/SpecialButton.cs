using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HWs_Generator
{
    class SpecialButton : Button
    {
        public SpecialButton(object[] args)
        {

        }
        public SpecialButton()
        {
            Gate = SIDE.LEFT;
        }


        private int gate_width;
        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            if (myDisable)
            {
                g.FillRectangle(new SolidBrush(Gate_color), ClientRectangle);
            }
            else
            {
                base.OnPaint(pevent);
            }
            if (Gate != SIDE.UP) g.DrawLine(new Pen(Gate_color, Gate_width), 0, Gate_width / 2, Width, Gate_width / 2);
            if (Gate != SIDE.DOWN) g.DrawLine(new Pen(Gate_color, Gate_width), 0, Height - Gate_width / 2, Width, Height - Gate_width / 2);
            if (Gate != SIDE.LEFT) g.DrawLine(new Pen(Gate_color, Gate_width), Gate_width / 2, 0, Gate_width / 2, Height);
            if (Gate != SIDE.RIGHT) g.DrawLine(new Pen(Gate_color, Gate_width), Width - Gate_width / 2, 0, Width - Gate_width / 2, Height);
        }

        public bool myDisable = false;
        bool entered = false;
        protected override void OnMouseMove(MouseEventArgs mevent)
        {
            base.OnMouseMove(mevent);
            if (!entered)
            {
                entered = true;
                SIDE enterance_direction = getClosestEdge(mevent.Location);
                if (enterance_direction != Gate)
                {
                    myDisable = true;
                    Refresh();
                }
            }
        }

        private SIDE getClosestEdge(Point location)
        {
            int distToLeft = location.X;
            int distToRight = Width - location.X;
            int distToTop = location.Y;
            int distToBottom = Height - location.Y;

            int min = Math.Min(Math.Min(distToLeft, distToRight), Math.Min(distToTop, distToBottom));
            if (min == distToLeft) return SIDE.LEFT;
            if (min == distToRight) return SIDE.RIGHT;
            if (min == distToBottom) return SIDE.DOWN;
            return SIDE.UP;
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            entered = false;
            myDisable = false;
            Refresh();
        }
        protected override void OnClick(EventArgs e)
        {
            if (!myDisable) base.OnClick(e);
        }

        private SIDE gate;

        public SIDE Gate
        {
            get
            {
                return gate;
            }

            set
            {
                gate = value;
            }
        }

        public int Gate_width
        {
            get
            {
                return gate_width;
            }

            set
            {
                gate_width = value;
            }
        }

        public Color Gate_color
        {
            get
            {
                return gate_color;
            }

            set
            {
                gate_color = value;
            }
        }

        private Color gate_color;

    }

    public enum SIDE
    {
        UP,
        DOWN,
        LEFT,
        RIGHT,
    }
}
