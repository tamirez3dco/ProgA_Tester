using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HWs_Generator
{
    public partial class MegaButton : UserControl
    {
        int time = 3;

        public  void clickFirstGateButton()
        {
            specialButton1.PerformClick();
        }
        public void clickThirdGateButton()
        {
            specialButton3.PerformClick();
        }
        public MegaButton()
        {
            InitializeComponent();
        }

        public MegaButton(Object[] args)
        {
            InitializeComponent();
            SpecialButton[] spButtons = { specialButton1,specialButton2,specialButton2,specialButton3 };
            foreach(SpecialButton sp in spButtons)
            {
                sp.Gate = (SIDE)args[(int)GUI3_ARGS.GATE_BUTTON_SIDE];
                sp.Gate_color = (Color)args[(int)GUI3_ARGS.GATE_DIS_COLOR];
                sp.Gate_width = (int)args[(int)GUI3_ARGS.GATE_PIX_WIDTH];
            }
            textBox1.Text = "Stam";
            adjustText();
            switch ((int)(args[(int)GUI3_ARGS.MEGA_PATTERN]))
            {
                case 0:
                    break;
                case 1:
                    tableLayoutPanel1.SetCellPosition(specialButton1, new TableLayoutPanelCellPosition(2, 1));
                    tableLayoutPanel1.SetCellPosition(specialButton2, new TableLayoutPanelCellPosition(0, 2));
                    tableLayoutPanel1.SetCellPosition(specialButton3, new TableLayoutPanelCellPosition(1, 3));
                    break;
                case 2:
                    tableLayoutPanel1.SetCellPosition(specialButton1, new TableLayoutPanelCellPosition(1, 1));
                    tableLayoutPanel1.SetCellPosition(specialButton2, new TableLayoutPanelCellPosition(0, 2));
                    tableLayoutPanel1.SetCellPosition(specialButton3, new TableLayoutPanelCellPosition(2, 3));
                    break;
                case 3:
                    tableLayoutPanel1.SetCellPosition(specialButton1, new TableLayoutPanelCellPosition(0, 1));
                    tableLayoutPanel1.SetCellPosition(specialButton2, new TableLayoutPanelCellPosition(2, 2));
                    tableLayoutPanel1.SetCellPosition(specialButton3, new TableLayoutPanelCellPosition(1, 3));
                    break;
            }
        }

        private void adjustText()
        {
            String text = textBox1.Text;
            specialButton1.Text = text + "1";
            specialButton2.Text = text + "2";
            specialButton3.Text = text + "3";
        }
        

        public int Time
        {
            get
            {
                return time;
            }

            set
            {
                time = value;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            adjustText();
        }

        private void specialButton1_Click(object sender, EventArgs e)
        {
            SpecialButton sp = (SpecialButton)sender;
            sp.BackColor = Color.Yellow;
            label1.Text = Time.ToString();
            label1.Visible = true;
        }
    }
}
