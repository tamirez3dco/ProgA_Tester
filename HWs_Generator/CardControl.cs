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
    public partial class CardControl : UserControl
    {
        public CardControl()
        {
            InitializeComponent();
        }

        private SUIT card_suit;
        private int card_value;

        public int Card_value
        {
            get
            {
                return card_value;
            }

            set
            {
                if (value > 13 || value < 1) return;
                card_value = value;
                String str = card_value.ToString();
                switch (value)
                {
                    case 1:
                        str = "A";
                        break;
                    case 10:
                        str = "T";
                        break;
                    case 11:
                        str = "J";
                        break;
                    case 12:
                        str = "Q";
                        break;
                    case 13:
                        str = "K";
                        break;
                }
                label1.Text = label2.Text = label3.Text = str;
            }
        }

        public SUIT Card_suit
        {
            get
            {
                return card_suit;
            }

            set
            {
                card_suit = value;
                switch (card_suit)
                {
                    case SUIT.CLUB:
                        tableLayoutPanel1.BackgroundImage = HWs_Generator.Properties.Resources.club;
                        break;
                    case SUIT.HART:
                        tableLayoutPanel1.BackgroundImage = HWs_Generator.Properties.Resources.hart;
                        break;
                    case SUIT.DAIMOND:
                        tableLayoutPanel1.BackgroundImage = HWs_Generator.Properties.Resources.daimond;
                        break;
                    case SUIT.SPADE:
                        tableLayoutPanel1.BackgroundImage = HWs_Generator.Properties.Resources.spade;
                        break;
                }
            }
        }

        public enum SUIT{
            HART,
            CLUB,
            DAIMOND,
            SPADE
        }
    }
}
