namespace HWs_Generator
{
    partial class GUI3_CardForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.cardControl1 = new HWs_Generator.CardControl();
            this.SuspendLayout();
            // 
            // cardControl1
            // 
            this.cardControl1.Card_suit = HWs_Generator.CardControl.SUIT.CLUB;
            this.cardControl1.Card_value = 7;
            this.cardControl1.Location = new System.Drawing.Point(69, 91);
            this.cardControl1.Name = "cardControl1";
            this.cardControl1.Size = new System.Drawing.Size(150, 150);
            this.cardControl1.TabIndex = 0;
            // 
            // GUI3_CardForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.cardControl1);
            this.Name = "GUI3_CardForm";
            this.Text = "GUI3_CardForm";
            this.ResumeLayout(false);

        }

        #endregion

        private CardControl cardControl1;
    }
}