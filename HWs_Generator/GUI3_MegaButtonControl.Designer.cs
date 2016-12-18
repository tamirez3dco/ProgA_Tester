namespace HWs_Generator
{
    partial class MegaButton
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.specialButton1 = new HWs_Generator.SpecialButton();
            this.specialButton2 = new HWs_Generator.SpecialButton();
            this.specialButton3 = new HWs_Generator.SpecialButton();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.textBox1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.specialButton1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.specialButton2, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.specialButton3, 2, 3);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(150, 150);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // textBox1
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.textBox1, 3);
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Location = new System.Drawing.Point(0, 0);
            this.textBox1.Margin = new System.Windows.Forms.Padding(0);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(150, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(3, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 44);
            this.label1.TabIndex = 4;
            this.label1.Text = "label1";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.Visible = false;
            // 
            // specialButton1
            // 
            this.specialButton1.BackColor = System.Drawing.SystemColors.Control;
            this.specialButton1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.specialButton1.Gate = HWs_Generator.SIDE.LEFT;
            this.specialButton1.Gate_color = System.Drawing.Color.Red;
            this.specialButton1.Gate_width = 4;
            this.specialButton1.Location = new System.Drawing.Point(0, 20);
            this.specialButton1.Margin = new System.Windows.Forms.Padding(0);
            this.specialButton1.Name = "specialButton1";
            this.specialButton1.Size = new System.Drawing.Size(49, 43);
            this.specialButton1.TabIndex = 1;
            this.specialButton1.Text = "specialButton1";
            this.specialButton1.UseVisualStyleBackColor = false;
            this.specialButton1.Click += new System.EventHandler(this.specialButton1_Click);
            // 
            // specialButton2
            // 
            this.specialButton2.BackColor = System.Drawing.SystemColors.Control;
            this.specialButton2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.specialButton2.Gate = HWs_Generator.SIDE.LEFT;
            this.specialButton2.Gate_color = System.Drawing.Color.Red;
            this.specialButton2.Gate_width = 4;
            this.specialButton2.Location = new System.Drawing.Point(49, 63);
            this.specialButton2.Margin = new System.Windows.Forms.Padding(0);
            this.specialButton2.Name = "specialButton2";
            this.specialButton2.Size = new System.Drawing.Size(50, 43);
            this.specialButton2.TabIndex = 2;
            this.specialButton2.Text = "specialButton2";
            this.specialButton2.UseVisualStyleBackColor = false;
            this.specialButton2.Click += new System.EventHandler(this.specialButton1_Click);
            // 
            // specialButton3
            // 
            this.specialButton3.BackColor = System.Drawing.SystemColors.Control;
            this.specialButton3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.specialButton3.Gate = HWs_Generator.SIDE.LEFT;
            this.specialButton3.Gate_color = System.Drawing.Color.Red;
            this.specialButton3.Gate_width = 4;
            this.specialButton3.Location = new System.Drawing.Point(99, 106);
            this.specialButton3.Margin = new System.Windows.Forms.Padding(0);
            this.specialButton3.Name = "specialButton3";
            this.specialButton3.Size = new System.Drawing.Size(51, 44);
            this.specialButton3.TabIndex = 3;
            this.specialButton3.Text = "specialButton3";
            this.specialButton3.UseVisualStyleBackColor = false;
            this.specialButton3.Click += new System.EventHandler(this.specialButton1_Click);
            // 
            // MegaButton
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkGray;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "MegaButton";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TextBox textBox1;
        private SpecialButton specialButton1;
        private SpecialButton specialButton2;
        private SpecialButton specialButton3;
        private System.Windows.Forms.Label label1;
    }
}
