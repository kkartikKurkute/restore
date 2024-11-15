namespace ExcelApp
{
    partial class Form4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form4));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.rjButton1 = new ExcelApp.RjButton();
            this.rjButton2 = new ExcelApp.RjButton();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(59, 23);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(294, 268);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // rjButton1
            // 
            this.rjButton1.BackColor = System.Drawing.Color.Silver;
            this.rjButton1.BorderColor = System.Drawing.Color.LightGray;
            this.rjButton1.BorderColour = System.Drawing.Color.LightGray;
            this.rjButton1.BorderRadius = 40;
            this.rjButton1.BorderSize = 0;
            this.rjButton1.FlatAppearance.BorderSize = 0;
            this.rjButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.rjButton1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rjButton1.ForeColor = System.Drawing.Color.White;
            this.rjButton1.GradientEndColor = System.Drawing.Color.DeepSkyBlue;
            this.rjButton1.GradientStartColor = System.Drawing.Color.Aquamarine;
            this.rjButton1.Location = new System.Drawing.Point(59, 529);
            this.rjButton1.Name = "rjButton1";
            this.rjButton1.Padding = new System.Windows.Forms.Padding(23);
            this.rjButton1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.rjButton1.Size = new System.Drawing.Size(294, 77);
            this.rjButton1.TabIndex = 0;
            this.rjButton1.Text = "Start";
            this.rjButton1.UseVisualStyleBackColor = false;
            this.rjButton1.Click += new System.EventHandler(this.rjButton1_Click);
            // 
            // rjButton2
            // 
            this.rjButton2.BackColor = System.Drawing.Color.Red;
            this.rjButton2.BorderColor = System.Drawing.Color.Red;
            this.rjButton2.BorderColour = System.Drawing.Color.Red;
            this.rjButton2.BorderRadius = 40;
            this.rjButton2.BorderSize = 0;
            this.rjButton2.FlatAppearance.BorderSize = 0;
            this.rjButton2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.rjButton2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rjButton2.ForeColor = System.Drawing.Color.Black;
            this.rjButton2.GradientEndColor = System.Drawing.Color.Aquamarine;
            this.rjButton2.GradientStartColor = System.Drawing.Color.DeepSkyBlue;
            this.rjButton2.Location = new System.Drawing.Point(-5, 356);
            this.rjButton2.Name = "rjButton2";
            this.rjButton2.Size = new System.Drawing.Size(74, 88);
            this.rjButton2.TabIndex = 3;
            this.rjButton2.Text = ">";
            this.rjButton2.UseVisualStyleBackColor = false;
            this.rjButton2.Click += new System.EventHandler(this.rjButton2_Click);
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(467, 773);
            this.Controls.Add(this.rjButton2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.rjButton1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form4";
            this.Text = "Form4";
            this.Load += new System.EventHandler(this.Form4_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private RjButton rjButton1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private RjButton rjButton2;
    }
}