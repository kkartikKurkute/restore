namespace ExcelApp
{
    partial class MsgBox
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblMsg = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSaveNext = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.btnTryAgain = new System.Windows.Forms.Button();
            this.picBox = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBox)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.btnSaveNext);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.btnTryAgain);
            this.panel1.Controls.Add(this.picBox);
            this.panel1.Controls.Add(this.lblMsg);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(399, 220);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lblMsg.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.Location = new System.Drawing.Point(109, 105);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(197, 22);
            this.lblMsg.TabIndex = 6;
            this.lblMsg.Text = "The Image are same";
            this.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(270, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 11;
            this.label3.Visible = false;
            // 
            // btnSaveNext
            // 
            this.btnSaveNext.AutoSize = true;
            this.btnSaveNext.BackColor = System.Drawing.Color.White;
            this.btnSaveNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnSaveNext.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnSaveNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSaveNext.Image = global::ExcelApp.Properties.Resources.SaveNext;
            this.btnSaveNext.Location = new System.Drawing.Point(113, 155);
            this.btnSaveNext.Margin = new System.Windows.Forms.Padding(0);
            this.btnSaveNext.Name = "btnSaveNext";
            this.btnSaveNext.Size = new System.Drawing.Size(119, 47);
            this.btnSaveNext.TabIndex = 10;
            this.btnSaveNext.UseVisualStyleBackColor = true;
            this.btnSaveNext.Click += new System.EventHandler(this.btnSaveNext_Click);
            // 
            // button2
            // 
            this.button2.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Image = global::ExcelApp.Properties.Resources.RevealHint;
            this.button2.Location = new System.Drawing.Point(250, 158);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(116, 44);
            this.button2.TabIndex = 9;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnTryAgain
            // 
            this.btnTryAgain.AutoSize = true;
            this.btnTryAgain.BackColor = System.Drawing.Color.White;
            this.btnTryAgain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnTryAgain.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnTryAgain.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTryAgain.Image = global::ExcelApp.Properties.Resources.TryAgain;
            this.btnTryAgain.Location = new System.Drawing.Point(10, 155);
            this.btnTryAgain.Margin = new System.Windows.Forms.Padding(0);
            this.btnTryAgain.Name = "btnTryAgain";
            this.btnTryAgain.Size = new System.Drawing.Size(91, 47);
            this.btnTryAgain.TabIndex = 8;
            this.btnTryAgain.UseVisualStyleBackColor = true;
            this.btnTryAgain.Click += new System.EventHandler(this.btnTryAgain_Click_1);
            // 
            // picBox
            // 
            this.picBox.Location = new System.Drawing.Point(171, 19);
            this.picBox.Name = "picBox";
            this.picBox.Size = new System.Drawing.Size(61, 57);
            this.picBox.TabIndex = 7;
            this.picBox.TabStop = false;
            // 
            // MsgBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(399, 220);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.HelpButton = true;
            this.Name = "MsgBox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.MsgBox_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBox)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnSaveNext;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnTryAgain;
        private System.Windows.Forms.PictureBox picBox;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.Label label3;
    }
}