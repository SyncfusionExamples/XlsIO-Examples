namespace IncomeTaxEmailDistribution
{
    partial class Form1
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
            this.EmailSendButton = new System.Windows.Forms.Button();
            this.excelPathTxtBox = new System.Windows.Forms.TextBox();
            this.partAPathTxtBox = new System.Windows.Forms.TextBox();
            this.partBPathTxtBox = new System.Windows.Forms.TextBox();
            this.ExcelPathLabel = new System.Windows.Forms.Label();
            this.PartAPathLabel = new System.Windows.Forms.Label();
            this.PartBPathLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panColumnIdTxtBox = new System.Windows.Forms.TextBox();
            this.emailColumnIdTxtBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // EmailSendButton
            // 
            this.EmailSendButton.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.EmailSendButton.Location = new System.Drawing.Point(309, 332);
            this.EmailSendButton.Name = "EmailSendButton";
            this.EmailSendButton.Size = new System.Drawing.Size(140, 42);
            this.EmailSendButton.TabIndex = 0;
            this.EmailSendButton.Text = "Send Emails";
            this.EmailSendButton.UseVisualStyleBackColor = true;
            this.EmailSendButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // excelPathTxtBox
            // 
            this.excelPathTxtBox.Location = new System.Drawing.Point(250, 162);
            this.excelPathTxtBox.Name = "excelPathTxtBox";
            this.excelPathTxtBox.Size = new System.Drawing.Size(403, 22);
            this.excelPathTxtBox.TabIndex = 1;
            // 
            // partAPathTxtBox
            // 
            this.partAPathTxtBox.Location = new System.Drawing.Point(250, 210);
            this.partAPathTxtBox.Name = "partAPathTxtBox";
            this.partAPathTxtBox.Size = new System.Drawing.Size(403, 22);
            this.partAPathTxtBox.TabIndex = 2;
            // 
            // partBPathTxtBox
            // 
            this.partBPathTxtBox.Location = new System.Drawing.Point(250, 260);
            this.partBPathTxtBox.Name = "partBPathTxtBox";
            this.partBPathTxtBox.Size = new System.Drawing.Size(403, 22);
            this.partBPathTxtBox.TabIndex = 3;
            // 
            // ExcelPathLabel
            // 
            this.ExcelPathLabel.AutoSize = true;
            this.ExcelPathLabel.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExcelPathLabel.Location = new System.Drawing.Point(90, 162);
            this.ExcelPathLabel.Name = "ExcelPathLabel";
            this.ExcelPathLabel.Size = new System.Drawing.Size(111, 21);
            this.ExcelPathLabel.TabIndex = 4;
            this.ExcelPathLabel.Text = "Excel File Path";
            // 
            // PartAPathLabel
            // 
            this.PartAPathLabel.AutoSize = true;
            this.PartAPathLabel.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PartAPathLabel.Location = new System.Drawing.Point(89, 210);
            this.PartAPathLabel.Name = "PartAPathLabel";
            this.PartAPathLabel.Size = new System.Drawing.Size(118, 21);
            this.PartAPathLabel.TabIndex = 5;
            this.PartAPathLabel.Text = "Part A File Path";
            // 
            // PartBPathLabel
            // 
            this.PartBPathLabel.AutoSize = true;
            this.PartBPathLabel.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PartBPathLabel.Location = new System.Drawing.Point(89, 261);
            this.PartBPathLabel.Name = "PartBPathLabel";
            this.PartBPathLabel.Size = new System.Drawing.Size(118, 21);
            this.PartBPathLabel.TabIndex = 6;
            this.PartBPathLabel.Text = "Part B File Path";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(89, 115);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 21);
            this.label1.TabIndex = 7;
            this.label1.Text = "Email ID Column";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(89, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(155, 21);
            this.label2.TabIndex = 8;
            this.label2.Text = "Tax Payer ID Column";
            // 
            // panColumnIdTxtBox
            // 
            this.panColumnIdTxtBox.Location = new System.Drawing.Point(250, 67);
            this.panColumnIdTxtBox.Name = "panColumnIdTxtBox";
            this.panColumnIdTxtBox.Size = new System.Drawing.Size(403, 22);
            this.panColumnIdTxtBox.TabIndex = 9;
            // 
            // emailColumnIdTxtBox
            // 
            this.emailColumnIdTxtBox.Location = new System.Drawing.Point(250, 115);
            this.emailColumnIdTxtBox.Name = "emailColumnIdTxtBox";
            this.emailColumnIdTxtBox.Size = new System.Drawing.Size(403, 22);
            this.emailColumnIdTxtBox.TabIndex = 10;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(689, 153);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(106, 31);
            this.button2.TabIndex = 12;
            this.button2.Text = "browse";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(689, 208);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(106, 32);
            this.button1.TabIndex = 13;
            this.button1.Text = "browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(689, 260);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(106, 31);
            this.button3.TabIndex = 14;
            this.button3.Text = "browse";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(907, 408);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.emailColumnIdTxtBox);
            this.Controls.Add(this.panColumnIdTxtBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PartBPathLabel);
            this.Controls.Add(this.PartAPathLabel);
            this.Controls.Add(this.ExcelPathLabel);
            this.Controls.Add(this.partBPathTxtBox);
            this.Controls.Add(this.partAPathTxtBox);
            this.Controls.Add(this.excelPathTxtBox);
            this.Controls.Add(this.EmailSendButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button EmailSendButton;
        private System.Windows.Forms.TextBox excelPathTxtBox;
        private System.Windows.Forms.TextBox partAPathTxtBox;
        private System.Windows.Forms.TextBox partBPathTxtBox;
        private System.Windows.Forms.Label ExcelPathLabel;
        private System.Windows.Forms.Label PartAPathLabel;
        private System.Windows.Forms.Label PartBPathLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox panColumnIdTxtBox;
        private System.Windows.Forms.TextBox emailColumnIdTxtBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}

