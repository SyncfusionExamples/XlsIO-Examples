
using System;
using System.Windows.Forms;

namespace Edit_Excel
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
        private Button btnEdit;
        private Label label;

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "Form1";

            btnEdit = new Button();
            label = new Label();

            //Label
            label.Location = new System.Drawing.Point(0, 40);
            label.Size = new System.Drawing.Size(426, 35);
            label.Text = "Click the button to view an Excel spreadsheet edited by Essential XlsIO. Please note that MS Excel Viewer or MS Excel is required to view the resultant document.";
            label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            //Button
            btnEdit.Location = new System.Drawing.Point(180, 110);
            btnEdit.Size = new System.Drawing.Size(85, 26);
            btnEdit.Text = "Edit Excel";
            btnEdit.Click += new EventHandler(btnEdit_Click);

            //Create Spreadsheet
            ClientSize = new System.Drawing.Size(450, 150);
            Controls.Add(label);
            Controls.Add(btnEdit);
            Text = "Edit Excel";
        }

        #endregion
    }
}

