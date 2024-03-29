﻿
using System;
using System.Windows.Forms;

namespace DataGridView_to_Worksheet
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
        private Button btnCreate;
        private Label label;
        private GroupBox groupBox;

        private void InitializeComponent()
        {
            btnCreate = new Button();
            label = new Label();
            groupBox = new GroupBox();

            //Button
            btnCreate.Location = new System.Drawing.Point(339, 280);
            btnCreate.Size = new System.Drawing.Size(115, 26);
            btnCreate.Text = "Create";
            btnCreate.Click += new EventHandler(btnCreate_Click);

            //Label
            label.Location = new System.Drawing.Point(0, 50);
            label.Size = new System.Drawing.Size(426, 48);
            label.Text = "Click the button to view an Excel spreadsheet generated by Essential XlsIO. Please note that MS Excel Viewer or MS Excel is required to view the resultant document.";
            label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            //Group Box
            groupBox.Location = new System.Drawing.Point(12, 100);
            groupBox.Size = new System.Drawing.Size(442, 151);
            groupBox.Text = "DataGridView";

            //DataGridView to Excel 
            ClientSize = new System.Drawing.Size(466, 333);
            Controls.Add(groupBox);
            Controls.Add(label);
            Controls.Add(btnCreate);
            Text = "DataGridView to Excel";
        }
        #endregion
    }
}

