using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IncomeTaxEmailDistribution
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string tinColumnId = panColumnIdTxtBox.Text.Trim('"');

            string emailColumnId = emailColumnIdTxtBox.Text.Trim('"');

            string excelFilePath = excelPathTxtBox.Text.Trim('"');

            string[] partAPathFiles = Directory.GetFiles(partAPathTxtBox.Text.Trim('"'));

            string[] partBPathFiles = Directory.GetFiles(partBPathTxtBox.Text.Trim('"'));

            if(partAPathFiles != null && partAPathFiles.Length == 0)
            {
                MessageBox.Show("No files found in Part A Path");
                return;
            }
            if(partBPathFiles != null && partBPathFiles.Length == 0)
            {
                MessageBox.Show("No files found in Part B Path");
                return;
            }
            if(!File.Exists(excelFilePath))
            {
                MessageBox.Show("Excel File not found");
                return;
            }
            
            using(ExcelEngine excelengine = new ExcelEngine())
            {
                IApplication application = excelengine.Excel;
                IWorkbook workbook = application.Workbooks.Open(excelFilePath);

                IWorksheet worksheet = workbook.Worksheets[0];

                IRange used = worksheet.UsedRange;

                int startRow = used.Row + 1;
                int endRow = used.LastRow;
                int statusColumn = used.LastColumn + 1;


                for(int row = startRow;row <= endRow;row++)
                {
                    //Check the automation status of the tax payer report
                    if (worksheet.Range[row, statusColumn].Value == "Sent")
                    {
                        continue;
                    }

                    //Converting the Column Id to Column Index
                    IRange tinRange = worksheet[tinColumnId + row.ToString()];

                    string tinID = worksheet.Range[row, tinRange.Column].Value;

                    if(string.IsNullOrEmpty(tinID))
                    {
                        worksheet[row, statusColumn].Value = "TIN not found";
                        continue;
                    }
                    
                    IRange emailRange = worksheet[emailColumnId + row.ToString()];

                    string emailId = worksheet.Range[row, emailRange.Column].Value;

                    string partAPath = partAPathFiles.Where(x => x.Contains(tinID + "_")).FirstOrDefault();

                    if (partAPath == null)
                    {
                        worksheet[row, statusColumn].Value = "Part A File not found";
                        continue;
                    }
                    
                    string partBPath = partBPathFiles.Where(x => x.Contains(tinID + "_")).FirstOrDefault();

                    if (partBPath == null)
                    {
                        worksheet[row, statusColumn].Value = "Part B File not found";
                        continue;
                    }
                    try
                    {
                        //Your maid id
                        string from = "Yourmail@abc.com";

                        string subject = "Form 16 for the financial year 2022- 2023";

                        string mailBody = "<p>Hi,</p>\r\n<p>Thank you for your kind cooperation</p>\r\n<p>We have attached the form 16 document pertaining to the financial year (2022-2023) tax deduction. Employees are requested to review and file their Income tax return before due date.</p>\r\n<p>The last date of filing of Income Tax return is July 31, 2023.</p>\r\n<p>For further assistance, please feel free to get back to us.</p>\r\n<p>Thanks and Regards,</p>\r\n<p>Accounts Team.</p>";

                        SendEMail(from, emailId, subject, mailBody, partAPath, partBPath);

                        worksheet[row, statusColumn].Value = "Sent";
                    }
                    catch(Exception ex){
                        worksheet[row, statusColumn].Value = ex.ToString();
                    }
                }

                Thread.Sleep(1000);

                // Saving the Workbook containing successful and failed records

                string filePath = Path.GetFileNameWithoutExtension(excelFilePath) + "_Updated_" + DateTime.Now.ToString().Replace(":", "-") + Path.GetExtension(excelFilePath); 
               
                workbook.SaveAs(filePath);
                
            }
            Close();
            
        }
        private static void SendEMail(string from, string recipients, string subject, string body, string attachmentPartA, string attachmentPartB)
        {
            //Creates the email message
            var emailMessage = new MailMessage(from, recipients);
            //Adds the subject for email
            emailMessage.Subject = subject;
            //Sets the HTML string as email body
            emailMessage.IsBodyHtml = true;
            emailMessage.Body = body;
            emailMessage.Attachments.Add(new Attachment(attachmentPartA));
            emailMessage.Attachments.Add(new Attachment(attachmentPartB));

            //Sends the email with prepared message
            using (var client = new SmtpClient())
            {
                //Update your SMTP Server address here
                client.Host = "outlook.office365.com";
                client.UseDefaultCredentials = false;
                //Update your email credentials here. Need to generate app password for your mail id
                client.Credentials = new System.Net.NetworkCredential(from, "apppassword");
                client.Port = 587;
                client.EnableSsl = true;
                client.Send(emailMessage);                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openFileDialog1.ShowDialog();
            // if a file is selected
            if (result == DialogResult.OK)
            {
                // Set the selected file URL to the textbox
                this.excelPathTxtBox.Text = this.openFileDialog1.FileName;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            // if a file is selected
            if (result == DialogResult.OK)
            {
                // Set the selected file URL to the textbox
                this.partAPathTxtBox.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            // if a file is selected
            if (result == DialogResult.OK)
            {
                // Set the selected file URL to the textbox
                this.partBPathTxtBox.Text = this.folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
