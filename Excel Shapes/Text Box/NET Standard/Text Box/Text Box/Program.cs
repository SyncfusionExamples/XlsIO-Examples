using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Text_Box
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Creates a new Text Box
                ITextBoxShape textbox = sheet.TextBoxes.AddTextBox(2, 2, 30, 200);
                textbox.Text = "Text Box 1";
                textbox = sheet.TextBoxes.AddTextBox(6, 2, 30, 200);
                textbox.Text = "Text Box 2";

                //Reads a Text Box
                ITextBoxShape shape1 = sheet.TextBoxes[0];
                shape1.Text = "TextBox";

                //Format the control
                shape1.Fill.ForeColor = Color.Gold;
                shape1.Fill.BackColor = Color.Black;
                shape1.Fill.Pattern = ExcelGradientPattern.Pat_90_Percent;

                //Remove a Text Box
                ITextBoxShape shape2 = sheet.TextBoxes[1];
                shape2.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("TextBox.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("TextBox.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
