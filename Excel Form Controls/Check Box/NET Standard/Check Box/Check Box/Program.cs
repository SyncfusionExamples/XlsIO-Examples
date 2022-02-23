using System.IO;
using Syncfusion.XlsIO;


namespace Check_Box
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

                //Create a check box with cell link
                ICheckBoxShape checkBoxRed = sheet.CheckBoxes.AddCheckBox(2, 4, 20, 75);
                checkBoxRed.Text = "Red";
                checkBoxRed.CheckState = ExcelCheckState.Unchecked;
                checkBoxRed.LinkedCell = sheet["B2"];
                ICheckBoxShape checkBoxBlue = sheet.CheckBoxes.AddCheckBox(4, 4, 20, 75);
                checkBoxBlue.Text = "Blue";
                checkBoxBlue.CheckState = ExcelCheckState.Checked;
                checkBoxBlue.LinkedCell = sheet["B4"];

                //Read a check box
                checkBoxRed = sheet.CheckBoxes[0];
                checkBoxRed.CheckState = ExcelCheckState.Checked;

                //Remove a check box
                checkBoxBlue = sheet.CheckBoxes[1];
                checkBoxBlue.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CheckBox.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CheckBox.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
