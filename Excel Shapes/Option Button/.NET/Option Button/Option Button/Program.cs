using System.IO;
using Syncfusion.XlsIO;

namespace Option_Button
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

                //Create an Option Button
                IOptionButtonShape optionButton1 = sheet.OptionButtons.AddOptionButton(2, 3);
                //Assign a value to the Option Button
                optionButton1.Text = "Fed Ex";

                //Format the control
                optionButton1.Fill.FillType = ExcelFillType.SolidColor;
                optionButton1.Fill.ForeColor = Syncfusion.Drawing.Color.Yellow;
                //Change the check state
                optionButton1.CheckState = ExcelCheckState.Checked;

                //Create an Option Button
                IOptionButtonShape optionButton2 = sheet.OptionButtons.AddOptionButton(5, 3);
                //Assign a value to the Option Button
                optionButton2.Text = "DHL";

                //Read an Option Button
                optionButton1 = sheet.OptionButtons[0];
                optionButton1.CheckState = ExcelCheckState.Unchecked;

                //Remove an Option Button
                optionButton2 = sheet.OptionButtons[1];
                optionButton2.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/OptionButton.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




