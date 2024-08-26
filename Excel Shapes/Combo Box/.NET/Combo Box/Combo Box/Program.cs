using System.IO;
using Syncfusion.XlsIO;

namespace Combo_Box
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

                //Filling Values
                sheet["A2"].Text = "RGB colors";
                sheet["A3"].Text = "Red";
                sheet["A4"].Text = "Green";
                sheet["A5"].Text = "Blue";
                sheet["B5"].Text = "Selected Index";

                //Create a Combo Box
                IComboBoxShape comboBox1 = sheet.ComboBoxes.AddComboBox(2, 3, 20, 100);
                //Assign a value to the Combo Box
                comboBox1.ListFillRange = sheet["A3:A5"];
                comboBox1.LinkedCell = sheet["C5"];
                comboBox1.SelectedIndex = 2;

                //Create a Combo Box
                IComboBoxShape comboBox2 = sheet.ComboBoxes.AddComboBox(5, 3, 20, 100);
                //Assign a value to the Combo Box
                comboBox2.ListFillRange = sheet["A3:A5"];
                comboBox2.LinkedCell = sheet["C5"];
                comboBox2.SelectedIndex = 1;

                //Read a Combo Box
                comboBox1 = sheet.ComboBoxes[0];
                comboBox1.SelectedIndex = 1;

                //Remove a Combo Box
                comboBox2 = sheet.ComboBoxes[1];
                comboBox2.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ComboBox.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




