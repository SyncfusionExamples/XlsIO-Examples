using System;
using System.IO;
using Syncfusion.XlsIO;

namespace SetAndFormatTimeValues
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

                TimeSpan ts = new TimeSpan(12, 32, 38);

                //Convert the TimeSpan to a fractional day value that Excel understands
                double excelTimeValue = ts.TotalDays;

                //Set value in cell
                sheet.SetValueRowCol(excelTimeValue, 1, 1);

                //Apply the time format to the cell to display it as 'hh:mm:ss'
                sheet.Range[1, 1].NumberFormat = "hh:mm:ss";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}