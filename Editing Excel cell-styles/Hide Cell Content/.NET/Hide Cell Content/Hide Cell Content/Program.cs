using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Cell_Content
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
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Hide Cell Content
                //Assign values to a range of cells in the worksheet
                worksheet.Range["A1:A10"].Text = "Hide Cell Content";

                //Apply number format for the cell to hide its content
                worksheet.Range["A5"].NumberFormat = ";;;";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/HideCellContent.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




