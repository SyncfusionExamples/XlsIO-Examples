using System.IO;
using Syncfusion.XlsIO;

namespace Ungroup_Rows_and_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate - ToUngroup.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Un-Group Rows
                //Ungroup Rows
                worksheet.Range["A3:A7"].Ungroup(ExcelGroupBy.ByRows);
                #endregion

                #region Un-Group Columns
                //Ungroup Columns
                worksheet.Range["C1:D1"].Ungroup(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/UngroupRowsandColumns.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





