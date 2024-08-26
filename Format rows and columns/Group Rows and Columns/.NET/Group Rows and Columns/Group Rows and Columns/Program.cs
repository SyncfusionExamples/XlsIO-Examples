using System.IO;
using Syncfusion.XlsIO;

namespace Group_Rows_and_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate - ToGroup.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Group Rows
                //Group Rows
                worksheet.Range["A3:A7"].Group(ExcelGroupBy.ByRows, true);
                worksheet.Range["A11:A16"].Group(ExcelGroupBy.ByRows);
                #endregion

                #region Group Columns
                //Group Columns
                worksheet.Range["C1:D1"].Group(ExcelGroupBy.ByColumns, false);
                worksheet.Range["F1:G1"].Group(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/GroupRowsandColumns.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




