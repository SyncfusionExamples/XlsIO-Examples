using System.IO;
using Syncfusion.XlsIO;

namespace Autofit_Rows_and_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Autofit Rows
                //Autofit applied to a single row
                worksheet.AutofitRow(3);

                //Autofit applied to multiple rows
                worksheet.Range["6:10"].AutofitRows();
                #endregion

                #region Autofit Columns
                //Autofit applied to a single column
                worksheet.AutofitColumn(2);

                //Autofit applied to multiple columns
                worksheet.Range["E:G"].AutofitColumns();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AutofitRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





