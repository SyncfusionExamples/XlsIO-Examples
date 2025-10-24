using System.IO;
using Syncfusion.XlsIO;

namespace Clear_Content
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

                #region Clear Content
                //Clearing content and formatting in C3
                worksheet.Range["C3"].Clear(true);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ClearContent.xlsx"));
                #endregion
            }
        }
    }
}




