using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Comments
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

                //Remove all the comments in worksheet
                worksheet.Comments.Clear();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveComments.xlsx"));
                #endregion
            }
        }
    }
}





