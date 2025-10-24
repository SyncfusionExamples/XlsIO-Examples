using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Hyperlink
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

                #region Remove Hyperlink
                //Removing Hyperlink from Range "C7"
                worksheet.Range["C7"].Hyperlinks.RemoveAt(0);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveHyperlink.xlsx"));
                #endregion
            }
        }
    }
}




