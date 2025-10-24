using System.IO;
using Syncfusion.XlsIO;

namespace Modify_Hyperlink
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

                #region Modify Hyperlink
                //Modifying hyperlink’s text to display
                IHyperLink hyperlink = worksheet.Range["C5"].Hyperlinks[0];
                hyperlink.TextToDisplay = "Syncfusion";
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ModifyHyperlink.xlsx"));
                #endregion
            }
        }
    }
}




