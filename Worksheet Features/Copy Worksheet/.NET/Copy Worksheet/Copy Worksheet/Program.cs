using System.IO;
using Syncfusion.XlsIO;

namespace Copy_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook sourceWorkbook = application.Workbooks.Open(Path.GetFullPath(@"Data/SourceTemplate.xlsx"));
                IWorkbook destinationWorkbook = application.Workbooks.Open(Path.GetFullPath(@"Data/DestinationTemplate.xlsx"));

                #region Copy Worksheet
                //Copy first worksheet from the source workbook to the destination workbook
                destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0]);
                destinationWorkbook.ActiveSheetIndex = 1;
                #endregion

                #region Save
                //Saving the workbook
                destinationWorkbook.SaveAs(Path.GetFullPath("Output/CopyWorksheet.xlsx"));
                #endregion
            }
        }
    }
}





