using System.IO;
using Syncfusion.XlsIO;

namespace Delete_Named_Range
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
                IWorksheet sheet = workbook.Worksheets[0];

                //Deleting named range object
                IName name = workbook.Names[0];
                name.Delete();

                //Deleting named range from workbook
                workbook.Names["BookLevelName3"].Delete();
                //Deleting named range from worksheet
                sheet.Names["SheetLevelName2"].Delete();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/DeleteNamedRange.xlsx"));
                #endregion
            }
        }
    }
}





