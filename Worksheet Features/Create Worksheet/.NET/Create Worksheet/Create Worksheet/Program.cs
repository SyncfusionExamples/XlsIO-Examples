using System.IO;
using Syncfusion.XlsIO;

namespace Create_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                #region Create
                //The new workbook is created with 5 worksheets
                IWorkbook workbook = application.Workbooks.Create(5);
                //Creating a new sheet
                IWorksheet worksheet = workbook.Worksheets.Create();
                //Creating a new sheet with name “Sample”
                IWorksheet namedSheet = workbook.Worksheets.Create("Sample");
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateWorksheet.xlsx"));
                #endregion
            }
        }
    }
}




