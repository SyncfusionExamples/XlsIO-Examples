using System.IO;
using Syncfusion.XlsIO;

namespace Create_Slicer
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Access the table.
                IListObject table = sheet.ListObjects[0];

                //Add slicer for the table.
                sheet.Slicers.Add(table, 3, 11, 2);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateSlicer.xlsx"));
                #endregion
            }
        }
    }
}





