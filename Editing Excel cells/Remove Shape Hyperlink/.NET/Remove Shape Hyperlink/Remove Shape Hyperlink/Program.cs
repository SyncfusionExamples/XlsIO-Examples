using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Shape_Hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath("Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Removing hyperlink from sheet with Index
                worksheet.HyperLinks.RemoveAt(0);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveShapeHyperlink.xlsx"));
                #endregion
            }
        }
    }
}




