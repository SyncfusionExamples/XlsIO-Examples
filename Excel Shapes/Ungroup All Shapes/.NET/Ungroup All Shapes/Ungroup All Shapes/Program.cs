using System.IO;
using Syncfusion.XlsIO;

namespace Ungroup_All_Shapes
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

                IShapes shapes = worksheet.Shapes;

                // Ungroup group shape and its all the inner shapes.
                shapes.Ungroup(shapes[0] as IGroupShape, true);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/UngroupAllShapes.xlsx"));
                #endregion
            }
        }
    }
}





