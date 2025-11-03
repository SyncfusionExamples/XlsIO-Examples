using System.IO;
using Syncfusion.XlsIO;

namespace Ungroup_Shapes
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

                // Ungroup group shape.
                shapes.Ungroup(shapes[0] as IGroupShape);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/UngroupShapes.xlsx"));
                #endregion
            }
        }
    }
}





