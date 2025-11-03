using System.IO;
using Syncfusion.XlsIO;

namespace AutoShapes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding an AutoShape
                IShape shape1 = worksheet.Shapes.AddAutoShapes(AutoShapeType.RoundedRectangle, 2, 7, 60, 192);
                IShape shape2 = worksheet.Shapes.AddAutoShapes(AutoShapeType.CircularArrow, 8, 7, 60, 192);

                //Set the value inside the shape
                shape1.TextFrame.TextRange.Text = "AutoShape";

                //Format the shape
                shape1.Fill.ForeColorIndex = ExcelKnownColors.Light_blue;
                shape1.TextFrame.VerticalAlignment = ExcelVerticalAlignment.MiddleCentered;

                //Read an AutoShape
                shape1 = worksheet.Shapes[0];
                shape1.TextFrame.TextRange.Text = "RoundedRectangle";

                //Remove an AutoShape
                shape2 = worksheet.Shapes[1];
                shape2.Remove();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AutoShapes.xlsx"));
                #endregion
            }
        }
    }
}




