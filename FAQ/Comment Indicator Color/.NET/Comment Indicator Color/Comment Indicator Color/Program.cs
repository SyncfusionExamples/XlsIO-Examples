using Syncfusion.Drawing;
using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Set the default version as Excel 2016
            excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
            // Create a new workbook
            IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            int height = 8;
            int width = 5;

            // Add a comment to cell A1
            ICommentShape comment = worksheet.Range["A1"].AddComment();

            // Create a right triangle shape to act as custom indicator
            IShape shape = worksheet.Shapes.AddAutoShapes(AutoShapeType.RightTriangle, 1, 1, width, height);
            // Position the shape at the upper-right corner of the cell
            shape.Left = (int)worksheet.GetColumnWidthInPixels(1) - height;
            // Rotate to point downward
            shape.ShapeRotation = 180;
            // Set the indicator (shape) color to blue
            shape.Fill.ForeColor = Color.Blue;

            // Save the workbook to a file
            workbook.SaveAs(Path.GetFullPath("Output/CommentIndicatorColor.xlsx"));
        }
    }
}
