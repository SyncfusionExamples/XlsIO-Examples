using System;
using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Add_Oval_Shape_Chart
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

                //Add chart to worksheet
                IChart chart = worksheet.Charts.Add();

                //Add oval shape to chart
                IShape shape = chart.Shapes.AddAutoShapes(AutoShapeType.Oval, 20, 60, 500, 400);

                //Format the shape
                shape.Line.ForeColorIndex = ExcelKnownColors.Red;

                //Add the text to the oval shape and set the text alignment on the shape
                shape.TextFrame.TextRange.Text = "This is an oval shape";
                shape.TextFrame.VerticalAlignment = ExcelVerticalAlignment.MiddleCentered;
                shape.TextFrame.HorizontalAlignment = ExcelHorizontalAlignment.CenterMiddle;

                #region Save
                //Saving the workbook
                workbook.SaveAs("Output.xlsx");
                #endregion
            }
        }
    }
}