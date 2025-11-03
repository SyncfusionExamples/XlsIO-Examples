using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Shapes;

namespace Hide_Unhide_Shapes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                IShapes shapes = worksheet.Shapes;
                AutoShapeImpl shape1 = shapes[0] as AutoShapeImpl;

                //Set shape1 to be hidden
                shape1.IsHidden = true;

                AutoShapeImpl shape2 = shapes[1] as AutoShapeImpl;

                //Set shape2 to be visible
                shape2.IsHidden = false;

                //Saving the workbook
                workbook.SaveAs("Output/Output.xlsx");
            }

        }

    }
}
