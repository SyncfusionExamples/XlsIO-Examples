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

                FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                IShapes shapes = worksheet.Shapes;
                AutoShapeImpl shape1 = shapes[0] as AutoShapeImpl;

                //Set shape1 to be hidden
                shape1.IsHidden = true;

                AutoShapeImpl shape2 = shapes[1] as AutoShapeImpl;

                //Set shape2 to be visible
                shape2.IsHidden = false;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                workbook.Close();
                excelEngine.Dispose();
            }

        }

    }
}
