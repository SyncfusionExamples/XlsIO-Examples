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
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                IShapes shapes = worksheet.Shapes;

                // Ungroup group shape and its all the inner shapes.
                shapes.Ungroup(shapes[0] as IGroupShape, true);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("UngroupAllShapes.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("UngroupAllShapes.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
