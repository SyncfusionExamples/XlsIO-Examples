using System.IO;
using Syncfusion.XlsIO;

namespace Modify_Shape_Hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Modify Shape Hyperlink
                //Modifying hyperlink’s screen tip through IWorksheet instance
                IHyperLink hyperlink = worksheet.HyperLinks[0];
                hyperlink.ScreenTip = "Syncfusion";

                //Modifying hyperlink’s screen tip through IShape instance
                hyperlink = worksheet.Shapes[1].Hyperlink;
                hyperlink.ScreenTip = "Mail Syncfusion";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ModifyShapeHyperlink.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ModifyShapeHyperlink.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
