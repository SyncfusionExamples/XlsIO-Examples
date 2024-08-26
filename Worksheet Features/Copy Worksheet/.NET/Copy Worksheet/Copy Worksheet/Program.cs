using System.IO;
using Syncfusion.XlsIO;

namespace Copy_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream sourceStream = new FileStream("../../../SourceTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook sourceWorkbook = application.Workbooks.Open(sourceStream);

                FileStream destinationStream = new FileStream("../../../DestinationTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook destinationWorkbook = application.Workbooks.Open(destinationStream);

                #region Copy Worksheet
                //Copy first worksheet from the source workbook to the destination workbook
                destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0]);
                destinationWorkbook.ActiveSheetIndex = 1;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/CopyWorksheet.xlsx"), FileMode.Create, FileAccess.Write);
                destinationWorkbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                destinationStream.Dispose();
                sourceStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CopyWorksheet.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
