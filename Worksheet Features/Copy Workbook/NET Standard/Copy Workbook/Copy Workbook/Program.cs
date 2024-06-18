using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Interfaces;

namespace Copy_Workbook
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream sourceStream = new FileStream("../../../Data/SourceWorkbookTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook sourceWorkbook = application.Workbooks.Open(sourceStream);
                FileStream destinationStream = new FileStream("../../../Data/DestinationWorkbookTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook destinationWorkbook = application.Workbooks.Open(destinationStream);

                //Clone the workbook
                destinationWorkbook = sourceWorkbook.Clone();
               
                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                destinationWorkbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                destinationStream.Dispose();
                sourceStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }   
        }
    }
}