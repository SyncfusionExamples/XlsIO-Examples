using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Edit_Excel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //A existing workbook is opened.             
                FileStream sampleFile = new FileStream("Data/InputTemplate.xlsx", FileMode.Open);
                IWorkbook workbook = application.Workbooks.Open(sampleFile);

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                //Saving the workbook as stream
                FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);

                //Dispose the stream
                stream.Dispose();
            }
        }
    }
}




