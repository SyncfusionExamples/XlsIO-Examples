using System;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.IO;


namespace Edit_Excel.Components.Data
{
    public class ExcelSercive
    {
        public MemoryStream EditExcel()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel workbook into IWorkbook
                IWorkbook workbook = application.Workbooks.Open("InputTemplate.xlsx");

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Save the document as a stream and retrun the stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created Excel document to MemoryStream.
                    workbook.SaveAs(stream);
                    return stream;
                }
            }
        }
    }
}
