﻿using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace ConvertExcelToImage
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream excelStream = new FileStream("Sample.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Create the MemoryStream to save the image  
                MemoryStream imageStream = new MemoryStream();

                //Save the converted image to MemoryStream
                worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
                imageStream.Position = 0;

                #region Save
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Sample.jpeg"), FileMode.Create, FileAccess.Write);
                imageStream.CopyTo(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                excelStream.Dispose();
            }
        }
    }
}