﻿using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Worksheet_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Image.png"), FileMode.Create, FileAccess.Write);
                sheet.ConvertToImage(sheet.UsedRange, outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}