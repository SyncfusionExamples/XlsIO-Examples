﻿using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Range_to_JSON_without_Schema
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

                //Custom range
                IRange range = worksheet.Range["A1:F100"];

                #region save as JSON
                //Saves the workbook to a JSON filestream, as schema by default
                FileStream outputStream = new FileStream("Excel-Range-To-JSON-filestream-without-schema.json", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAsJson(outputStream, range, false);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                #region Open JSON 
                //Open default JSON
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Excel-Range-To-JSON-filestream-without-schema.json")
                {
                    UseShellExecute = true
                };
                process.Start();
                #endregion
            }
        }
    }
}
