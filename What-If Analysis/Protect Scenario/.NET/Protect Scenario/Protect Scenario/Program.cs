using System.IO;
using Syncfusion.XlsIO;
using System;
using static System.Net.Mime.MediaTypeNames;

namespace Protect_Scenario
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/WhatIfAnalysisTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                inputStream.Dispose();

                IWorksheet worksheet = workbook.Worksheets[0];

                //Enable worksheet protection
                worksheet.Protect("scenario");

                // Access the collection of scenarios in the worksheet
                IScenarios scenarios = worksheet.Scenarios;

                //To make a scenario editable after protecting the sheet
                scenarios[0].Locked = false;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ProtectScenario.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

            }
        }
    }
}




