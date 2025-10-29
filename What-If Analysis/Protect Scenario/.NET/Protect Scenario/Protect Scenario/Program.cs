using System.IO;
using Syncfusion.XlsIO;
using System;

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

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/WhatIfAnalysisTemplate.xlsx"), ExcelOpenType.Automatic);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Enable worksheet protection
                worksheet.Protect("scenario");

                // Access the collection of scenarios in the worksheet
                IScenarios scenarios = worksheet.Scenarios;

                //To make a scenario editable after protecting the sheet
                scenarios[0].Locked = false;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ProtectScenario.xlsx"));
                #endregion

            }
        }
    }
}




