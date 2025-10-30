using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Hide_Scenario
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

                //Access the collection of scenarios in the worksheet
                IScenarios scenarios = worksheet.Scenarios;

                //Disable the protection for a specific scenario
                scenarios["Increased % of Change"].Hidden = true;

                //Enable worksheet protection
                worksheet.Protect("Scenario");

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HideScenario.xlsx"));
                #endregion
            }
        }
    }
}




