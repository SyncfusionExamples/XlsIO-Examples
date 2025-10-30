using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Apply_Scenario
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

                for (int pos = 0; pos < scenarios.Count; pos++)
                {
                    //Apply scenarios
                    scenarios[pos].Show();

                    IWorkbook newBook = excelEngine.Excel.Workbooks.Create(0);

                    IWorksheet newSheet = newBook.Worksheets.AddCopy(worksheet);

                    newSheet.Name = scenarios[pos].Name;

                    //Saving the new workbook
                    newBook.SaveAs(Path.GetFullPath(@"Output/" + scenarios[pos].Name + ".xlsx"));

                    //To restore the cell values from the previous scenario results
                    scenarios["Current % of Change"].Show();
                    scenarios["Current Quantity"].Show();
                }
            }
        }
    }
}




