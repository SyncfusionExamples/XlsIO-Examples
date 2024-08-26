using System.IO;
using Syncfusion.XlsIO;
using System;
using static System.Net.Mime.MediaTypeNames;

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

                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/WhatIfAnalysisTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                inputStream.Dispose();

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

                    //Saving the new workbook as a stream
                    using (FileStream stream = new FileStream(scenarios[pos].Name + ".xlsx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        newBook.SaveAs(stream);
                    }

                    //To restore the cell values from the previous scenario results
                    scenarios["Current % of Change"].Show();
                    scenarios["Current Quantity"].Show();
                }
            }
        }
    }
}




