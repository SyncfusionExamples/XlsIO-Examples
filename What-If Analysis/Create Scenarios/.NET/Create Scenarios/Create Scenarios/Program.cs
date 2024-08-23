using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Create_Scenarios
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

                // Access the collection of scenarios in the worksheet
                IScenarios scenarios = worksheet.Scenarios;

                //Initialize list objects with different values for scenarios
                List<object> currentChangePercentage_Values = new List<object> { 0.23, 0.8, 1.1, 0.5, 0.35, 0.2 };
                List<object> increasedChangePercentage_Values = new List<object> { 0.45, 0.56, 0.9, 0.5, 0.58, 0.43 };
                List<object> decreasedChangePercentage_Values = new List<object> { 0.3, 0.2, 0.5, 0.3, 0.5, 0.23 };
                List<object> currentQuantity_Values = new List<object> { 1500, 3000, 5000, 4000, 500, 4000 };
                List<object> increasedQuantity_Values = new List<object> { 1000, 5000, 4500, 3900, 10000, 8900 };
                List<object> decreasedQuantity_Values = new List<object> { 1000, 2000, 3000, 3000, 300, 4000 };

                //Add scenarios in the worksheet with different values for the same cells
                scenarios.Add("Current % of Change", worksheet.Range["F5:F10"], currentChangePercentage_Values);
                scenarios.Add("Increased % of Change", worksheet.Range["F5:F10"], increasedChangePercentage_Values);
                scenarios.Add("Decreased % of Change", worksheet.Range["F5:F10"], decreasedChangePercentage_Values);
                scenarios.Add("Current Quantity", worksheet.Range["D5:D10"], currentQuantity_Values);
                scenarios.Add("Increased Quantity", worksheet.Range["D5:D10"], increasedQuantity_Values);
                scenarios.Add("Decreased Quantity", worksheet.Range["D5:D10"], decreasedQuantity_Values);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateScenarios.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}
