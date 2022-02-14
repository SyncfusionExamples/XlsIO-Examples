using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Types_of_Calculated_Value
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                bool B1_PreviousValue = sheet["B1"].FormulaBoolValue;
                DateTime C1_PreviousValue = sheet["C1"].FormulaDateTime;
                double D1_PreviousValue = sheet["D1"].FormulaNumberValue;

                //Previous Value '2'
                sheet["E1"].Number = 3;

                sheet.EnableSheetCalculations();

                //It has formula 'ISEVEN(E1)'
                //Returns the calculated value of a formula as Boolean                
                string value = sheet.Range["B1"].CalculatedValue;
                bool B1_LatestValue = sheet["B1"].FormulaBoolValue;

                //It has formula 'TODAY()'
                //Returns the calculated value of a formula as DateTime                
                value = sheet.Range["C1"].CalculatedValue;
                DateTime C1_LatestValue = sheet["C1"].FormulaDateTime;

                //It has formula '=E1'
                //Returns the calculated value of a formula as double                
                value = sheet.Range["D1"].CalculatedValue;
                double D1_LatestValue = sheet["D1"].FormulaNumberValue;

                sheet.DisableSheetCalculations();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Formula.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Formula.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
