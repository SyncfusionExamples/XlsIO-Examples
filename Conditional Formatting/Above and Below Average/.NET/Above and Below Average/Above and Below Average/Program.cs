﻿using System.IO;
using Syncfusion.XlsIO;

namespace Above_and_Below_Average
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

                //Applying conditional formatting to "M6:M35"
                IConditionalFormats formats = worksheet.Range["M6:M35"].ConditionalFormats;
                IConditionalFormat format = formats.AddCondition();

                //Applying above or below average rule in the conditional formatting
                format.FormatType = ExcelCFType.AboveBelowAverage;
                IAboveBelowAverage aboveBelowAverage = format.AboveBelowAverage;

                //Set AverageType as Below for AboveBelowAverage rule.
                aboveBelowAverage.AverageType = ExcelCFAverageType.Below;

                //Set color for Conditional Formattting.
                format.FontColorRGB = Syncfusion.Drawing.Color.FromArgb(255, 255, 255);
                format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(166, 59, 38);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("AboveAndBelowAverage.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("AboveAndBelowAverage.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
