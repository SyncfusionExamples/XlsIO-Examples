﻿using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Create_Sparkline
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
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Add SparklineGroups
                ISparklineGroup sparklineGroup = sheet.SparklineGroups.Add();

                //Add SparkLineType
                sparklineGroup.SparklineType = SparklineType.Line;
                sparklineGroup.MarkersColor = Color.BlueViolet;

                //Add sparklines
                ISparklines sparklines = sparklineGroup.Add();
                IRange dataRange = sheet.Range["D6:G17"];
                IRange referenceRange = sheet.Range["H6:H17"];
                sparklines.Add(dataRange, referenceRange);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Sparklines.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Sparklines.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
