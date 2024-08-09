﻿using System.IO;
using Syncfusion.XlsIO;
using System;
using static System.Net.Mime.MediaTypeNames;

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

                FileStream inputStream = new FileStream("../../../Data/WhatIfAnalysisTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                inputStream.Dispose();

                IWorksheet worksheet = workbook.Worksheets[0];

                //Access the collection of scenarios in the worksheet
                IScenarios scenarios = worksheet.Scenarios;

                //Disable the protection for a specific scenario
                scenarios["Increased % of Change"].Hidden = true;

                //Enable worksheet protection
                worksheet.Protect("Scenario");

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("HideScenario.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("HideScenario.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();

            }
        }
    }
}