﻿using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Edit_Macro
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xls", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Accessing Vba project
                IVbaProject project = workbook.VbaProject;

                //Accessing vba modules collection
                IVbaModules vbaModules = project.Modules;

                //Access a Vba Module
                IVbaModule vbaModule = vbaModules["Module1"];

                //Edit the macro
                vbaModule.Name = "Module1";

                vbaModule.Code = "Sub Auto_Open()\n MsgBox \"Macro is edited\" \n End Sub ";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("EditMacro.xlsm", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacro);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("EditMacro.xlsm")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
