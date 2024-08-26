using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;
namespace Save_as_Stream
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Creating Vba project
                IVbaProject project = workbook.VbaProject;

                //Accessing vba modules collection
                IVbaModules vbaModules = project.Modules;

                //Adding a vba module
                IVbaModule vbaModule = vbaModules.Add("StdModule", VbaModuleType.StdModule);

                //Adding vba code to the module
                vbaModule.Code = "Sub Auto_Open\n MsgBox \"Macro Added\" \n End Sub";

                #region Save
                //Saving the workbook Macro in XLTM format
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/SaveAsStream.xltm"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacroTemplate);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("SaveAsStream.xltm")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
