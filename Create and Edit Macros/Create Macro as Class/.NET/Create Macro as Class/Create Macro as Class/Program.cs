using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Create_Macro_as_Class
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

                //Adding a class module
                IVbaModule clsModule = vbaModules.Add("Test", VbaModuleType.ClassModule);
                clsModule.Code = "Public Sub Create()\n MsgBox \"Created a class module\" \n End Sub";

                //Adding a vba module
                IVbaModule vbaModule = vbaModules.Add("Module1", VbaModuleType.StdModule);

                //Using class in StdModule
                vbaModule.Code = "Sub Auto_Open()\n Dim obj As New test \n obj.Create \n End Sub";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("MacroAsClass.xlsm", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacro);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MacroAsClass.xlsm")
                {
                    UseShellExecute = true
                };
                process.Start();
            }

        }
    }
}
