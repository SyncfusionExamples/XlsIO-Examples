using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Shapes_with_Macro
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
                vbaModule.Code = "Sub Invoke()\n MsgBox \"Macro Added\" \n End Sub";

                //Adding a auto shape
                IShape shape = sheet.Shapes.AddAutoShapes(AutoShapeType.Rectangle, 1, 2, 60, 70);
                shape.Name = "Shape1";

                //Assigning a Macro to shape
                shape.OnAction = "StdModule.Invoke";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ShapesWithMacro.xlsm", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacro);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ShapesWithMacro.xlsm")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
