using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Create_Macro_as_MSForm
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

                //Opening form module existing workbook
                FileStream input = new FileStream("../../../Data/InputTemplate.xls", FileMode.Open, FileAccess.ReadWrite);
                IWorkbook newBook = application.Workbooks.Open(input);

                IVbaProject newProject = newBook.VbaProject;

                //Accessing existing form module
                IVbaModule form = newProject.Modules["UserForm1"];

                //Adding a form module in new workbook
                IVbaModule formModule = project.Modules.Add(form.Name, VbaModuleType.MsForm);

                //Copying the form code behind
                formModule.Code = form.Code;

                //Copying the designer of the form
                formModule.DesignerStorage = form.DesignerStorage;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("MacroAsMSForm.xlsm", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacro);
                #endregion

                //Dispose streams
                input.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MacroAsMSForm.xlsm")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
