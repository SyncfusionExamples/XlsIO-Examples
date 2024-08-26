using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;
namespace Clear_All_Macros
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xls"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Accessing Vba project
                IVbaProject project = workbook.VbaProject;

                //Accessing vba modules collection
                IVbaModules vbaModules = project.Modules;

                //Remove all macros
                vbaModules.Clear();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ClearAllMacro.xlsm"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsMacro);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}

