using System.IO;
using Syncfusion.XlsIO;

namespace Create_Slicer
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

                //Access the table.
                IListObject table = sheet.ListObjects[0];

                //Add slicer for the table.
                sheet.Slicers.Add(table, 3, 11, 2);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateSlicer.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CreateSlicer.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
