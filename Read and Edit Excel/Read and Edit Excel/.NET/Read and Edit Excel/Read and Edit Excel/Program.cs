using Syncfusion.XlsIO;
using System.IO;

namespace Read_and_Edit_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance for ExcelEngine
            ExcelEngine excelEngine = new ExcelEngine();

            #region Open
            //Loads or open an existing workbook through Open method of IWorkbook
            FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(inputStream);
            #endregion

            //Set the version of the workbook
            workbook.Version = ExcelVersion.Xlsx;

            #region Edit
            //Set a value in Excel cell
            workbook.Worksheets[0].Range["A2"].Value = "Hello World";
            #endregion

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/ReadandEditExcel.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);            
            #endregion

            #region Close
            //Close the instance of IWorkbook
            workbook.Close();
            #endregion

            //Dispose streams
            outputStream.Dispose();
            inputStream.Dispose();

            //Dispose the instance of ExcelEngine
            excelEngine.Dispose();

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo("ReadandEditExcel.xlsx")
            {
                UseShellExecute = true
            };
            process.Start();
        }
    }
}
