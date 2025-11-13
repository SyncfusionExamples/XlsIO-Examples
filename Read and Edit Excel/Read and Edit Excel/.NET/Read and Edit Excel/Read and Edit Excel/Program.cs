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
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(@"Data/InputTemplate.xlsx");
            #endregion

            //Set the version of the workbook
            workbook.Version = ExcelVersion.Xlsx;

            #region Edit
            //Set a value in Excel cell
            workbook.Worksheets[0].Range["A2"].Value = "Hello World";
            #endregion

            #region Save
            //Saving the workbook
            workbook.SaveAs("Output/ReadandEditExcel.xlsx");            
            #endregion

            #region Close
            //Close the instance of IWorkbook
            workbook.Close();
            #endregion

            //Dispose the instance of ExcelEngine
            excelEngine.Dispose();
        }
    }
}





