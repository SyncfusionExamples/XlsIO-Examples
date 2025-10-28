using System.IO;
using Syncfusion.XlsIO;

namespace Dynamic_Refresh
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet pivotSheet = workbook.Worksheets[0];

                //Change the range values that the Pivot Tables range refers to
                workbook.Names["PivotRange"].RefersToRange = pivotSheet.Range["A1:H25"];

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PivotTable.xlsx"));
                #endregion
            }
        }
    }
}





