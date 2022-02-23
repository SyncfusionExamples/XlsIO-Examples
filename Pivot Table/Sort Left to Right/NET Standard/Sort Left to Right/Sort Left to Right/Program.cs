using System.IO;
using Syncfusion.XlsIO;

namespace Sort_Left_to_Right
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[1];
                IPivotTable pivotTable = sheet.PivotTables[0];

                // Pivot table Left to Right sorting.
                IPivotField columnField = pivotTable.ColumnFields[0];
                columnField.AutoSort(PivotFieldSortType.Descending, 1);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PivotSort.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotSort.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
