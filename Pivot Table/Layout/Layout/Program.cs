using Syncfusion.XlsIO;

namespace Layout
{
    class Prgoram
    {
        public static void Main(String[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                FileStream fileStream = new FileStream("../../../Data/PivotTable.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[1];

                IPivotTable pivotTable = worksheet.PivotTables[0];
                //Layout the pivot table.
                pivotTable.Layout();

                string fileName = "PivotTable_Layout.xlsx";
                //Saving the workbook as stream
                FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotTable_Layout.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}