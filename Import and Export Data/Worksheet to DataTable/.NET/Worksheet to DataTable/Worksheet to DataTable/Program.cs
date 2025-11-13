using System.IO;
using Syncfusion.XlsIO;
using System.Data;

namespace Worksheet_to_DataTable
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Read all data from the used range of worksheet and Export to the DataTable
                DataTable customersTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.ComputedFormulaValues);

                IRange range = worksheet.Range["A1:B10"];

                //Read specified range from the worksheet and Export to DataTable
                DataTable customersTable1 = worksheet.ExportDataTable(range, ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.ComputedFormulaValues);
            }
        }
    }
}





