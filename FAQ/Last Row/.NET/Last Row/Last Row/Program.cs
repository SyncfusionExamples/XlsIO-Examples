using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Implementation.Collections;

class Program
{
    static void Main(string[] args)
    {
        // Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(1);

            IWorksheet sheet = workbook.Worksheets[0];
            sheet["A1:B10"].Text = "10";
            sheet["C1:C5"].Text = "20";

            int lastRow = GetLastRow(3, sheet as WorksheetImpl);
            Console.WriteLine("Last Row in Column C: " + lastRow);

            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
        }
    }

    private static int GetLastRow(int column, WorksheetImpl worksheet)
    {
        int firstRow = worksheet.UsedRange.Row;
        int lastRow = worksheet.UsedRange.LastRow;
        for (int iRow = lastRow; iRow >= firstRow; iRow--)
        {
            RowStorage rowStorage = WorksheetHelper.GetOrCreateRow(worksheet, iRow - 1, false);
            if (rowStorage != null)
            {
                RowStorageEnumerator enumerator = rowStorage.GetEnumerator(worksheet.RecordExtractor) as RowStorageEnumerator;
                while (enumerator.MoveNext())
                {
                    if (enumerator.ColumnIndex + 1 == column)
                        return iRow;
                }
            }
        }
        return -1;
    }
}
