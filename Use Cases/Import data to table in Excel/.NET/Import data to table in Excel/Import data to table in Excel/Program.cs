using Syncfusion.XlsIO;

class Program
{
    public static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Sample data
            object[,] data = new object[,]
            {
                { "ID", "Name", "Category", "Price" },
                { 1, "Apple", "Fruit", 0.99 },
                { 2, "Carrot", "Vegetable", 0.49 },
                { 3, "Milk", "Dairy", 1.49 }
            };

            //Import data to worksheet
            worksheet.ImportArray(data, 1, 1);

            //Calculate range from data size
            int rowCount = data.GetLength(0);
            int colCount = data.GetLength(1);

            IRange dataRange = worksheet.Range[1, 1, rowCount, colCount];

            //Create a table (ListObject)
            IListObject table = worksheet.ListObjects.Create("SalesTable", dataRange);

            //Apply built-in table style
            table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;

            //Auto-fit columns
            worksheet.UsedRange.AutofitColumns();

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
        }
    }
}