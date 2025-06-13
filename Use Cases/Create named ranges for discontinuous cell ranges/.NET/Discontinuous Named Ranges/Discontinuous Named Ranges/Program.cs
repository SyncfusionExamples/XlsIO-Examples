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
            IWorksheet sheet = workbook.Worksheets[0];

            //Create a range collection for discontinuous cells
            IRanges ranges = sheet.CreateRangesCollection();

            //Add different ranges to the collection
            ranges.Add(sheet["D2:D3"]);
            ranges.Add(sheet["D10:D11"]);

            //Create a named range with the collection
            workbook.Names.Add("test", ranges);

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