using Newtonsoft.Json;
using Syncfusion.XlsIO;
using System.Data;

using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
    IWorksheet worksheet = workbook.Worksheets[0];

    //Extract worksheet data
    DataTable sheetData = worksheet.ExportDataTable( worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.ComputedFormulaValues);

    //Save the data as JSON file
    string json = JsonConvert.SerializeObject(sheetData, Formatting.Indented);

    File.WriteAllText(Path.GetFullPath(@"Output/Output.json"), json);
}