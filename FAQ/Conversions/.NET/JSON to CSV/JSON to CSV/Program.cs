using System.Data;
using Newtonsoft.Json;
using Syncfusion.XlsIO;

class Program
{
    static void Main()
    {
        //Load JSON file
        string jsonPath = Path.GetFullPath("Data/Input.json");
        string jsonData = File.ReadAllText(jsonPath);

        //Deserialize to DataTable
        DataTable dataTable = JsonConvert.DeserializeObject<DataTable>(jsonData);

        //Initialize ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            //Import DataTable to worksheet
            sheet.ImportDataTable(dataTable, true, 1, 1);

            //Saving the workbook as CSV
            workbook.SaveAs(Path.GetFullPath("Output/Sample.csv"), ",");
        }
    }
}
