# How to convert JSON document to CSV format document?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) and [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json) NuGet packages as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System.Data;
using Newtonsoft.Json;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to convert JSON document to CSV format document.
```csharp
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
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/Sample.csv"), FileMode.Create, FileAccess.ReadWrite);
    workbook.SaveAs(outputStream, ",");

    //Dispose streams
    outputStream.Dispose();
}
```		