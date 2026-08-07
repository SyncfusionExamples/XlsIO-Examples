# Edit Excel document using C#

The [.NET Excel Library](https://www.syncfusion.com/document-sdk/net-excel-library) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. 

## Steps to edit an Excel document programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
```
Step 5: Include the below code snippet in **Program.cs** to edit an Excel document.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;

    IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
    IWorksheet sheet = workbook.Worksheets[0];

    sheet.Range["D1"].Text = "In Stock";
    sheet.Range["E1"].Text = "Total Price";
    sheet.Range["F1"].Text = "Discount (10%)";
    sheet.Range["G1"].Text = "Final Price";

    sheet.Range["D2"].Boolean = true;   
    sheet.Range["D3"].Boolean = true;   

    sheet.Range["A4"].Text = "Bag";
    sheet.Range["B4"].Number = 500;
    sheet.Range["C4"].Number = 5;
    sheet.Range["D4"].Boolean = false;

    sheet.Range["A5"].Text = "Bottle";
    sheet.Range["B5"].Number = 100;
    sheet.Range["C5"].Number = 10;
    sheet.Range["D5"].Boolean = true;

    sheet.Range["E2"].Formula = "B2*C2";   
    sheet.Range["F2"].Formula = "E2*0.1";  
    sheet.Range["G2"].Formula = "E2-F2";   

    sheet.Range["E3"].Formula = "B3*C3";   
    sheet.Range["F3"].Formula = "E3*0.1";  
    sheet.Range["G3"].Formula = "E3-F3";   

    sheet.Range["E4"].Formula = "B4*C4";   
    sheet.Range["F4"].Formula = "E4*0.1";  
    sheet.Range["G4"].Formula = "E4-F4";  

    sheet.Range["E5"].Formula = "B5*C5";  
    sheet.Range["F5"].Formula = "E5*0.1";  
    sheet.Range["G5"].Formula = "E5-F5";  

    sheet.Range["A6"].Text = "Totals";
    sheet.Range["E6"].Formula = "SUM(E2:E5)";
    sheet.Range["F6"].Formula = "SUM(F2:F5)";
    sheet.Range["G6"].Formula = "SUM(G2:G5)";

    sheet.Range["A1:G1"].CellStyle.Font.Bold = true;
    sheet.Range["A1:G6"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

    IListObject table = sheet.ListObjects.Create("SalesTable", sheet.Range["A1:G6"]);
    table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium23;

    workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
    workbook.Close();
}
```
