# How to set column width for a pivot table range in an Excel Document?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;
```

Step 5: Include the below code snippet in **Program.cs** to set column width for a pivot table range in an Excel Document.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];
    IWorksheet worksheet1 = workbook.Worksheets[1];

    //Create pivot cache with the given data range
    IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H5"]);

    //Create pivot table with the cache at the specified range
    IPivotTable pivotTable = worksheet1.PivotTables.Add("PivotTable1", worksheet1["A1"], cache);

    PivotTableImpl pivotTableImpl = pivotTable as PivotTableImpl;

    //Add Pivot table fields 
    pivotTable.Fields[0].Axis = PivotAxisTypes.Row;
    pivotTable.Fields[1].Axis = PivotAxisTypes.Row;
    pivotTable.DataFields.Add(pivotTable.Fields["Total"], "Sum", PivotSubtotalTypes.Sum);

    //Set column width
    worksheet1.Range["A10"].ColumnWidth = 50;

    //Disable pivot table autoformat    
    (pivotTable.Options as PivotTableOptions).IsAutoFormat = false;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
```	