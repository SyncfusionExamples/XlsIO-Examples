# How to set and format time values in Excel using TimeSpan?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to set and format time values in Excel using TimeSpan.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet sheet = workbook.Worksheets[0];

    TimeSpan ts = new TimeSpan(12, 32, 38);

    //Convert the TimeSpan to a fractional day value that Excel understands
    double excelTimeValue = ts.TotalDays;

    //Set value in cell
    sheet.SetValueRowCol(excelTimeValue, 1, 1);

    //Apply the time format to the cell to display it as 'hh:mm:ss'
    sheet.Range[1, 1].NumberFormat = "hh:mm:ss";

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
}
```
