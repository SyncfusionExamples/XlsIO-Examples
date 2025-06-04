# How to autofill a series in a worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs**  to autofill a series in a worksheet.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Setting the values to the cells
    worksheet["A1"].Number = 1;
    worksheet["A2"].Number = 3;
    worksheet["A3"].Number = 5;

    //Define the source range
    IRange source = worksheet["A1:A3"];

    //Define the destination range
    IRange destinationRange = worksheet["A4:A10"];

    //Use AutoFill method to fill the values based on ExcelAutoFillType
    source.AutoFill(destinationRange, ExcelAutoFillType.FillSeries);

    //Saving the workbook as stream
    FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);

    //Dispose streams
    outputStream.Dispose();
}
```