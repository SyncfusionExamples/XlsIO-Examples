# How to set the default font and font size in an Excel Workbook?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to set the default font and font size in an Excel Workbook.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet sheet = workbook.Worksheets[0];

    //Set default font and font size
    workbook.StandardFont = "Calibri";
    workbook.StandardFontSize = 12;

    //Add some text
    sheet.Range["A1"].Text = "This is default font and size";

    //Save to file
    FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);

    //Dispose streams
    outputStream.Dispose();

}
```