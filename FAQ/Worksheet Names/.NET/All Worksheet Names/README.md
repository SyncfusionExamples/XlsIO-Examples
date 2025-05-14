# Retrieve All Worksheet Names

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
```

Step 5: Include the below code snippet in **Program.cs** to retrieve all worksheet names.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;

    FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);

    //Get the worksheets collection
    WorksheetsCollection worksheets = workbook.Worksheets as WorksheetsCollection;

    //Print all worksheet names
    foreach (IWorksheet worksheet in worksheets)
    {
        Console.WriteLine(worksheet.Name);
    }

    //Dispose streams
    inputStream.Dispose();
}
```