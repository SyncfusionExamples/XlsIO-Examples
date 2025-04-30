# How to compute the size of the Excel file?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to compute the size of the Excel file.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    worksheet.Range["A1"].Text = "Sample Data";

    //Save to memory stream
    using (MemoryStream stream = new MemoryStream())
    {
        workbook.SaveAs(stream);

        //Compute file size in bytes
        long sizeInBytes = stream.Length;
        Console.WriteLine($"File size: {sizeInBytes} bytes");

        //Convert to KB 
        double sizeInKB = sizeInBytes / 1024.0;
        Console.WriteLine($"File size: {sizeInKB:F2} KB");
    }
}
```