# How to add Barcode in Excel document using C#?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to add Barcode in Excel document.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    // Load barcodes from local files
    FileStream barcode1 = new FileStream("Data/Barcode1.png", FileMode.Open, FileAccess.Read);
    FileStream barcode2 = new FileStream("Data/Barcode2.png", FileMode.Open, FileAccess.Read);
    worksheet.Pictures.AddPicture(1, 1, barcode1);
    worksheet.Pictures.AddPicture(15, 1, barcode2);
    worksheet.Pictures.AddPicture(1, 10, barcode1);
    worksheet.Pictures.AddPicture(15, 10, barcode2);

    // Save to file system
    FileStream stream = new FileStream("Output/Output.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
    workbook.SaveAs(stream);
    workbook.Close();
    excelEngine.Dispose();
}
```

