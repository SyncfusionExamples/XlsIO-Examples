# How to add Oval shape to Excel chart using XlsIO?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to get the RGB values of the cell color using XlsIO.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Apply cell color
    worksheet.Range["A1"].CellStyle.ColorIndex = ExcelKnownColors.Custom50;

    //Get the RGB values of the cell color
    Color color = worksheet.Range["A1"].CellStyle.Color;
    byte red = color.R;
    byte green = color.G;
    byte blue = color.B;

    //Print the RGB values
    Console.WriteLine($"Red: {red}, Green: {green}, Blue: {blue}");

    //Save the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);

    //Dispose stream
    outputStream.Dispose();
}
```

