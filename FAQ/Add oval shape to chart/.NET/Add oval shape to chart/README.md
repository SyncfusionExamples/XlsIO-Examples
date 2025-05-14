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

Step 5: Include the below code snippet in **Program.cs** to add Oval shape to Excel chart using XlsIO.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Add chart to worksheet
    IChart chart = worksheet.Charts.Add();

    //Add oval shape to chart
    IShape shape = chart.Shapes.AddAutoShapes(AutoShapeType.Oval, 20, 60, 500, 400);

    //Format the shape
    shape.Line.ForeColorIndex = ExcelKnownColors.Red;

    //Add the text to the oval shape and set the text alignment on the shape
    shape.TextFrame.TextRange.Text = "This is an oval shape";
    shape.TextFrame.VerticalAlignment = ExcelVerticalAlignment.MiddleCentered;
    shape.TextFrame.HorizontalAlignment = ExcelHorizontalAlignment.CenterMiddle;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
}
```

