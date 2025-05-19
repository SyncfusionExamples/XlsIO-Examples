# How to set the background color for Excel Chart in C#?

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

Step 5: Include the below code snippet in **Program.cs** to set the background color for Excel Chart in C#.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
    IWorkbook workbook = application.Workbooks.Open(inputStream);
    IWorksheet worksheet = workbook.Worksheets[0];

    //Get the first chart in the worksheet
    IChartShape chart = worksheet.Charts[0];

    //Applying background color for plot area
    chart.PlotArea.Fill.ForeColor = Color.LightYellow;

    //Applying background color for chart area
    chart.ChartArea.Fill.ForeColor = Color.LightGreen;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams
    outputStream.Dispose();
    inputStream.Dispose();
}
```