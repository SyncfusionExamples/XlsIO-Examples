# How to show the leader line on Excel chart using XlsIO?

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to show the leader line on Excel chart.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet sheet = workbook.Worksheets[0];

    //Add data
    sheet.Range["A1"].Text = "Fruit";
    sheet.Range["B1"].Text = "Quantity";
    sheet.Range["A2"].Text = "Apple";
    sheet.Range["A3"].Text = "Banana";
    sheet.Range["A4"].Text = "Cherry";
    sheet.Range["B2"].Number = 40;
    sheet.Range["B3"].Number = 30;
    sheet.Range["B4"].Number = 30;

    //Add a Pie chart 
    IChart chart = sheet.Charts.Add();
    chart.ChartType = ExcelChartType.Pie;
    chart.DataRange = sheet.Range["A1:B4"];
    chart.IsSeriesInRows = false;
    chart.ChartTitle = "Fruit Distribution";

    //Enable data labels with values, and leader lines
    IChartSerie series = chart.Series[0];
    series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
    series.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;

    //Manually resizing data label area using Manual Layout
    chart.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Left = 0.09;
    chart.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Top = 0.01;

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams   
    outputStream.Dispose();
}
```			