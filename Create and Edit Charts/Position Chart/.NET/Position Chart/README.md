# Positioning a Chart

Step 1: Create a New C# Console Application Project.

Step 2: Name the Project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to position a chart in Excel.

```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;

    // Create a new workbook
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];

    // Fill sample data
    worksheet.Range["A1"].Text = "Category";
    worksheet.Range["B1"].Text = "Value";
    worksheet.Range["A2"].Text = "A";
    worksheet.Range["A3"].Text = "B";
    worksheet.Range["A4"].Text = "C";
    worksheet.Range["B2"].Number = 10;
    worksheet.Range["B3"].Number = 20;
    worksheet.Range["B4"].Number = 30;

    // Add a chart
    IChartShape chart = worksheet.Charts.Add();
    chart.DataRange = worksheet.Range["A1:B4"];
    chart.ChartType = ExcelChartType.Column_Clustered;

    //Set chart position 
    chart.Top = 100;     
    chart.Left = 150;

    //Set height and width
    IChart chart1 = worksheet.Charts[0];
    chart1.Height = 300;  
    chart1.Width = 500;   

    #region Save
    //Saving the workbook
    FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    #endregion

    //Dispose streams   
    outputStream.Dispose();
}
```			