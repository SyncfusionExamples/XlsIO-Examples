# How to fill a date series in a worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.

```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs**  to fill a date series in a worksheet.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Setting the values to the cells
	worksheet["A1"].DateTime = new DateTime(2025, 1, 1);

	//Define the range
	IRange range = worksheet["A1:A10"];

	//Use FillSeries method to fill the values based on ExcelFillSeries
	range.FillSeries(ExcelSeriesBy.Columns, ExcelFillSeries.Years, 3, new DateTime(2060, 1, 1));

	//Saving the workbook 
	FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);

	//Dispose streams
	outputStream.Dispose();
}
```