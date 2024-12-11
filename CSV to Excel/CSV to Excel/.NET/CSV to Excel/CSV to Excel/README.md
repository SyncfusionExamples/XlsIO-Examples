# Convert a CSV to an Excel document using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **convert a CSV to an Excel document** using C#.

## Steps to convert a CSV to an Excel document programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to convert a CSV to an Excel document.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.csv"), FileMode.Open, FileAccess.Read);

	//Open the CSV file
	IWorkbook workbook = application.Workbooks.Open(inputStream, ",");
	IWorksheet worksheet = workbook.Worksheets[0];

	//Saving the workbook as stream
	FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
```

More information about converting a CSV to an Excel document can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/conversions/csv-to-excel/net/csv-to-excel-conversion) section.