# Convert an Excel document to CSV using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **convert an Excel document to CSV** using C#.

## Steps to convert an Excel document to CSV programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to convert an Excel document to CSV.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);

	//Saving the workbook as streams
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/Sample.csv"), FileMode.Create, FileAccess.ReadWrite);
	workbook.SaveAs(outputStream, ",");

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
```

More information about converting an Excel document to CSV can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/conversions/excel-to-csv/net/excel-to-csv-conversion) section.