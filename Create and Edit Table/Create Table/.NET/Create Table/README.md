# Create a table in the worksheet using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **create a table in the worksheet** using C#.

## Steps to create a table in the worksheet programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to create a table in the worksheet.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];
	
	//Create for the given data
	IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C5"]);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/CreateTable.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	inputStream.Dispose();
	outputStream.Dispose();
}
```

More information about creating a table in the worksheet can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/excel-library/net/working-with-excel-tables#creating-a-table) section.