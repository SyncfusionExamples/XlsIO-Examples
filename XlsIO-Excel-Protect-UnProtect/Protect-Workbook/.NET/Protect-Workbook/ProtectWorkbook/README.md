# Protect a workbook with a password using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **protect a workbook with a password** using C#. After adding workbook protection, the structural changes of the workbook are disabled.

## Steps to protect a workbook with a password programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```chsharp
using System;
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to protect a workbook with a password.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputWorkbook.xlsx"), FileMode.Open, FileAccess.ReadWrite);

	//Open Excel
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Protect workbook with password
	workbook.Protect(true, true, "syncfusion");
	
	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/ProtectedWorkbook.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
```

More information about protecting a workbook with a password can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/excel-library/net/security#protect-workbook) section.