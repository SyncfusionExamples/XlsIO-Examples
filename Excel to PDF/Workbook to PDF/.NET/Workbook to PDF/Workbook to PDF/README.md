# Convert an Excel workbook to PDF using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **convert an Excel workbook to PDF** using C#.

## Steps to convert an Excel workbook to PDF programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIORenderer.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
```

Step 5: Include the below code snippet in **Program.cs** to convert an Excel workbook to PDF.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);

	//Initialize XlsIO renderer.
	XlsIORenderer renderer = new XlsIORenderer();

	//Convert Excel document into PDF document 
	PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/WorkbookToPDF.pdf"), FileMode.Create, FileAccess.Write);
	pdfDocument.Save(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
```

More information about converting an Excel workbook to PDF can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/conversions/excel-to-pdf/net/excel-to-pdf-conversion#workbook-to-pdf) section.