# Apply conditional formatting in the worksheet using C#

The Syncfusion&reg; [.NET Excel Library](https://www.syncfusion.com/document-processing/excel-framework/net/excel-library) (XlsIO) enables you to create, read, and edit Excel documents programmatically without Microsoft Excel or interop dependencies. Using this library, you can **apply conditional formatting in the worksheet** using C#.

## Steps to apply conditional formatting in the worksheet programmatically

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
```csharp
using System.IO;
using Syncfusion.XlsIO;
```

Step 5: Include the below code snippet in **Program.cs** to apply conditional formatting in the worksheet.
```csharp
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Applying conditional formatting to "A1"
	IConditionalFormats condition = worksheet.Range["A1"].ConditionalFormats;
	IConditionalFormat condition1 = condition.AddCondition();

	//Represents conditional format rule that the value in target range should be between 10 and 20
	condition1.FormatType = ExcelCFType.CellValue;
	condition1.Operator = ExcelComparisonOperator.Between;
	condition1.FirstFormula = "10";
	condition1.SecondFormula = "20";
	worksheet.Range["A1"].Text = "Enter a number between 10 and 20";

	//Setting back color and font style to be applied for target range
	condition1.BackColor = ExcelKnownColors.Light_orange;
	condition1.IsBold = true;
	condition1.IsItalic = true;

	//Applying conditional formatting to "A3"
	condition = worksheet.Range["A3"].ConditionalFormats;
	IConditionalFormat condition2 = condition.AddCondition();

	//Represents conditional format rule that the cell value should be 1000
	condition2.FormatType = ExcelCFType.CellValue;
	condition2.Operator = ExcelComparisonOperator.Equal;
	condition2.FirstFormula = "1000";
	worksheet.Range["A3"].Text = "Enter the Number as 1000";

	//Setting fill pattern and back color to target range
	condition2.FillPattern = ExcelPattern.LightUpwardDiagonal;
	condition2.BackColor = ExcelKnownColors.Yellow;

	//Applying conditional formatting to "A5"
	condition = worksheet.Range["A5"].ConditionalFormats;
	IConditionalFormat condition3 = condition.AddCondition();

	//Setting conditional format rule that the cell value for target range should be less than or equal to 1000
	condition3.FormatType = ExcelCFType.CellValue;
	condition3.Operator = ExcelComparisonOperator.LessOrEqual;
	condition3.FirstFormula = "1000";
	worksheet.Range["A5"].Text = "Enter a Number which is less than or equal to 1000";

	//Setting back color to target range
	condition3.BackColor = ExcelKnownColors.Light_green;

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/ConditionalFormat.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
```

More information about applying conditional formatting in the worksheet can be found in this [documentation](https://help.syncfusion.com/document-processing/excel/excel-library/net/working-with-conditional-formatting#create-a-conditional-format) section.