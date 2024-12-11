# How to apply different font settings in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply different font settings in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Adding text for a range
	worksheet.Range["A1:B6"].Text = "Hello World";

	#region Font Settings
	//Setting Font Type
	worksheet.Range["A1"].CellStyle.Font.FontName = "Arial Black";
	worksheet.Range["A3"].CellStyle.Font.FontName = "Castellar";

	//Setting Font Styles
	worksheet.Range["A2"].CellStyle.Font.Bold = true;
	worksheet.Range["A4"].CellStyle.Font.Italic = true;

	//Setting Font Size
	worksheet.Range["A5"].CellStyle.Font.Size = 18;

	//Setting Font Effects
	worksheet.Range["A6"].CellStyle.Font.Strikethrough = true;
	worksheet.Range["B3"].CellStyle.Font.Subscript = true;
	worksheet.Range["B5"].CellStyle.Font.Superscript = true;

	//Setting UnderLine Types
	worksheet.Range["B1"].CellStyle.Font.Underline = ExcelUnderline.Double;
	worksheet.Range["B2"].CellStyle.Font.Underline = ExcelUnderline.Single;
	worksheet.Range["B4"].CellStyle.Font.Underline = ExcelUnderline.DoubleAccounting;
	worksheet.Range["B6"].CellStyle.Font.Underline = ExcelUnderline.SingleAccounting;

	//Setting Font Color
	worksheet.Range["B6"].CellStyle.Font.Color = ExcelKnownColors.Green;
	#endregion

	worksheet.UsedRange.AutofitColumns();
	worksheet.UsedRange.AutofitRows();

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/FontSettings.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}