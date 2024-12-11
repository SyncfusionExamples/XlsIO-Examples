# How to apply different border settings in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply different border settings in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	#region Border Settings
	//Apply borders
	worksheet.Range["A2"].CellStyle.Borders.LineStyle = ExcelLineStyle.Medium;
	worksheet.Range["A4"].CellStyle.Borders.LineStyle = ExcelLineStyle.Double;
	worksheet.Range["A6"].CellStyle.Borders.LineStyle = ExcelLineStyle.Dash_dot;
	worksheet.Range["A8"].CellStyle.Borders.LineStyle = ExcelLineStyle.Thick;
	worksheet.Range["C2"].CellStyle.Borders.LineStyle = ExcelLineStyle.Slanted_dash_dot;
	worksheet.Range["C4"].CellStyle.Borders.LineStyle = ExcelLineStyle.Hair;
	worksheet.Range["C6"].CellStyle.Borders.LineStyle = ExcelLineStyle.Medium_dash_dot_dot;
	worksheet.Range["C8"].CellStyle.Borders.LineStyle = ExcelLineStyle.Thin;

	//Apply Border using Border Index
	//Top Border
	worksheet.Range["E2"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Medium;
	//Left Border
	worksheet.Range["E4"].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Double;
	//Bottom Border
	worksheet.Range["E6"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Dashed;
	//Right Border
	worksheet.Range["E8"].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thick;
	//DiagonalUp Border
	worksheet.Range["E10"].CellStyle.Borders[ExcelBordersIndex.DiagonalUp].LineStyle = ExcelLineStyle.Thin;
	//DiagonalDown Border
	worksheet.Range["E12"].CellStyle.Borders[ExcelBordersIndex.DiagonalDown].LineStyle = ExcelLineStyle.Dotted;

	//Apply border color
	worksheet.Range["A2"].CellStyle.Borders.Color = ExcelKnownColors.Blue;

	//Setting the Border as Range
	worksheet.Range["G2:I8"].BorderAround();
	worksheet.Range["G2:I8"].BorderInside(ExcelLineStyle.Dash_dot, ExcelKnownColors.Red);
	#endregion

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/BorderSettings.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}