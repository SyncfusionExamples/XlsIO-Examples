# How to apply a global style in the worksheet?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to apply a global style in the worksheet.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Adding values to a worksheet range
	worksheet.Range["A1"].Text = "CustomerID";
	worksheet.Range["B1"].Text = "CompanyName";
	worksheet.Range["C1"].Text = "ContactName";
	worksheet.Range["D1"].Text = "TotalSales (in USD)";
	worksheet.Range["A2"].Text = "ALFKI";
	worksheet.Range["A3"].Text = "ANATR";
	worksheet.Range["A4"].Text = "BONAP";
	worksheet.Range["A5"].Text = "BSBEV";
	worksheet.Range["B2"].Text = "Alfred Futterkiste";
	worksheet.Range["B3"].Text = "Ana Trujillo Emparedados y helados";
	worksheet.Range["B4"].Text = "Bon App";
	worksheet.Range["B5"].Text = "B's Beverages";
	worksheet.Range["C2"].Text = "Maria Anders";
	worksheet.Range["C3"].Text = "Ana Trujillo";
	worksheet.Range["C4"].Text = "Laurence Lebihan";
	worksheet.Range["C5"].Text = "Victoria Ashworth";
	worksheet.Range["D2"].Number = 15000.107;
	worksheet.Range["D3"].Number = 27000.208;
	worksheet.Range["D4"].Number = 18700.256;
	worksheet.Range["D5"].Number = 25000.450;

	#region Global Style
	//Formatting
	//Global styles should be used when the same style needs to be applied to more than one cell. This usage of a global style reduces memory usage.
	//Add custom colors to the palette
	workbook.SetPaletteColor(8, Color.FromArgb(255, 174, 33));

	//Defining header style
	IStyle headerStyle = workbook.Styles.Add("HeaderStyle");
	headerStyle.BeginUpdate();
	headerStyle.Color = Color.FromArgb(255, 174, 33);
	headerStyle.Font.Bold = true;
	headerStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
	headerStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
	headerStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
	headerStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
	headerStyle.EndUpdate();

	//Add custom colors to the palette
	workbook.SetPaletteColor(9, Color.FromArgb(239, 243, 247));

	//Defining body style
	IStyle bodyStyle = workbook.Styles.Add("BodyStyle");
	bodyStyle.BeginUpdate();
	bodyStyle.Color = Color.FromArgb(239, 243, 247);
	bodyStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
	bodyStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
	bodyStyle.EndUpdate();

	//Defining number format style
	IStyle numberformatStyle = workbook.Styles.Add("NumberFormatStyle");
	numberformatStyle.BeginUpdate();
	numberformatStyle.NumberFormat = "0.00";
	numberformatStyle.EndUpdate();

	//Apply Header style
	worksheet.Rows[0].CellStyle = headerStyle;
	//Apply Body Style
	worksheet.Range["A2:C5"].CellStyle = bodyStyle;
	//Apply Number Format style
	worksheet.Range["D2:D5"].CellStyle = numberformatStyle;
	#endregion

	//Auto-fit the columns
	worksheet.UsedRange.AutofitColumns();

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/GlobalStyle.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}