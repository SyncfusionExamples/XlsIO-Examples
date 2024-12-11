# How to highlight cells using conditional formatting?

Step 1: Create a new C# Console Application project.

Step 2: Name the project.

Step 3: Install the [Syncfusion.XlsIO.Net.Core](https://www.nuget.org/packages/Syncfusion.XlsIO.Net.Core) NuGet package as reference to your .NET Standard applications from [NuGet.org](https://www.nuget.org).

Step 4: Include the following namespaces in the **Program.cs** file.
{% tabs %}  
{% highlight c# tabtitle="C#" %}
using System.IO;
using Syncfusion.XlsIO;
{% endhighlight %}
{% endtabs %}  

Step 5: Include the below code snippet in **Program.cs** to highlight cells by formatting unique and duplicate values using conditional formatting.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Excel2016;
	IWorkbook workbook = application.Workbooks.Create(1);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Fill worksheet with data
	worksheet.Range["A1:B1"].Merge();
	worksheet.Range["A1:B1"].CellStyle.Font.RGBColor = Color.FromArgb(255, 102, 102, 255);
	worksheet.Range["A1:B1"].CellStyle.Font.Size = 14;
	worksheet.Range["A1:B1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
	worksheet.Range["A1"].Text = "Global Internet Usage";
	worksheet.Range["A1:B1"].CellStyle.Font.Bold = true;

	worksheet.Range["A3:B21"].CellStyle.Font.RGBColor = Color.FromArgb(255, 64, 64, 64);
	worksheet.Range["A3:B3"].CellStyle.Font.Bold = true;
	worksheet.Range["B3"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;

	worksheet.Range["A3"].Text = "Country";
	worksheet.Range["A4"].Text = "Northern America";
	worksheet.Range["A5"].Text = "Central America";
	worksheet.Range["A6"].Text = "The Caribbean";
	worksheet.Range["A7"].Text = "South America";
	worksheet.Range["A8"].Text = "Northern Europe";
	worksheet.Range["A9"].Text = "Eastern Europe";
	worksheet.Range["A10"].Text = "Western Europe";
	worksheet.Range["A11"].Text = "Southern Europe";
	worksheet.Range["A12"].Text = "Northern Africa";
	worksheet.Range["A13"].Text = "Eastern Africa";
	worksheet.Range["A14"].Text = "Middle Africa";
	worksheet.Range["A15"].Text = "Western Africa";
	worksheet.Range["A16"].Text = "Southern Africa";
	worksheet.Range["A17"].Text = "Central Asia";
	worksheet.Range["A18"].Text = "Eastern Asia";
	worksheet.Range["A19"].Text = "Southern Asia";
	worksheet.Range["A20"].Text = "SouthEast Asia";
	worksheet.Range["A21"].Text = "Oceania";

	worksheet.Range["B3"].Text = "Usage";
	worksheet.Range["B4"].Value = "88%";
	worksheet.Range["B5"].Value = "61%";
	worksheet.Range["B6"].Value = "49%";
	worksheet.Range["B7"].Value = "68%";
	worksheet.Range["B8"].Value = "94%";
	worksheet.Range["B9"].Value = "74%";
	worksheet.Range["B10"].Value = "90%";
	worksheet.Range["B11"].Value = "77%";
	worksheet.Range["B12"].Value = "49%";
	worksheet.Range["B13"].Value = "27%";
	worksheet.Range["B14"].Value = "12%";
	worksheet.Range["B15"].Value = "39%";
	worksheet.Range["B16"].Value = "51%";
	worksheet.Range["B17"].Value = "50%";
	worksheet.Range["B18"].Value = "58%";
	worksheet.Range["B19"].Value = "36%";
	worksheet.Range["B20"].Value = "58%";
	worksheet.Range["B21"].Value = "69%";

	worksheet.SetColumnWidth(1, 23.45);
	worksheet.SetColumnWidth(2, 8.09);

	IConditionalFormats conditionalFormats =
	worksheet.Range["A4:B21"].ConditionalFormats;
	IConditionalFormat condition = conditionalFormats.AddCondition();

	//conditional format to set duplicate format type
	condition.FormatType = ExcelCFType.Duplicate;
	condition.BackColorRGB = Color.FromArgb(255, 255, 199, 206);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/UniqueandDuplicate.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}