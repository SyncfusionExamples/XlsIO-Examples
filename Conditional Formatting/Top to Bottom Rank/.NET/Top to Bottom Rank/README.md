# How to apply conditional formatting to the top and bottom N rank values?

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

Step 5: Include the below code snippet in **Program.cs** to apply conditional formatting to format the top 10 rank values.
{% tabs %}
{% highlight c# tabtitle="C#" %}
using (ExcelEngine excelEngine = new ExcelEngine())
{
	IApplication application = excelEngine.Excel;
	application.DefaultVersion = ExcelVersion.Xlsx;
	FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
	IWorkbook workbook = application.Workbooks.Open(inputStream);
	IWorksheet worksheet = workbook.Worksheets[0];

	//Applying conditional formatting to "N6:N35".
	IConditionalFormats formats = worksheet.Range["N6:N35"].ConditionalFormats;
	IConditionalFormat format = formats.AddCondition();

	//Applying top or bottom rule in the conditional formatting.
	format.FormatType = ExcelCFType.TopBottom;
	ITopBottom topBottom = format.TopBottom;

	//Set type as Top for TopBottom rule.
	topBottom.Type = ExcelCFTopBottomType.Top;

	//Set rank value for the TopBottom rule.
	topBottom.Rank = 10;

	//Set color for Conditional Formattting.
	format.BackColorRGB = Syncfusion.Drawing.Color.FromArgb(51, 153, 102);

	#region Save
	//Saving the workbook
	FileStream outputStream = new FileStream(Path.GetFullPath("Output/TopToBottomRank.xlsx"), FileMode.Create, FileAccess.Write);
	workbook.SaveAs(outputStream);
	#endregion

	//Dispose streams
	outputStream.Dispose();
	inputStream.Dispose();
}
{% endhighlight %}
{% endtabs %}